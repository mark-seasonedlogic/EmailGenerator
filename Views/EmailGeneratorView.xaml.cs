using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using System.Threading.Tasks;
using Windows.Storage.Pickers;
using Windows.Storage;
using OutlookDeviceEmailer;
using CsvHelper;
using System.Globalization;
using CsvHelper.Configuration;
using EmailGenerator.Services;
using Microsoft.Extensions.DependencyInjection;
using static Org.BouncyCastle.Math.EC.ECCurve;
using Microsoft.Extensions.Configuration;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace EmailGenerator.Views
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class EmailGeneratorView : Page
    {
        private readonly EmailDataService _emailDataService;
        private readonly EmailBuilderService _emailBuilderService;
        private string deviceFilePath;
        private string emailFilePath;
        private readonly IConfiguration _config;
        public EmailGeneratorView()
    : this(App.Host.Services.GetRequiredService<IConfiguration>(),
           App.Host.Services.GetRequiredService<EmailDataService>(),
           App.Host.Services.GetRequiredService<EmailBuilderService>())
        {
        }
        public EmailGeneratorView(IConfiguration config, EmailDataService emailDataService, EmailBuilderService emailBuilderService)
        {
            _emailDataService = emailDataService;
            _emailBuilderService = emailBuilderService;
            _config = config;
            this.InitializeComponent();
        }
        private async Task<string> SelectFile(string title)
        {
            var picker = new FileOpenPicker();

            // Required in WinUI 3
            var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(App.MainWindow);
            WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);

            picker.SuggestedStartLocation = PickerLocationId.Desktop;
            picker.FileTypeFilter.Add(".csv");

            StorageFile file = await picker.PickSingleFileAsync();
            return file?.Path ?? string.Empty;
        }

        private async void SelectDeviceFile_Click(object sender, RoutedEventArgs e)
        {
            deviceFilePath = await SelectFile("Select Device List CSV");
            txtDeviceFilePath.Text = string.IsNullOrEmpty(deviceFilePath) ? "No file selected" : deviceFilePath;
        }

        private async void SelectEmailFile_Click(object sender, RoutedEventArgs e)
        {
            emailFilePath = await SelectFile("Select Email List CSV");
            txtEmailFilePath.Text = string.IsNullOrEmpty(emailFilePath) ? "No file selected" : emailFilePath;
        }

        private async Task<ContentDialogResult> ShowConfirmationDialog(string title, string message)
        {
            ContentDialog dialog = new ContentDialog
            {
                Title = title,
                Content = message,
                PrimaryButtonText = "OK",
                CloseButtonText = "Cancel",
                XamlRoot = this.Content.XamlRoot
            };

            return await dialog.ShowAsync();
        }


        #region Refactor This!!
        private async void SendEmails_Click(object sender, RoutedEventArgs e)
        {

            if (string.IsNullOrEmpty(deviceFilePath) || string.IsNullOrEmpty(emailFilePath))
            {
                await ShowMessage("Error", "Please select both CSV files before proceeding.");
                return;
            }
            var fieldMappings = _config.GetSection("FieldMappings").Get<Dictionary<string, string>>();
            var conceptMappings = _config.GetSection("ConceptMappings").Get<Dictionary<string, string>>();

            Func<string, string> resolveField = key =>
                fieldMappings.TryGetValue(key, out var val) ? val : key;

            Func<string, string> mapConcept = code =>
                conceptMappings.TryGetValue(code, out var abbr) ? abbr : "UNKNOWN";



            Dictionary<string, Dictionary<string, string>> emailLookup = _emailDataService.LoadEmailLookup(emailFilePath,resolveField,mapConcept);
            Dictionary<string, List<string>> emailDict = _emailDataService.GenerateEmailDictionary(deviceFilePath, emailLookup,resolveField);

            string generatedEmailDirectory = await Task.Run(() =>
            {
                return _emailBuilderService.CreateEmailFilesFromHtmlAsync(emailDict, resolveField).GetAwaiter().GetResult();
            });
            var result = await ShowConfirmationDialog("Emails Generated", "Preview Emails?");
            if (result == ContentDialogResult.Primary) // User clicked OK
            {
                PreviewWindow previewWindow = new PreviewWindow(generatedEmailDirectory);
                previewWindow.Activate();
            }
        }

        private void EditTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(deviceFilePath) || string.IsNullOrEmpty(emailFilePath))
            {
                ShowMessage("Error", "Please select both CSV files before editing the template.");
                return;
            }

            var deviceHeaders = GetHeadersFromCsv(deviceFilePath);
            var emailHeaders = GetHeadersFromCsv(emailFilePath);

            EmailHTMLEditorWindow editorWindow = new EmailHTMLEditorWindow(deviceHeaders, emailHeaders);
            editorWindow.Activate();

        }

        private static bool isDialogOpen = false;
        private async System.Threading.Tasks.Task ShowMessage(string title, string message)
        {
            if (isDialogOpen)
                return; // Prevent multiple dialogs from opening

            isDialogOpen = true; // Set flag

            ContentDialog dialog = new ContentDialog
            {
                Title = title,
                Content = message,
                CloseButtonText = "OK",
                XamlRoot = this.Content.XamlRoot // Required in WinUI 3
            };

            await dialog.ShowAsync();
            isDialogOpen = false;
        }

        private List<string> GetHeadersFromCsv(string filePath)
        {
            using var reader = new StreamReader(filePath);
            using var csv = new CsvReader(reader, new CsvHelper.Configuration.CsvConfiguration(CultureInfo.InvariantCulture));

            if (csv.Read() && csv.ReadHeader())
            {
                return csv.HeaderRecord.ToList();
            }

            return new List<string>();
        }

  

        #endregion

    }
}
