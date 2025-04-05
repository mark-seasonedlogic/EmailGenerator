using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System.Threading.Tasks;
using Windows.Storage.Pickers;
using Windows.Storage;
using System.Net.Mail;
using System.Text;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI;
using Windows.UI;
using Microsoft.UI.Windowing;
using Windows.Graphics;
using MimeKit;
using OpenMcdf;
using MsgKit;
using Microsoft.Extensions.Configuration;
using MimeKit.Utils;
using System.Text.RegularExpressions;
using EmailGenerator.Interfaces;
using System.Diagnostics;
using EmailGenerator.Models;
using Outlook = Microsoft.Office.Interop.Outlook;
using EmailGenerator.Helpers;

namespace OutlookDeviceEmailer
{
    public partial class MainWindow : Window
    {
        private readonly IConfiguration _config;
        private string deviceFilePath;
        private string emailFilePath;
        private readonly Dictionary<string, string> conceptMappings;
        private readonly IFedExShippingService _shippingService;
        public MainWindow(IFedExShippingService shippingService)
        {
            _shippingService = shippingService;

            _config = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

            conceptMappings = _config.GetSection("ConceptMappings").Get<Dictionary<string, string>>();

            InitializeComponent();
            ApplyAcrylicEffect();
            SetRoundedCorners();
            SetWindowSize(600, 400);
        }
        public MainWindow()
        {
            _config = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

            conceptMappings = _config.GetSection("ConceptMappings").Get<Dictionary<string, string>>();

            InitializeComponent();
            ApplyAcrylicEffect();
            SetRoundedCorners();
            SetWindowSize(600, 400);
        }
        private void SetWindowSize(int width, int height)
        {
            IntPtr hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WindowId myWndId = Win32Interop.GetWindowIdFromWindow(hWnd);
            AppWindow appWindow = AppWindow.GetFromWindowId(myWndId);

            appWindow.Resize(new SizeInt32(width, height));
        }

        private string GetConceptAbbreviationFromCode(string conceptCode)
        {
            return conceptMappings.TryGetValue(conceptCode, out string abbreviation) ? abbreviation : "UNKNOWN";
        }

        private string GetConfiguredField(string key)
        {
            return _config[$"FieldMappings:{key}"] ?? key;
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
        private void SetRoundedCorners()
{
    IntPtr hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
    WindowId myWndId = Win32Interop.GetWindowIdFromWindow(hWnd);
    AppWindow appWindow = AppWindow.GetFromWindowId(myWndId);

    if (appWindow.Presenter is OverlappedPresenter presenter)
    {
        presenter.IsResizable = false;
        presenter.IsMaximizable = false; // Prevent full-screen to keep rounded effect
                
        presenter.SetBorderAndTitleBar(false, false); // Remove default borders
    }

    appWindow.TitleBar.ExtendsContentIntoTitleBar = true; // Extend content into the title bar
    appWindow.SetPresenter(AppWindowPresenterKind.CompactOverlay); // Ensure standard windowing with rounded corners
}
        private void ApplyAcrylicEffect()
        {
            if (RootGrid != null)
            {
                RootGrid.Background = new AcrylicBrush()
                {
                    TintColor = Color.FromArgb(255, 245, 245, 245),
                    TintOpacity = 0.6,
                    FallbackColor = Colors.LightGray
                };
            }
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

        private async Task<string> SelectFile(string title)
        {
            var picker = new FileOpenPicker();

            // Required in WinUI 3
            var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);

            picker.SuggestedStartLocation = PickerLocationId.Desktop;
            picker.FileTypeFilter.Add(".csv");

            StorageFile file = await picker.PickSingleFileAsync();
            return file?.Path ?? string.Empty;
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

        private async void SendEmails_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(deviceFilePath) || string.IsNullOrEmpty(emailFilePath))
            {
                await ShowMessage("Error", "Please select both CSV files before proceeding.");
                return;
            }

            Dictionary<string, Dictionary<string,string>> emailLookup = LoadEmailLookup(emailFilePath);
            Dictionary<string, List<string>> emailDict = GenerateEmailDictionary(deviceFilePath, emailLookup);

            string generatedEmailDirectory = await CreateEmailFilesFromHTML(emailDict); // Now returns directory

            var result = await ShowConfirmationDialog("Emails Generated", "Preview Emails?");
            if (result == ContentDialogResult.Primary) // User clicked OK
            {
                PreviewWindow previewWindow = new PreviewWindow(generatedEmailDirectory);
                previewWindow.Activate();
            }
        }

        private Dictionary<string, Dictionary<string, string>> LoadEmailLookup(string emailFile)
        {
            var emailLookup = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

            using (var reader = new StreamReader(emailFile))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
            {
                if (!csv.Read() || !csv.ReadHeader())
                {
                    throw new Exception("Restaurant directory CSV file is empty or missing headers.");
                }

                var headers = csv.HeaderRecord.ToList();

                while (csv.Read())
                {
                    string conceptCode = csv.GetField(GetConfiguredField("CONCEPT_CD"))?.Trim();
                    string restaurantNumber = csv.GetField(GetConfiguredField("RSTRNT_NBR"))?.Trim();

                    if (string.IsNullOrEmpty(conceptCode) || string.IsNullOrEmpty(restaurantNumber))
                        continue;

                    string conceptAbbreviation = GetConceptAbbreviationFromCode(conceptCode);
                    string paddedRestaurantNumber = restaurantNumber.PadLeft(4, '0');
                    string lookupKey = $"{paddedRestaurantNumber}{conceptAbbreviation}";

                    var restaurantData = new Dictionary<string, string>();
                    foreach (var header in headers)
                    {
                        string value = csv.GetField(header)?.Trim();
                        restaurantData[header] = value;
                    }

                    emailLookup[lookupKey] = restaurantData;
                }
            }

            return emailLookup;
        }

        private Dictionary<string, List<string>> GenerateEmailDictionary(string deviceFile, Dictionary<string, Dictionary<string, string>> emailLookup)
        {
            var emailDict = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            string jvpEmail = "";
            string mvpEmail = "";
            using (var reader = new StreamReader(deviceFile))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
            {
                if (csv.Read())
                {
                    csv.ReadHeader();
                }
                else
                {
                    throw new Exception("Device CSV file is empty.");
                }

                while (csv.Read())
                {
                    string userName = csv.GetField("Username")?.Trim();
                    if (string.IsNullOrEmpty(userName))
                        continue;
                    string storeNumber = userName.Substring(3, 4);
                    string concept = userName.Substring(0, 3);
                    string ip = csv.GetField("MD + Serial")?.Replace(", "," - SN: ").Trim();
                    string serial = csv.GetField("Serial Number")?.Trim();
                    jvpEmail = csv.GetField("JVP EMAIL")?.Trim();
                    mvpEmail = csv.GetField("MVP EMAIL")?.Trim();
                    //RVP_ADMIN_EMAIL
                    //JVP_EMAIL_ADDR
                    //MARKTNG_MGR_NAME


                    string lookupKey = $"{storeNumber}{concept}";

                    // Find matching restaurant based on store number & concept
                    var matchingRestaurants = emailLookup.Keys
                        .Where(k => k.StartsWith(lookupKey, StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    if (matchingRestaurants.Count > 0)
                    {
                        string restaurantKey = matchingRestaurants.First();
                        var restaurantInfo = emailLookup[restaurantKey];

                        // Extract the recipient email separately
                        string recipientEmail = restaurantInfo["STORE_EMAIL_ADDR"];

                        if (!emailDict.ContainsKey(recipientEmail))
                            emailDict[recipientEmail] = new List<string>();

                        // Merge restaurant & device data into a single key-value string
                        string formattedData = string.Join(", ",
                           restaurantInfo.Where(kv => kv.Key != "STORE_EMAIL_ADDR")
                                         .Select(kv => $"{kv.Key}: {kv.Value}")) +
                           $", MD + Serial: {ip}, Serial Number: {serial}, JVP EMAIL: {jvpEmail}, MVP EMAIL: {mvpEmail}";

                        

                        emailDict[recipientEmail].Add(formattedData);
                    }
                }
            }

            return emailDict;
        }
        private Dictionary<string, Dictionary<string, string>> LoadRestaurantData(string emailFile)
        {
            var restaurantData = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

            using (var reader = new StreamReader(emailFile))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
            {
                if (!csv.Read() || !csv.ReadHeader())
                {
                    throw new Exception("Email CSV file is empty or missing headers.");
                }

                while (csv.Read())
                {
                    string restaurantName = csv.GetField(GetConfiguredField("RSTRNT_LEGAL_NAME"))?.Trim();
                    string email = csv.GetField(GetConfiguredField("STORE_EMAIL_ADDR"))?.Trim();
                    string address = csv.GetField(GetConfiguredField("ADDR_LINE1_TXT"))?.Trim();
                    string city = csv.GetField(GetConfiguredField("CITY_NAME"))?.Trim();
                    string state = csv.GetField(GetConfiguredField("STATE_CD"))?.Trim();
                    string phone = csv.GetField(GetConfiguredField("STORE_PHONE_NO"))?.Trim();

                    if (!string.IsNullOrEmpty(restaurantName) && !string.IsNullOrEmpty(email))
                    {
                        restaurantData[restaurantName] = new Dictionary<string, string>
                        {
                            { GetConfiguredField("STORE_EMAIL_ADDR"), email },
                            { "Restaurant Name", restaurantName },
                            { "Address", address },
                            { "City", city },
                            { "State", state },
                            { "Phone", phone }
                        };
                    }
                }
            }

            return restaurantData;
        }
        private string LoadHtmlEmailTemplate()
        {
            string htmlPath = "EmailTemplates/email_template_content.html";
            return File.Exists(htmlPath) ? File.ReadAllText(htmlPath, Encoding.UTF8) : "<p>[No HTML Template Found]</p>";
        }

        private async Task<string>  CreateEmailFilesFromHTML(Dictionary<string, List<string>> emailDict)
        {
            string saveDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Generated Emails (HTML)");
            Directory.CreateDirectory(saveDirectory);

            string emailTemplate = LoadHtmlEmailTemplate(); // Load the HTML template

            foreach (var recipient in emailDict.Keys)
            {
                string formattedEmail = emailTemplate; // Work on a fresh copy per recipient
                StringBuilder deviceDetails = new StringBuilder(); // Store all device entries

                Dictionary<string, string> restaurantDetailsDict = new Dictionary<string, string>();
                Dictionary<string,string> keyValuePairs = new Dictionary<string, string>();

                foreach (var dataString in emailDict[recipient])
                {
                    // Parse key-value pairs from CSV-formatted string
                    keyValuePairs = new Dictionary<string, string>();

                    foreach (var entry in dataString.Split(','))
                    {
                        var pair = entry.Split(':', 2);
                        if (pair.Length == 2)
                        {
                            keyValuePairs[pair[0].Trim()] = pair[1].Trim();
                        }
                    }

                    restaurantDetailsDict = keyValuePairs;

                    // Format device info as list item
                    //Try using new device detail columns:
                    //deviceDetails.AppendLine($"<li>{keyValuePairs["Wi-Fi IP Address"]}, Serial: {keyValuePairs["Serial Number"]}</li>");
                    deviceDetails.AppendLine($"<li>{keyValuePairs["MD + Serial"]}</li>");
                }
                
                // Replace device list placeholder with actual HTML list
                formattedEmail = formattedEmail.Replace("{{Device List}}", $"<ul>{deviceDetails}</ul>");

                // Replace placeholders for restaurant/device fields
                foreach (var key in restaurantDetailsDict.Keys)
                {
                    string placeholder = $"{{{{{key}}}}}"; // e.g., {{Wi-Fi IP Address}}
                    formattedEmail = formattedEmail.Replace(placeholder, restaurantDetailsDict[key]);
                }
                //string supportEmail = ReplacePlaceholders("{{Support Email}}", keyValuePairs);
                string jvpEmail = ReplacePlaceholders("{{JVP EMAIL}}", keyValuePairs);
                string mvpEmail = ReplacePlaceholders("{{MVP EMAIL}}", keyValuePairs);
                //string mpEmail = ReplacePlaceholders("{{MARKTNG_MGR_NAME}}", keyValuePairs);


                string restaurantNumber = recipient.Split('@')[0]; // e.g.,
                restaurantNumber = restaurantNumber.Substring(restaurantNumber.Length - 4);// { emailDict.TryGetValue("RSTRNT_NBR") };
                List<MailAddress> ccRecipientList = new List<MailAddress>();
                List<MailAddress> toRecipientsList = new List<MailAddress>();   
                // Create the MailMessage
                MailMessage mail = new MailMessage();
                mail.To.Add(recipient);
                toRecipientsList.Add(new MailAddress(recipient));
                mail.CC.Add(new MailAddress("KarlyLopez@BloominBrands.com"));
                ccRecipientList.Add(new MailAddress("KarlyLopez@BloominBrands.com"));
                mail.CC.Add(new MailAddress("JoelCapo@BloominBrands.com"));
                ccRecipientList.Add(new MailAddress("JoelCapo@BloominBrands.com"));
                try
                {
                    mail.CC.Add(new MailAddress(jvpEmail));
                    ccRecipientList.Add(new MailAddress(jvpEmail));
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
                try
                {
                    mail.CC.Add(new MailAddress(mvpEmail));
                    ccRecipientList.Add(new MailAddress(mvpEmail));
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                mail.Subject = "Duplicate Android POSi Tablets in your Restaurant";
                mail.Body = formattedEmail;
                mail.IsBodyHtml = true;

                string safeName = recipient.Replace("@", "_at_").Replace(".", "_");
                string emlFileName = Path.Combine(saveDirectory, $"{safeName}.eml");
                string msgFileName = Path.Combine(saveDirectory, $"{safeName}.msg");


                //mail.From = new MailAddress("jugaldez@Bloominbrands.com");
                //Try using Outlook Interop
                OutlookInteropHelper.GenerateMsgWithEmbeddedImagesAndPdf(
                    subject: mail.Subject,
                    htmlInput: mail.Body,
                    imageDirectory: "C:\\Users\\MarkYoung\\source\\repos\\EmailGenerator\\bin\\x64\\Debug\\net8.0-windows10.0.19041.0\\win-x64\\EmailTemplates\\Untitled_files\\",
                    msgOutputPath: msgFileName,
                    pdfFolderPath: "C:\\Users\\MarkYoung\\Documents\\Tablet CleanUp Return Labels",
                    restaurantNumber: restaurantNumber,
                    toRecipients: toRecipientsList,
                    ccRecipients: ccRecipientList
                );
                // Convert to .eml
                //MimeMessage emlMessage = await ConvertToEmlWithEmbeddedImages(mail, "C:\\Users\\MarkYoung\\Documents\\Tablet CleanUp Return Labels",restaurantNumber);
                
                /*
                emlMessage.From.Clear();
                emlMessage.From.Add(MailboxAddress.Parse("MarkYoung@Bloominbrands.com"));
                */
                
                // Save to EML format
                //using (var stream = File.Create(emlFileName))
                //{
                //    emlMessage.WriteTo(stream);
                //}
                //ConvertEmlToMsg(emlMessage, msgFileName);
                //ConvertEmlToMsgWithOutlook(emlFileName, msgFileName);
                //CleanMsgWithOutlookLateBinding(msgFileName, msgFileName.Replace(".msg", "_clean.msg"));


            }

            return saveDirectory;
        }

        public static void CleanMsgWithOutlookLateBinding(string msgInputPath, string msgOutputPath)
        {
            // Ensure Outlook is installed and accessible
            Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
            if (outlookType == null)
                throw new InvalidOperationException("Outlook is not installed or accessible.");

            // Create Outlook instance (late bound)
            dynamic outlookApp = Activator.CreateInstance(outlookType);

            // Load the .msg file as a MailItem
            dynamic mailItem = outlookApp.CreateItemFromTemplate(msgInputPath);

            if (mailItem == null)
                throw new InvalidOperationException("Outlook failed to load the .msg file.");

            // Save it again (cleans it, links cid: images, sets proper sender)
            const int olMSGUnicode = 3; // Outlook constant for .msg (Unicode format)
            mailItem.SaveAs(msgOutputPath, olMSGUnicode);
        }

        public static void ConvertEmlToMsgWithOutlook(string emlPath, string msgPath)
    {
        var outlookApp = new Outlook.Application();
        Outlook.MailItem mailItem = outlookApp.CreateItemFromTemplate(emlPath) as Outlook.MailItem;

        if (mailItem == null)
            throw new InvalidOperationException("Failed to load .eml file into Outlook.");

        mailItem.SaveAs(msgPath, Outlook.OlSaveAsType.olMSGUnicode);
    }

    private string CreateEmailFiles(Dictionary<string, List<string>> emailDict)
        {
            string saveDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Generated Emails");
            Directory.CreateDirectory(saveDirectory);

            string emailTemplate = LoadEmailTemplate(); // Load the saved email template

            foreach (var recipient in emailDict.Keys)
            {
                string formattedEmail = emailTemplate; // Load the original template
                StringBuilder deviceDetails = new StringBuilder(); // Store all device entries

                Dictionary<string,string> restaurantDetailsDict = new Dictionary<string, string>();
                //This is confusing but, each data string contains all the restaurant data
                //  including the device IP and Serial Number
                //  REFACTOR!!
                foreach (var dataString in emailDict[recipient])
                {
                    // Extract key-value pairs from the CSV data
                    var keyValuePairs = new Dictionary<string, string>();

                    foreach (var entry in dataString.Split(','))
                    {
                        var pair = entry.Split(':', 2); // Ensure we split only on the first colon
                        if (pair.Length == 2)
                        {
                            keyValuePairs[pair[0].Trim()] = pair[1].Trim();
                        }
                    }
                    restaurantDetailsDict = keyValuePairs;
                    // Append device details instead of overwriting the full email
                    deviceDetails.AppendLine($"•\t: {keyValuePairs["Wi-Fi IP Address"]}, Serial: {keyValuePairs["Serial Number"]}");
                }

                // Replace placeholders, ensuring devices are added properly
                formattedEmail = formattedEmail.Replace("{Device List}", deviceDetails.ToString());
                foreach(var key in restaurantDetailsDict.Keys)
                {
                    string placeholder = string.Format("{{{0}}}", key);
                    formattedEmail = formattedEmail.Replace(placeholder, restaurantDetailsDict[key]);
                }
                //Now replace any restaurant-level placeholder from the last restaurant info saved

                // Create email file
                MailMessage mail = new MailMessage();
                mail.To.Add(recipient);
                mail.Subject = "Duplicate Android POSi Tablets in your Restaurant";
                mail.Body = formattedEmail;
                mail.From = new MailAddress("no-reply@example.com");

                string emlContent = ConvertToEml(mail);
                string emlFileName = Path.Combine(saveDirectory, $"{recipient.Replace("@", "_at_").Replace(".", "_")}.eml");
                string msgFileName = Path.Combine(saveDirectory, $"{recipient.Replace("@", "_at_").Replace(".", "_")}.msg");
                File.WriteAllText(emlFileName, emlContent, new UTF8Encoding(false));
                //ConvertEmlToMsg(emlFileName, msgFileName);
            }

            return saveDirectory;
        }
        public static string ReplacePlaceholders(string template, Dictionary<string, string> values)
        {
            foreach (var pair in values)
            {
                string placeholder = $"{{{{{pair.Key}}}}}";
                template = template.Replace(placeholder, pair.Value);
            }
            return template;
        }
        private string LoadEmailTemplate()
        {
            string templatePath = "email_template.txt";
            return File.Exists(templatePath) ? File.ReadAllText(templatePath, Encoding.UTF8) : "Default Email Body";
        }
        private static string? FindHtmlBody(MimeEntity entity)
        {
            if (entity is TextPart text && text.IsHtml)
            {
                return text.Text;
            }

            if (entity is Multipart multipart)
            {
                foreach (var part in multipart)
                {
                    string? result = FindHtmlBody(part);
                    if (!string.IsNullOrWhiteSpace(result))
                        return result;
                }
            }

            return null;
        }

        public static void ConvertEmlToMsg(MimeMessage message, string msgPath)
        {
            string senderEmail = message.From.Mailboxes.FirstOrDefault()?.Address ?? "no-reply@example.com";
            string senderName = message.From.Mailboxes.FirstOrDefault()?.Name ?? "Unknown Sender";
            string emailBody = FindHtmlBody(message.Body) ?? "[No Content]";
            using (var msg = new Email(
                sender: null,
                new Representing("", ""),
                message.Subject,
                draft: true,
                readReceipt: false,
                leaveAttachmentStreamsOpen: false
            ))
            {
                var tempFiles = new List<string>();

                foreach (var part in message.BodyParts.OfType<MimePart>())
                {
                    if (part.Content == null || string.IsNullOrEmpty(part.FileName))
                        continue;

                    using var ms = new MemoryStream();
                    part.Content.DecodeTo(ms);
                    ms.Position = 0;

                    if (!string.IsNullOrEmpty(part.ContentId) &&
                        part.ContentDisposition?.Disposition == ContentDisposition.Inline &&
                        part.ContentType.MediaType == "image")
                    {
                        // Handle inline image
                        string cid = part.ContentId.Trim('<', '>');
                        string base64 = Convert.ToBase64String(ms.ToArray());
                        string pattern = $"cid:{Regex.Escape(cid)}";

                        emailBody = Regex.Replace(
    emailBody,
    $"src=[\"']{pattern}[\"']",
    $"src=\"data:{part.ContentType.MimeType};base64,{base64}\"",
    RegexOptions.IgnoreCase
);

                    }
                    else
                    {
                        // Save to temp file with unique name
                        string tempFilePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}_{part.FileName}");
                        File.WriteAllBytes(tempFilePath, ms.ToArray());
                        msg.Attachments.Add(tempFilePath);
                        tempFiles.Add(tempFilePath);
                    }
                }

                msg.BodyHtml = emailBody;

                foreach (var recipient in message.To.Mailboxes)
                    msg.Recipients.AddTo(recipient.Address, recipient.Name);
                foreach (var recipient in message.Cc.Mailboxes)
                    msg.Recipients.AddCc(recipient.Address, recipient.Name);
                foreach (var recipient in message.Bcc.Mailboxes)
                    msg.Recipients.AddBcc(recipient.Address, recipient.Name);

                msg.Save(msgPath);

                // Clean up temp files
                foreach (var tempFile in tempFiles)
                {
                    try { File.Delete(tempFile); } catch { }
                }
            }
        }
        static void WriteStream(CFStorage storage, string streamName, string value)
        {
            byte[] data = Encoding.Unicode.GetBytes(value);
            CFStream stream = storage.AddStream(streamName);
            stream.SetData(data);
        }
        private async Task<MimeMessage> ConvertToEmlWithEmbeddedImages(MailMessage mail, string pdfFolderPath, string restaurantNumber)
        {
            string path = "missing.log";

            //using (var stream = File.Create(path))
            using (var writer = new StreamWriter(path, append:true))
            {
                var message = new MimeMessage();

            // Set addresses
            //message.From.Add(MailboxAddress.Parse(mail.From.Address));
            foreach (var to in mail.To)
                message.To.Add(MailboxAddress.Parse(to.Address));
            foreach (var cc in mail.CC)
                message.Cc.Add(MailboxAddress.Parse(cc.Address));

            message.Subject = mail.Subject;

            // Mark as draft so it can be uploaded to Outlook as a draft
            message.Headers.Add("X-Unsent", "1");

            // Main body and image processing
            var html = mail.Body;

            var imageRegex = new Regex(
   "<img[^>]+src=\\\"data:image/(?<type>[^;]+);base64,(?<data>.*?)\\\"",
   RegexOptions.IgnoreCase | RegexOptions.Singleline);
            var matches = imageRegex.Matches(html);
            var linkedResources = new List<MimePart>();

            foreach (Match match in matches)
            {
                string imageType = match.Groups["type"].Value;
                string base64Data = match.Groups["data"].Value;
                byte[] imageBytes = Convert.FromBase64String(base64Data);

                string contentId = MimeUtils.GenerateMessageId();

                var mimePart = new MimePart("image", imageType)
                {
                    Content = new MimeContent(new MemoryStream(imageBytes)),
                    ContentId = contentId,
                    ContentDisposition = new ContentDisposition(ContentDisposition.Inline),
                    ContentTransferEncoding = ContentEncoding.Base64,
                    FileName = $"image.{imageType}"
                };

                linkedResources.Add(mimePart);

                string originalImgTag = match.Value;
                string newImgTag = Regex.Replace(originalImgTag, "src=\\\".*?\\\"", $"src=\"cid:{contentId}\"");
                html = html.Replace(originalImgTag, newImgTag);
            }

            // Create the inner related multipart
            var related = new Multipart("related");


            var htmlPart = new TextPart("html");
            htmlPart.SetText(Encoding.UTF8, html);


            related.Add(htmlPart);
            foreach (var resource in linkedResources)
                related.Add(resource);

                // Optional plain text fallback
                var plainText = new TextPart("plain")
                {
                    Text = "This message contains HTML content with embedded images."
                };

                // Wrap HTML + images in multipart/alternative
                var alternative = new Multipart("alternative");
                alternative.Add(plainText);
                alternative.Add(related);

                // Now build the mixed wrapper for attachments
                var mixed = new Multipart("mixed");
                mixed.Add(alternative);


                // Find and attach the matching PDF
                // Find and attach the matching PDF
                var pdfMatches = Directory.GetFiles(pdfFolderPath, "*.pdf")
            .Where(f => Regex.IsMatch(Path.GetFileName(f), $".*{Regex.Escape(restaurantNumber)}.*\\.pdf", RegexOptions.IgnoreCase))
            .ToList();
            if (pdfMatches == null || pdfMatches.Count == 0)
            {
                    /**************************************************
                     * * THIS IS A MOCK REQUEST FOR TESTING ONLY!!!!
                     * *****************************************************/
                    /*                  var shipmentRequest = new ShipmentRequest
                     {

                                          accountNumber = new AccountNumber
                         {
                             value = "740561073"
                         },
                         requestedShipment = new RequestedShipment
                         {
                             shipper = new Shipper
                             {
                                 contact = new EmailGenerator.Models.Contact
                                 {
                                     personName = "MANAGING PARTNER",
                                     phoneNumber = "7349814144"
                                 },
                                 address = new EmailGenerator.Models.Address
                                 {
                                     streetLines = new List<string> { "42871 FORD RD." },
                                     city = "CANTON",
                                     stateOrProvinceCode = "MI", // Michigan abbreviation
                                     postalCode = "48187",
                                     countryCode = "US"
                                 }
                             },
                             recipients = new List<EmailGenerator.Models.Recipient>
                             {
                                 new EmailGenerator.Models.Recipient
                                 {
                                     contact = new EmailGenerator.Models.Contact
                                     {
                                         personName = "MTECH MOBILITY",
                                         phoneNumber = "8448642463"
                                     },
                                     address = new EmailGenerator.Models.Address
                                     {
                                         streetLines = new List<string> { "15827 GUILD COURT" },
                                         city = "JUPITER",
                                         stateOrProvinceCode = "FL", // Florida abbreviation
                                         postalCode = "33478",
                                         countryCode = "US"
                                     }
                                 }
                             },
                             serviceType = "STANDARD_OVERNIGHT",
                             packagingType = "YOUR_PACKAGING",
                             pickupType = "DROPOFF_AT_FEDEX_LOCATION",
                             shippingChargesPayment = new ShippingChargesPayment
                             {
                                 paymentType = "THIRD_PARTY",
                                 payor = new Payor
                                 {
                                     responsibleParty = new ResponsibleParty
                                     {
                                         accountNumber = new AccountNumber
                                         {
                                             value = "740561073",
                                             key = ""
                                         }
                                     },
                                     address = new EmailGenerator.Models.Address
                                     {
                                         streetLines = new List<string> { "42871 FORD RD." },
                                         city = "CANTON",
                                         stateOrProvinceCode = "MI",
                                         postalCode = "48187",
                                         countryCode = "US"
                                     }
                                 }
                             },
                             labelSpecification = new LabelSpecification(), // Fill in if needed
                             requestedPackageLineItems = new List<RequestedPackageLineItem>
                             {
                                 new RequestedPackageLineItem
                                 {
                                     weight = new Weight
                                     {
                                         units = "LB",
                                         value = "15"
                                     }
                                 }
                             }
                         }
                     };



                                                    byte[] pdfBytes = await _shippingService.CreateShipmentLabelAsync(shipmentRequest);
                                     // base64EncodedBytes is your byte[] containing base64-encoded data
                                     string base64String = Encoding.UTF8.GetString(pdfBytes);

                                     // Decode the base64 string back to raw bytes
                                     byte[] decodedBytes = Convert.FromBase64String(base64String);

                                     File.WriteAllBytes($"{mail.To}.label.pdf", decodedBytes);







                            */
                    writer.WriteLine($"No PDF found for {mail.To}: {restaurantNumber}");
                }
            else
            {
                if (pdfMatches.Count == 1)
                {
                    var pdfMatch = pdfMatches.First();
                    byte[] pdfBytes = File.ReadAllBytes(pdfMatch);
                    var attachment = new MimePart("application", "pdf")
                    {
                        Content = new MimeContent(new MemoryStream(pdfBytes), ContentEncoding.Base64),
                        ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                        ContentTransferEncoding = ContentEncoding.Base64,
                        FileName = Path.GetFileName(pdfMatch)
                    };

                    mixed.Add(attachment);
                }
                else if (pdfMatches.Count > 1)
                {
                    var dialog = new ContentDialog
                    {
                        Title = $"Select PDF for {restaurantNumber}",
                        PrimaryButtonText = "OK",
                        CloseButtonText = "Cancel"

                    };

                    var comboBox = new ComboBox { ItemsSource = pdfMatches.Select(Path.GetFileName).ToList() };
                    dialog.Content = comboBox;
                    dialog.XamlRoot = RootGrid.XamlRoot;
                    var result = await dialog.ShowAsync();
                    if (result == ContentDialogResult.Primary && comboBox.SelectedItem != null)
                    {
                        var selectedPdf = pdfMatches[comboBox.SelectedIndex];
                        byte[] pdfBytes = File.ReadAllBytes(selectedPdf);
                        var attachment = new MimePart("application", "pdf")
                        {
                            Content = new MimeContent(new MemoryStream(pdfBytes), ContentEncoding.Base64),
                            ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                            ContentTransferEncoding = ContentEncoding.Base64,
                            FileName = Path.GetFileName(selectedPdf)
                        };

                        mixed.Add(attachment);
                    }
                }
            }
            message.Body = mixed;


            return message;
        }
        }
        private MimeMessage ConvertToEmlWithEmbeddedImages(MailMessage mail)
        {
            var message = new MimeMessage();

            // Set addresses
            message.From.Add(MailboxAddress.Parse(mail.From.Address));
            foreach (var to in mail.To)
                message.To.Add(MailboxAddress.Parse(to.Address));
            message.Subject = mail.Subject;

            // Mark as draft so it can be uploaded to Outlook as a draft
            message.Headers.Add("X-Unsent", "1");


            // Main body and image processing
            var html = mail.Body;
            var builder = new BodyBuilder();

            var imageRegex = new Regex(
    "<img[^>]+src=\\\"data:image/(?<type>[^;]+);base64,(?<data>.*?)\\\"",
    RegexOptions.IgnoreCase | RegexOptions.Singleline);
            
            var matches = imageRegex.Matches(html);
            var cidMap = new Dictionary<string, string>();

            var linkedResources = new List<MimePart>();

            foreach (Match match in matches)
            {
                string imageType = match.Groups["type"].Value;
                string base64Data = match.Groups["data"].Value;
                base64Data = base64Data.Replace("\n", "").Replace("\r", "");

                byte[] imageBytes = Convert.FromBase64String(base64Data);

                string contentId = MimeUtils.GenerateMessageId();

                var mimePart = new MimePart("image", imageType)
                {
                    Content = new MimeContent(new MemoryStream(imageBytes)),
                    ContentId = contentId,
                    ContentDisposition = new ContentDisposition(ContentDisposition.Inline),
                    ContentTransferEncoding = ContentEncoding.Base64,
                    FileName = $"image.{imageType}"
                };

                linkedResources.Add(mimePart);

                string originalImgTag = match.Value;
                string newImgTag = Regex.Replace(originalImgTag, "src=\\\".*?\\\"", $"src=\"cid:{contentId}\"");
                html = html.Replace(originalImgTag, newImgTag);
            }

            var related = new Multipart("related");

            var htmlPart = new TextPart("html");
            htmlPart.SetText(Encoding.UTF8, html);


            related.Add(htmlPart);
            foreach (var resource in linkedResources)
                related.Add(resource);

            // Wrap with a mixed multipart to support file attachments later
            var mixed = new Multipart("mixed");
            mixed.Add(related);

            message.Body = mixed;


            return message;
        }

        private void OpenSettings_Click(object sender, RoutedEventArgs e)
        {
            var settingsWindow = new Microsoft.UI.Xaml.Window();
            settingsWindow.Content = new EmailGenerator.Views.SettingsEditorView(); // Use correct namespace
            settingsWindow.Activate();
        }

        private string ConvertToEml(MailMessage mail)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                using (StreamWriter writer = new StreamWriter(memoryStream, Encoding.UTF8))
                {
                    writer.WriteLine($"Date: {DateTime.UtcNow.ToString("ddd, dd MMM yyyy HH:mm:ss +0000")}");
                    writer.WriteLine($"From: {mail.From.Address}");
                    writer.WriteLine($"To: {string.Join(", ", mail.To)}");
                    writer.WriteLine($"Subject: {mail.Subject}");
                    writer.WriteLine("MIME-Version: 1.0");
                    writer.WriteLine("Content-Type: text/html; charset=UTF-8");
                    writer.WriteLine("Content-Transfer-Encoding: quoted-printable");
                    writer.WriteLine();
                    writer.WriteLine(mail.Body);
                    writer.Flush();
                    return Encoding.UTF8.GetString(memoryStream.ToArray());
                }
            }
        }

        private string GetConceptAbbreviationFromName(string conceptName)
        {
            var conceptDict = new Dictionary<string, string>
            {
                { "Outback", "OBS" },
                { "Bonefish", "BFG" },
                { "Carrabbas", "CIG" },
                { "Flemmings", "FLM" }
            };

            return conceptDict.TryGetValue(conceptName, out string abbreviation) ? abbreviation : conceptName;
        }
    }
}
