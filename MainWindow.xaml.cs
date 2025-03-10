using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
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

namespace OutlookDeviceEmailer
{
    public partial class MainWindow : Window
    {
        private string deviceFilePath;
        private string emailFilePath;

        public MainWindow()
        {
            InitializeComponent();
            ApplyAcrylicEffect();
            SetRoundedCorners();
            SetWindowSize(600, 400); // Set width and height
        }

        private void SetWindowSize(int width, int height)
        {
            IntPtr hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WindowId myWndId = Win32Interop.GetWindowIdFromWindow(hWnd);
            AppWindow appWindow = AppWindow.GetFromWindowId(myWndId);

            appWindow.Resize(new SizeInt32(width, height));
        }
        private void EditTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(deviceFilePath) || string.IsNullOrEmpty(emailFilePath))
            {
                ShowMessage("Error", "Please select both CSV files before editing the template.");
                return;
            }

            EmailEditorWindow editorWindow = new EmailEditorWindow(deviceFilePath, emailFilePath);
            editorWindow.Activate();
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

            string generatedEmailDirectory = CreateEmailFiles(emailDict); // Now returns directory

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

                // Get column headers dynamically
                var headers = csv.HeaderRecord.ToList();

                while (csv.Read())
                {
                    // Extract `CONCEPT_CD` and `RSTRNT_NBR`
                    string conceptCode = csv.GetField("CONCEPT_CD")?.Trim();
                    string restaurantNumber = csv.GetField("RSTRNT_NBR")?.Trim();

                    // Validate values
                    if (string.IsNullOrEmpty(conceptCode) || string.IsNullOrEmpty(restaurantNumber))
                        continue;

                    // Convert `CONCEPT_CD` and pad `RSTRNT_NBR` to 4 digits
                    string conceptAbbreviation = GetConceptAbbreviationFromCode(conceptCode);
                    string paddedRestaurantNumber = restaurantNumber.PadLeft(4, '0');

                    // Create the unique lookup key
                    string lookupKey = $"{paddedRestaurantNumber}{conceptAbbreviation}";

                    // Store all restaurant details dynamically
                    var restaurantData = new Dictionary<string, string>();

                    foreach (var header in headers)
                    {
                        string value = csv.GetField(header)?.Trim();
                        restaurantData[header] = value; // Store dynamically
                    }

                    emailLookup[lookupKey] = restaurantData;
                }
            }

            return emailLookup;
        }

        private Dictionary<string, List<string>> GenerateEmailDictionary(string deviceFile, Dictionary<string, Dictionary<string, string>> emailLookup)
        {
            var emailDict = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

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
                    string storeNumber = userName.Substring(3, 4);
                    string concept = userName.Substring(0, 3);
                    string ip = csv.GetField("Wi-Fi IP Address")?.Trim();
                    string serial = csv.GetField("Serial Number")?.Trim();

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
                            restaurantInfo.Where(kv => kv.Key != "STORE_EMAIL_ADDR") // Exclude email
                                          .Select(kv => $"{kv.Key}: {kv.Value}")) +
                            $", Wi-Fi IP Address: {ip}, Serial Number: {serial}";

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
                if (csv.Read())
                {
                    csv.ReadHeader();
                }
                else
                {
                    throw new Exception("Email CSV file is empty.");
                }

                while (csv.Read())
                {
                    string restaurantName = csv.GetField("RSTRNT_LEGAL_NAME")?.Trim();
                    string email = csv.GetField("STORE_EMAIL_ADDR")?.Trim();
                    string address = csv.GetField("ADDR_LINE1_TXT")?.Trim();
                    string city = csv.GetField("CITY_NAME")?.Trim();
                    string state = csv.GetField("STATE_CD")?.Trim();
                    string phone = csv.GetField("STORE_PHONE_NO")?.Trim();

                    if (!string.IsNullOrEmpty(restaurantName) && !string.IsNullOrEmpty(email))
                    {
                        restaurantData[restaurantName] = new Dictionary<string, string>
                {
                    { "STORE_EMAIL_ADDR", email },
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
                    deviceDetails.AppendLine($"- IP: {keyValuePairs["Wi-Fi IP Address"]}, Serial: {keyValuePairs["Serial Number"]}");
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
                mail.Subject = "Device & Restaurant Information Update";
                mail.Body = formattedEmail;
                mail.From = new MailAddress("no-reply@example.com");

                string emlContent = ConvertToEml(mail);
                string emlFileName = Path.Combine(saveDirectory, $"{recipient.Replace("@", "_at_").Replace(".", "_")}.eml");
                string msgFileName = Path.Combine(saveDirectory, $"{recipient.Replace("@", "_at_").Replace(".", "_")}.msg");
                File.WriteAllText(emlFileName, emlContent, new UTF8Encoding(false));
                ConvertEmlToMsg(emlFileName, msgFileName);
            }

            return saveDirectory;
        }
        private string LoadEmailTemplate()
        {
            string templatePath = "email_template.txt";
            return File.Exists(templatePath) ? File.ReadAllText(templatePath, Encoding.UTF8) : "Default Email Body";
        }
        static void ConvertEmlToMsg(string emlPath, string msgPath)
        {
            // Load the EML file using MimeKit
            var message = MimeMessage.Load(emlPath);

            // Extract sender details safely
            string senderEmail = message.From.Mailboxes.FirstOrDefault()?.Address ?? "no-reply@example.com";
            string senderName = message.From.Mailboxes.FirstOrDefault()?.Name ?? "Unknown Sender";

            // Extract email body (prefer HTML if available)
            string emailBody = message.HtmlBody ?? message.TextBody ?? "[No Content]";

            // Create a new MSG email with the correct constructor
            using (var msg = new Email(
                new Sender(senderEmail, senderName), // Sender info
                new Representing("", ""), // Representing field (optional)
                message.Subject, // Email subject
                draft: true, // IsDraft flag
                readReceipt: false, // Read receipt requested
                leaveAttachmentStreamsOpen: false // Keep attachment streams closed after saving
            ))
            {
                msg.BodyText = emailBody; // Set the email body content

                // Add recipients (To)
                foreach (var recipient in message.To.Mailboxes)
                {
                    msg.Recipients.AddTo(recipient.Address, recipient.Name);
                }

                // Add CC
                foreach (var recipient in message.Cc.Mailboxes)
                {
                    msg.Recipients.AddCc(recipient.Address, recipient.Name);
                }

                // Add BCC
                foreach (var recipient in message.Bcc.Mailboxes)
                {
                    msg.Recipients.AddBcc(recipient.Address, recipient.Name);
                }

                // Attachments Handling
                foreach (var attachment in message.Attachments.OfType<MimePart>())
                {
                    if (attachment.Content != null)
                    {
                        string tempFilePath = Path.Combine(Path.GetTempPath(), attachment.FileName);

                        // Save attachment as a temporary file
                        using (var fileStream = File.Create(tempFilePath))
                        {
                            attachment.Content.DecodeTo(fileStream);
                        }

                        // Add attachment by file path (MsgKit requires string, not byte[])
                        msg.Attachments.Add(tempFilePath);
                    }
                }

                // Save the MSG file
                msg.Save(msgPath);
            }
        }
        static void WriteStream(CFStorage storage, string streamName, string value)
        {
            byte[] data = Encoding.Unicode.GetBytes(value);
            CFStream stream = storage.AddStream(streamName);
            stream.SetData(data);
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
                    writer.WriteLine("Content-Type: text/plain; charset=UTF-8");
                    writer.WriteLine("Content-Transfer-Encoding: quoted-printable");
                    writer.WriteLine();
                    writer.WriteLine(mail.Body);
                    writer.Flush();
                    return Encoding.UTF8.GetString(memoryStream.ToArray());
                }
            }
        }
        private string GetConceptAbbreviationFromCode(string conceptCode)
        {
            var conceptMapping = new Dictionary<string, string>
    {
        { "1", "OBS" },
        { "2", "FPS" },
        { "3", "ROY" },
        { "4", "DOC" },
        { "6", "BFG" },
        { "7", "CIG" },
        { "10", "INT" },
        { "99", "ALC" }
    };

            return conceptMapping.TryGetValue(conceptCode, out string abbreviation) ? abbreviation : "UNKNOWN";
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
