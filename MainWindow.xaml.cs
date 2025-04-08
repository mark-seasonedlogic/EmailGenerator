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
using EmailGenerator.Models.Settings;
using EmailGenerator.Views;

namespace OutlookDeviceEmailer
{
    public partial class MainWindow : Window
    {
        public readonly IConfiguration _config;

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
            this.ExtendsContentIntoTitleBar = true;

            var hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WindowId windowId = Win32Interop.GetWindowIdFromWindow(hWnd);
            AppWindow appWindow = AppWindow.GetFromWindowId(windowId);

            appWindow.SetPresenter(AppWindowPresenterKind.Overlapped);
            var displayArea = DisplayArea.GetFromWindowId(windowId, DisplayAreaFallback.Primary);
            appWindow.MoveAndResize(displayArea.WorkArea);
            MainFrame.Navigate(typeof(DashboardView)); // Default page
            appWindow.Title = "Device Detail Mail Merge";
        }
        private void NavigationView_SelectionChanged(NavigationView sender, NavigationViewSelectionChangedEventArgs args)
        {
            if (args.IsSettingsSelected)
            {
                MainFrame.Navigate(typeof(SettingsEditorView));
            }
            else if (args.SelectedItem is NavigationViewItem selectedItem)
            {
                switch (selectedItem.Tag)
                {
                    case "Editor":
                        // Replace with your editor view when available
                        // MainFrame.Navigate(typeof(HtmlEditorView));
                        break;
                    case "Generator":
                        MainFrame.Navigate(typeof(EmailGeneratorView));
                        break;
                    case "Settings":
                        MainFrame.Navigate(typeof(SettingsEditorView));
                        break;
                    case "dashboard":
                        MainFrame.Navigate(typeof(EmailGenerator.Views.DashboardView));
                        break;

                }
            }
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

                        // Merge restaurant & device data into a single base64-encoded key-value string
                        var formattedDataParts = restaurantInfo
                            .Where(kv => kv.Key != "STORE_EMAIL_ADDR")
                            .Select(kv => $"{kv.Key}: {Convert.ToBase64String(Encoding.UTF8.GetBytes(kv.Value ?? ""))}")
                            .ToList();

                        // Add device-specific values (also base64-encoded)
                        formattedDataParts.Add($"MD + Serial: {Convert.ToBase64String(Encoding.UTF8.GetBytes(ip ?? ""))}");
                        formattedDataParts.Add($"Serial Number: {Convert.ToBase64String(Encoding.UTF8.GetBytes(serial ?? ""))}");
                        formattedDataParts.Add($"JVP EMAIL: {Convert.ToBase64String(Encoding.UTF8.GetBytes(jvpEmail ?? ""))}");
                        formattedDataParts.Add($"MVP EMAIL: {Convert.ToBase64String(Encoding.UTF8.GetBytes(mvpEmail ?? ""))}");

                        string formattedData = string.Join(", ", formattedDataParts);



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
            try
            {
                var settingsWindow = new Microsoft.UI.Xaml.Window();
                settingsWindow.Content = new EmailGenerator.Views.SettingsEditorView(); // Use correct namespace
                settingsWindow.Activate();
            }
            catch(Exception ex)
            {
                Debug.WriteLine($"SettingsEditorView error: {ex}");
                throw;
            }
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
