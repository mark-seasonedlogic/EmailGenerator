using EmailGenerator.Helpers;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace EmailGenerator.Services
{
    /// <summary>
    /// Responsible for generating and saving HTML-based email content from template and structured data.
    /// </summary>
    public class EmailBuilderService
    {
        private readonly string _templateContentPath;
        private readonly string _templateShellPath;
        private readonly string _outputDirectory;
        private readonly string _pdfFolderPath;
        private readonly string _imageFolderPath;
        private readonly string _emailSubject;
        private readonly IConfiguration _config;
        private  List<MailAddress> _ccRecipientList;
        private  List<MailAddress> _toRecipientList;
        private  List<MailAddress> _bccRecipientList;

        public EmailBuilderService(IConfiguration config)
        {
            _templateContentPath = config["Email:HtmlTemplateContentPath"];
            _templateShellPath = config["Email:HtmlTemplateShellPath"];
            _pdfFolderPath = config["Email:PdfFolderPath"];
            _imageFolderPath = config["Email:ImageFolderPath"];
            _pdfFolderPath = config["Email:PdfFolderPath"];
            _emailSubject = config["Email:SubjectLine"];
            _config = config;

            _outputDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Generated Emails (HTML)");
        }
        /// <summary>
        /// Converts a saved .html email to a .msg file using Outlook automation.
        /// </summary>
        public void ConvertHtmlToMsgWithOutlook(string htmlPath, string msgPath)
        {
            var outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.MailItem mailItem = outlookApp.CreateItemFromTemplate(htmlPath) as Microsoft.Office.Interop.Outlook.MailItem;

            if (mailItem == null)
                throw new InvalidOperationException("Failed to load .html file into Outlook.");

            mailItem.SaveAs(msgPath, Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSGUnicode);
        }

        /// <summary>
        /// Creates email files using a template and dictionary of base64-encoded device info strings.
        /// </summary>
        public async Task<string> CreateEmailFilesFromHtmlAsync(
            Dictionary<string, List<string>> emailDict,
            Func<string, string> fieldResolver)
        {
            Directory.CreateDirectory(_outputDirectory);

            string emailTemplate = File.Exists(_templateContentPath)
                ? await File.ReadAllTextAsync(_templateContentPath)
                : "<p>[No HTML Template Found]</p>";

            foreach (var recipient in emailDict.Keys)
            {
                string formattedEmail = emailTemplate;
                var deviceDetails = new StringBuilder();
                Dictionary<string, string> lastDevice = null;
                //Set recipients
                string _staticCcRecipients = _config["Email:StaticCcRecipients"] ?? "";
                _ccRecipientList = _staticCcRecipients
                    .Split(';', StringSplitOptions.RemoveEmptyEntries)
                    .Select(address => new MailAddress(address.Trim()))
                    .ToList();
                string _staticToRecipients = _config["Email:StaticToRecipients"] ?? "";
                _toRecipientList = _staticToRecipients
                    .Split(';', StringSplitOptions.RemoveEmptyEntries)
                    .Select(address => new MailAddress(address.Trim()))
                    .ToList();
                string _staticBccRecipients = _config["Email:StaticBccRecipients"] ?? "";
                _bccRecipientList = _staticBccRecipients
                    .Split(';', StringSplitOptions.RemoveEmptyEntries)
                    .Select(address => new MailAddress(address.Trim()))
                    .ToList();

                foreach (var encoded in emailDict[recipient])
                {
                    string json = Encoding.UTF8.GetString(Convert.FromBase64String(encoded));
                    var data = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                    lastDevice = data;

                    if (data != null && data.TryGetValue("MD + Serial", out var mdSerial))
                    {
                        deviceDetails.AppendLine($"<li>{mdSerial}</li>");
                    }
                }

                // Insert device list into template
                formattedEmail = formattedEmail.Replace("{{Device List}}", $"<ul>{deviceDetails}</ul>");

                // Replace other tokens using lastDevice info
                if (lastDevice != null)
                {
                    foreach (var kv in lastDevice)
                    {
                        formattedEmail = formattedEmail.Replace($"{{{{{kv.Key}}}}}", kv.Value);
                    }
                }
                // process recipient fields:
                _toRecipientList.Add(new MailAddress(recipient));
                // Save email as .msg file
                string safeName = recipient.Replace("@", "_at_").Replace(".", "_");
                string msgFileName = Path.Combine(_outputDirectory, $"{safeName}.msg");
                string restaurantNumber = lastDevice != null && lastDevice.TryGetValue("RSTRNT_NBR", out var rNum) ? rNum.PadLeft(4, '0') : "";
                
                OutlookInteropHelper.GenerateMsgWithEmbeddedImagesAndPdf(
    subject: _emailSubject,
    htmlInput: formattedEmail,
    imageDirectory: _imageFolderPath,  // make this configurable too
    msgOutputPath: msgFileName,
    pdfFolderPath: _pdfFolderPath,
    restaurantNumber: restaurantNumber,
    toRecipients: _toRecipientList,//new List<MailAddress> { new MailAddress(recipient) },
    ccRecipients: _ccRecipientList//new List<MailAddress>() // optional or pulled from settings
);

            }

            return _outputDirectory;
        }
        /// <summary>
        /// Builds a MailItem in Outlook using late binding and saves it as a .msg file.
        /// </summary>
        public void ConvertHtmlToMsgWithOutlook(string formattedEmail, string recipient, string msgPath, string restaurantNumber)
        {
            Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
            if (outlookType == null)
                throw new InvalidOperationException("Outlook is not installed or accessible.");

            dynamic outlookApp = Activator.CreateInstance(outlookType);
            dynamic mailItem = outlookApp.CreateItem(0); // 0 = olMailItem

            mailItem.HTMLBody = formattedEmail;
            mailItem.To = recipient;
            mailItem.Subject = "Duplicate Android POSi Tablets in your Restaurant";

            if (!string.IsNullOrEmpty(restaurantNumber) && Directory.Exists(_pdfFolderPath))
            {
                var pdfMatch = Directory.GetFiles(_pdfFolderPath, $"*{restaurantNumber}*.pdf").FirstOrDefault();
                if (!string.IsNullOrEmpty(pdfMatch))
                {
                    mailItem.Attachments.Add(pdfMatch);
                }
            }


            mailItem.SaveAs(msgPath, 3); // 3 = olMSGUnicode
        }
    }
}
