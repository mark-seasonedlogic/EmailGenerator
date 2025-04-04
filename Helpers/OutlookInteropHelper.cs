using Microsoft.Extensions.FileSystemGlobbing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text.RegularExpressions;

namespace EmailGenerator.Helpers
{
    public static class OutlookInteropHelper
    {
        public static void GenerateMsgWithEmbeddedImagesAndPdf(
              string subject,
              string htmlInput,
              string imageDirectory,
              string msgOutputPath,
              string pdfFolderPath,
              string restaurantNumber,
              List<MailAddress> toRecipients,
              List<MailAddress> ccRecipients,
              Action<string>? log = null)
        {
            // Match all linked image src paths (e.g., Untitled_files/image001.jpg)
            var imgTagRegex = new Regex("<img[^>]+src=\"(?<src>[^\"]+)\"", RegexOptions.IgnoreCase);
            var matches = imgTagRegex.Matches(htmlInput);

            var inlineImages = new List<(string FilePath, string ContentId)>();
            string updatedHtml = htmlInput;

            foreach (Match match in matches)
            {
                string originalSrc = match.Groups["src"].Value;
                string fileName = Path.GetFileName(originalSrc);
                // Convert to full path using working directory if relative
                string fullPath = Path.IsPathRooted(originalSrc)
                    ? originalSrc
                    : Path.Combine(imageDirectory, fileName);
                if (!File.Exists(fullPath))
                {
                    log?.Invoke($"Missing image: {fullPath}");
                    continue;
                }

                string cid = Guid.NewGuid().ToString() + "@bloomin";
                updatedHtml = updatedHtml.Replace(originalSrc, $"cid:{cid}");
                inlineImages.Add((fullPath, cid));
            }

            // Launch Outlook (late bound)
            Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
            if (outlookType == null)
                throw new InvalidOperationException("Outlook is not installed or not available via COM.");

            dynamic outlookApp = Activator.CreateInstance(outlookType);
            dynamic mailItem = outlookApp.CreateItem(0); // olMailItem

            mailItem.Subject = subject;
            mailItem.HTMLBody = updatedHtml;

            // Add To recipients
            foreach (var to in toRecipients)
            {
                mailItem.To += to.Address + ";";
            }

            // Add CC recipients
            foreach (var cc in ccRecipients)
            {
                mailItem.CC += cc.Address + ";";
            }

            // Embed inline images
            foreach (var image in inlineImages)
            {
                dynamic attachment = mailItem.Attachments.Add(
                    image.FilePath,
                    1, // olByValue
                    Type.Missing,
                    Path.GetFileName(image.FilePath)
                );

                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                    image.ContentId);
            }

            // Attach matching PDF
            var pdfMatches = Directory.GetFiles(pdfFolderPath, "*.pdf")
                .Where(f => Regex.IsMatch(Path.GetFileName(f), $".*{Regex.Escape(restaurantNumber)}.*\\.pdf", RegexOptions.IgnoreCase))
                .ToList();

            if (pdfMatches.Count == 0)
            {
                log?.Invoke($"No PDF found for restaurant {restaurantNumber}");
            }
            else
            {
                string selectedPdf = pdfMatches.First(); // Or prompt if multiple
                dynamic pdfAttachment = mailItem.Attachments.Add(
                    selectedPdf,
                    1, // olByValue
                    Type.Missing,
                    Path.GetFileName(selectedPdf)
                );
            }

            // Save as .msg
            const int olMSGUnicode = 3;
            mailItem.SaveAs(msgOutputPath, olMSGUnicode);
        }
    }
}
