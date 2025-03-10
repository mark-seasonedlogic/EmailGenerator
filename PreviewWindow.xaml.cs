using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using MimeKit;
using Windows.Storage;
using Windows.ApplicationModel.DataTransfer;
using MsgKit;
using msgReader = MsgReader;
using MsgReader.Outlook;

namespace OutlookDeviceEmailer
{
    public sealed partial class PreviewWindow : Window
    {
        private List<string> emailFiles;
        private int currentIndex = 0;

        public PreviewWindow(string directoryPath)
        {
            InitializeComponent();
            LoadEmails(directoryPath);
        }

        private void LoadEmails(string directoryPath)
        {
            if (Directory.Exists(directoryPath))
            {
                emailFiles = Directory.GetFiles(directoryPath, "*.msg").ToList();
                if (emailFiles.Count > 0)
                {
                    DisplayMsgEmail(currentIndex);
                }
                else
                {
                    txtEmailBody.Text = "No emails found.";
                }
            }
        }

        private void Prev_Click(object sender, RoutedEventArgs e)
        {
            if (currentIndex > 0)
            {
                currentIndex--;
                DisplayMsgEmail(currentIndex);
            }
        }

        private void Next_Click(object sender, RoutedEventArgs e)
        {
            if (currentIndex < emailFiles.Count - 1)
            {
                currentIndex++;
                DisplayMsgEmail(currentIndex);
            }
        }
        private MimeMessage LoadEmailWithFixedDate(string filePath)
        {
            string[] lines = File.ReadAllLines(filePath, Encoding.UTF8);

            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].StartsWith("Date:", StringComparison.OrdinalIgnoreCase))
                {
                    // Replace with valid RFC 2822 date format
                    lines[i] = "Date: " + DateTime.UtcNow.ToString("ddd, dd MMM yyyy HH:mm:ss +0000");
                }
            }

            // Write back the corrected file
            File.WriteAllLines(filePath, lines, Encoding.UTF8);

            // Load the email with MimeKit
            return MimeMessage.Load(filePath);
        }
        private MimeMessage LoadEmailWithoutDateHeader(string filePath)
        {
            string[] lines = File.ReadAllLines(filePath, Encoding.UTF8);
            List<string> cleanedLines = new List<string>();

            foreach (string line in lines)
            {
                if (!line.StartsWith("Date:", StringComparison.OrdinalIgnoreCase)) // Remove invalid Date header
                {
                    cleanedLines.Add(line);
                }
            }

            File.WriteAllLines(filePath, cleanedLines, Encoding.UTF8);

            // Load the email without a Date header
            return MimeMessage.Load(filePath);
        }

        private void DisplayEmlEmail(int index)
        {
            if (index >= 0 && index < emailFiles.Count)
            {
                string emailPath = emailFiles[index];

                try
                {
                    // Read file with UTF-8 encoding and clean invalid characters
                    /* string cleanedContent = File.ReadAllText(emailPath, Encoding.UTF8)
                         .Replace("\r\n\r\n", "\n") // Remove extra new lines
                         .Replace("\r\n", "\n") // Normalize line endings
                         .Trim(); // Remove extra spaces

                     // Write cleaned content back to file
                     File.WriteAllText(emailPath, cleanedContent, Encoding.UTF8);
                    */
                     // Load email after cleanup
                     var message = MimeMessage.Load(emailPath);
                    
                    //var message = LoadEmailWithoutDateHeader(emailPath);
                    
                    // Display email details
                    txtEmailSubject.Text = message.Subject ?? "[No Subject]";
                    txtEmailTo.Text = "To: " + (message.To?.ToString() ?? "[Unknown Recipient]");
                    txtEmailBody.Text = message.TextBody ?? message.HtmlBody ?? "[No body content]";

                    // Update email index display
                    txtEmailIndex.Text = $"{index + 1} / {emailFiles.Count}";
                }
                catch (Exception ex)
                {
                    txtEmailBody.Text = $"Error loading email: {ex.Message}";
                }
            }
        }

        private void DisplayMsgEmail(int index)
        {
            if (index >= 0 && index < emailFiles.Count)
            {
                string emailPath = emailFiles[index];

                try
                {
                    // Load MSG file using MsgReader
                    using (var msgFile = new Storage.Message(emailPath))
                    {
                        // Extract email details
                        txtEmailSubject.Text = msgFile.Subject ?? "[No Subject]";
                        txtEmailTo.Text = "To: " + (msgFile.Recipients.Count > 0 ? msgFile.Recipients[0].Email : "[Unknown Recipient]");
                        txtEmailBody.Text = msgFile.BodyText ?? "[No body content]";

                        // Update email index display
                        txtEmailIndex.Text = $"{index + 1} / {emailFiles.Count}";
                    }
                }
                catch (Exception ex)
                {
                    txtEmailBody.Text = $"Error loading email: {ex.Message}";
                }
            }
        }
        private void Border_DragOver(object sender, DragEventArgs e)
        {
            e.AcceptedOperation = DataPackageOperation.Copy;
        }

        private async void Border_Drop(object sender, DragEventArgs e)
        {
            var items = await e.DataView.GetStorageItemsAsync();
            if (items.Any())
            {
                StorageFile file = items.First() as StorageFile;
                if (file != null && file.FileType.ToLower() == ".pdf")
                {
                    string pdfPath = file.Path;
                    AttachPdfToMsg(pdfPath);
                }
            }
        }
        private void AttachPdfToEml(string pdfPath)
        {
            if (currentIndex >= 0 && currentIndex < emailFiles.Count)
            {
                string emailPath = emailFiles[currentIndex];
                var message = MimeMessage.Load(emailPath);

                // Check if the current body is already multipart
                Multipart multipart;
                if (message.Body is Multipart existingMultipart)
                {
                    multipart = existingMultipart;
                }
                else
                {
                    // Create a new multipart container to wrap existing content
                    multipart = new Multipart("mixed");
                    if (message.Body != null)
                    {
                        multipart.Add(message.Body);
                    }
                }

                // Create a new attachment for the PDF
                var attachment = new MimePart("application", "pdf")
                {
                    Content = new MimeContent(File.OpenRead(pdfPath), ContentEncoding.Base64),
                    ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                    ContentTransferEncoding = ContentEncoding.Base64,
                    FileName = Path.GetFileName(pdfPath)
                };

                // Add the attachment to the multipart container
                multipart.Add(attachment);

                // Set the new body with attachments
                message.Body = multipart;

                // Save the modified email (ensure encoding without BOM)
                File.WriteAllText(emailPath, message.ToString(), new UTF8Encoding(false));

                // Confirm the attachment was added
                txtEmailBody.Text += $"\n\n[Attached: {Path.GetFileName(pdfPath)}]";
            }
        }
        private void AttachPdfToMsg(string pdfPath)
        {
            if (currentIndex >= 0 && currentIndex < emailFiles.Count)
            {
                string msgPath = emailFiles[currentIndex];

                // Load the existing MSG file
                using (var msgFile = new Storage.Message(msgPath))
                {
                    // Extract existing email details
                    string senderEmail = msgFile.Sender?.Email ?? "no-reply@example.com";
                    string senderName = msgFile.Sender?.DisplayName ?? "Unknown Sender";
                    string subject = msgFile.Subject;
                    string body = msgFile.BodyText;

                    // Create a new MSG with the same details
                    using (var msg = new Email(
                        new Sender(senderEmail, senderName),
                        new Representing("", ""),
                        subject,
                        draft: true,
                        readReceipt: false,
                        leaveAttachmentStreamsOpen: false
                    ))
                    {
                        msg.BodyText = body;

                        // Re-add recipients
                        foreach (var recipient in msgFile.Recipients)
                        {
                            msg.Recipients.AddTo(recipient.Email, recipient.DisplayName);
                        }

                        // Add existing attachments properly
                        foreach (var attachment in msgFile.Attachments)
                        {
                            if (attachment is Storage.Attachment fileAttachment)
                            {
                                // Write the attachment to a temp file
                                string tempFilePath = Path.Combine(Path.GetTempPath(), fileAttachment.FileName);
                                File.WriteAllBytes(tempFilePath, fileAttachment.Data);

                                // Attach the file from disk
                                msg.Attachments.Add(tempFilePath);
                            }
                        }

                        // Attach the new PDF
                        msg.Attachments.Add(pdfPath);

                        // Save the modified MSG file (overwrite the original)
                        msg.Save(msgPath);
                    }
                }

                // Confirm the attachment was added
                txtEmailBody.Text += $"\n\n[Attached: {Path.GetFileName(pdfPath)}]";
            }
        }
    }
}
