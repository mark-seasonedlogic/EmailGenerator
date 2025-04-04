// Updated EmailHTMLEditorWindow.xaml.cs to add WebView2 preview and live HTML loading

using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.Web.WebView2.Core;
using Microsoft.UI.Xaml.Controls;
using System;
using System.IO;
using System.Text;
using Windows.Graphics;
using System.Collections.Generic;

namespace OutlookDeviceEmailer
{
    public sealed partial class EmailHTMLEditorWindow : Window
    {
        private const string TemplateFilePath = "EmailTemplates/email_template.html";
        private List<string> File1Headers;
        private List<string> File2Headers;

        public EmailHTMLEditorWindow()
        {
            InitializeComponent();
            InitializeWebView();
        }
        private async void OnPlaceholderSelected(object sender, SelectionChangedEventArgs e)
        {
            var comboBox = sender as ComboBox;
            if (comboBox?.SelectedItem is string placeholder)
            {
                string js = $"tinymce.activeEditor.execCommand('mceInsertContent', false, '{{{{{placeholder}}}}}');";
                await WebEditor.ExecuteScriptAsync(js);
                comboBox.SelectedIndex = -1; // Reset selection
            }
        }

        public EmailHTMLEditorWindow(List<string> file1Headers, List<string> file2Headers)
        {
            this.InitializeComponent();
            this.File1Headers = file1Headers;
            this.File2Headers = file2Headers;
            InitializeWebView();
        }


        private async void InitializeWebView()
        {
            await WebEditor.EnsureCoreWebView2Async();

            WebEditor.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
            WebEditor.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = true;
            WebEditor.CoreWebView2.Settings.IsWebMessageEnabled = true;

            string resourcePath = Path.Combine(AppContext.BaseDirectory, "wwwroot", "tinymce", "js", "tinymce");
            WebEditor.CoreWebView2.SetVirtualHostNameToFolderMapping(
                "myfiles.local",
                resourcePath,
                CoreWebView2HostResourceAccessKind.Allow);

            string htmlContent = LoadHtmlTemplate();
            WebEditor.NavigateToString(htmlContent);
        }

        private string LoadHtmlTemplate()
        {
            string shellPath = "EmailTemplates/email_template_shell.html";
            string contentPath = "EmailTemplates/email_template_content.html";

            string contentHtml = File.Exists(contentPath)
                ? File.ReadAllText(contentPath, Encoding.UTF8)
                : "<p>Hello Team,</p>";  // fallback content

            string shellHtml = File.ReadAllText(shellPath, Encoding.UTF8);
            return shellHtml.Replace("{{CONTENT}}", contentHtml);
        }

        private void SetWindowSize(int width, int height)
        {
            IntPtr hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WindowId myWndId = Win32Interop.GetWindowIdFromWindow(hWnd);
            AppWindow appWindow = AppWindow.GetFromWindowId(myWndId);
            appWindow.Resize(new SizeInt32(width, height));
        }

        private async void SaveTemplate_Click(object sender, RoutedEventArgs e)
        {
            string htmlContent = await WebEditor.ExecuteScriptAsync("tinymce.get('editor').getContent()");
            htmlContent = System.Text.Json.JsonSerializer.Deserialize<string>($"\"{htmlContent.Trim('"')}\"");

            File.WriteAllText("EmailTemplates/email_template_content.html", htmlContent, Encoding.UTF8);
            this.Close();
        }
    }
}
