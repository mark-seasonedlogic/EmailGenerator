using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.Web.WebView2.Core;
using System;
using Windows.UI.WebUI;

namespace OutlookDeviceEmailer
{
    public sealed partial class EmailHTMLEditorNativeWindow : Window
    {
        public EmailHTMLEditorNativeWindow()
        {
            this.InitializeComponent();
            webView.NavigationCompleted += WebView_NavigationCompleted;
            LoadEditor();
        }

        private async void LoadEditor()
        {
            string htmlContent = @"
                <html>
                <head>
                    <style>
                        body { font-family: Arial, sans-serif; padding: 10px; }
                        #editor { border: 1px solid #ccc; padding: 10px; min-height: 400px; outline: none; }
                    </style>
                </head>
                <body>
                    <div id='editor' contenteditable='true'>Type your email here...</div>
                </body>
                </html>"
            ;

            if (webView.CoreWebView2 == null)
            {
                await webView.EnsureCoreWebView2Async();
            }
            webView.NavigateToString(htmlContent);
        }

        private void WebView_NavigationCompleted(WebView2 sender, CoreWebView2NavigationCompletedEventArgs args)
        {
            // JavaScript execution can be done here after the page loads
        }

        private void FormatText(object sender, RoutedEventArgs e)
        {
            if (sender is Button button)
            {
                string command = button.Tag.ToString();
                webView.CoreWebView2.ExecuteScriptAsync($"document.execCommand('{command}', false, null);");
            }
        }

        private async void InsertLink(object sender, RoutedEventArgs e)
        {
            string url = "https://example.com"; // You could show a dialog to get input
            await webView.CoreWebView2.ExecuteScriptAsync($"document.execCommand('createLink', false, '{url}');");
        }
    }
}
