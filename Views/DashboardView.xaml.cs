using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml;
using Microsoft.Web.WebView2.Core;
using Microsoft.Extensions.Configuration;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using EmailGenerator.Models;
using Microsoft.Extensions.DependencyInjection;
using System.Text;
using System.IO;

namespace EmailGenerator.Views
{
    public sealed partial class DashboardView : Page
    {
        public DashboardViewModel ViewModel { get; }

        public DashboardView()
        {
            this.InitializeComponent();
            ViewModel = App.Host.Services.GetRequiredService<DashboardViewModel>();
            this.DataContext = ViewModel;

            this.Loaded += async (s, e) =>
            {
                await EmailEditor.EnsureCoreWebView2Async();

                EmailEditor.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
                EmailEditor.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = true;
                EmailEditor.CoreWebView2.Settings.IsWebMessageEnabled = true;

                string resourcePath = Path.Combine(AppContext.BaseDirectory, "wwwroot", "tinymce", "js", "tinymce");
                EmailEditor.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "myfiles.local",
                    resourcePath,
                    CoreWebView2HostResourceAccessKind.Allow);

                string htmlContent = LoadHtmlTemplate();
                EmailEditor.NavigateToString(htmlContent);
            };
        }

        private string LoadHtmlTemplate()
        {
            string shellPath = Path.Combine(AppContext.BaseDirectory, "EmailTemplates", "email_template_shell.html");
            string contentPath = Path.Combine(AppContext.BaseDirectory, "EmailTemplates", "email_template_content.html");

            string contentHtml = File.Exists(contentPath)
                ? File.ReadAllText(contentPath, Encoding.UTF8)
                : "<p>Hello Team,</p>";

            string shellHtml = File.Exists(shellPath)
                ? File.ReadAllText(shellPath, Encoding.UTF8)
                : "<html><body>{{CONTENT}}</body></html>";

            return shellHtml.Replace("{{CONTENT}}", contentHtml);
        }
        private void OnSaveClicked(object sender, RoutedEventArgs e)
        {
            ViewModel.Save();
        }

    }
}