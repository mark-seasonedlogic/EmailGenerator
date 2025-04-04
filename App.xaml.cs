using System.IO;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.UI.Xaml;
using EmailGenerator.Interfaces;
using EmailGenerator.Services;
using EmailGenerator.Models.Settings;
using OutlookDeviceEmailer;

namespace EmailGenerator
{
    public partial class App : Application
    {
        public static IHost Host { get; private set; }

        public App()
        {
            this.InitializeComponent();

            Host = Microsoft.Extensions.Hosting.Host
                .CreateDefaultBuilder()
                .ConfigureAppConfiguration((context, config) =>
                {
                    config.SetBasePath(Directory.GetCurrentDirectory());
                    config.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
                })
                .ConfigureServices((context, services) =>
                {
                    services.Configure<FedExSettings>(context.Configuration.GetSection("FedEx"));

                    // Configure HttpClient for FedExAuthProvider with BaseAddress
                    services.AddHttpClient<IFedExAuthProvider, FedExAuthProvider>((sp, client) =>
                    {
                        var settings = sp.GetRequiredService<Microsoft.Extensions.Options.IOptions<FedExSettings>>().Value;
                        client.BaseAddress = new System.Uri(settings.ApiBaseUrl);
                    });

                    // FedExShippingService just uses IFedExAuthProvider, so no special setup needed
                    services.AddHttpClient<IFedExShippingService, FedExShippingService>();

                    services.AddSingleton<MainWindow>();
                })
                .Build();
        }

        protected override void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
        {
            var mainWindow = Host.Services.GetRequiredService<MainWindow>();
            mainWindow.Activate();
        }
    }
}
