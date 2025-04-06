using EmailGenerator.ViewModels;
using Microsoft.UI.Xaml.Controls;
using EmailGenerator.Models.Settings;
using EmailGenerator.Helpers;
using Microsoft.UI.Xaml;


namespace EmailGenerator.Views
{
    public sealed partial class SettingsEditorView : Page
    {
        public SettingsEditorViewModel ViewModel { get; }

        public SettingsEditorView()
        {
            this.InitializeComponent();

            var settings = AppSettingsLoader.LoadFromFile("appsettings.json");
            ViewModel = new SettingsEditorViewModel(settings);
        }
        private void OnSaveClick(object sender, RoutedEventArgs e)
        {
            ViewModel.Save("appsettings.json");
        }
    }
}
