using CommunityToolkit.Mvvm.ComponentModel;
using EmailGenerator.Models.Settings;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailGenerator.ViewModels
{
    public class SettingsEditorViewModel : ObservableObject
    {
        private AppSettings _settings;

        public ObservableCollection<KeyValuePair<string, string>> FieldMappings { get; set; }
        public ObservableCollection<KeyValuePair<string, string>> ConceptMappings { get; set; }
        public FedExSettings FedEx => _settings.FedEx;

        public SettingsEditorViewModel(AppSettings settings)
        {
            _settings = settings;
            FieldMappings = new(settings.FieldMappings);
            ConceptMappings = new(settings.ConceptMappings);
        }

        public AppSettings GetUpdatedSettings()
        {
            return new AppSettings
            {
                FieldMappings = FieldMappings.ToDictionary(kv => kv.Key, kv => kv.Value),
                ConceptMappings = ConceptMappings.ToDictionary(kv => kv.Key, kv => kv.Value),
                FedEx = FedEx
            };
        }
    }
}
