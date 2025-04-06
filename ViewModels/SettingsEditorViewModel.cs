using CommunityToolkit.Mvvm.ComponentModel;
using EmailGenerator.Helpers;
using EmailGenerator.Models;
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
        private readonly AppSettings _originalSettings;
        public ObservableCollection<MappedField> FieldMappings { get; set; }
        public ObservableCollection<MappedField> ConceptMappings { get; set; }
        public FedExSettings FedEx => _settings.FedEx;

        public SettingsEditorViewModel(AppSettings settings)
        {
            _originalSettings = settings;
            _settings = settings;

            // Convert dictionaries to observable collections of editable MappedField objects
            FieldMappings = new ObservableCollection<MappedField>(
                settings.FieldMappings.Select(kv => new MappedField(kv.Key, kv.Value)));

            ConceptMappings = new ObservableCollection<MappedField>(
                settings.ConceptMappings.Select(kv => new MappedField(kv.Key, kv.Value)));

            // Attach change tracking to item-level PropertyChanged for IsDirty or Save control
            foreach (var item in FieldMappings)
                item.PropertyChanged += (s, e) => OnPropertyChanged(nameof(IsDirty));
            foreach (var item in ConceptMappings)
                item.PropertyChanged += (s, e) => OnPropertyChanged(nameof(IsDirty));

            FieldMappings.CollectionChanged += (s, e) => OnPropertyChanged(nameof(IsDirty));
            ConceptMappings.CollectionChanged += (s, e) => OnPropertyChanged(nameof(IsDirty));

            // Store a deep copy of the original settings for IsDirty comparison
            _originalSettings = new AppSettings
            {
                FieldMappings = settings.FieldMappings.ToDictionary(entry => entry.Key, entry => entry.Value),
                ConceptMappings = settings.ConceptMappings.ToDictionary(entry => entry.Key, entry => entry.Value),
                FedEx = new FedExSettings
                {
                    ClientId = settings.FedEx.ClientId,
                    ClientSecret = settings.FedEx.ClientSecret,
                    AccountNumber = settings.FedEx.AccountNumber,
                    MeterNumber = settings.FedEx.MeterNumber,
                    ApiBaseUrl = settings.FedEx.ApiBaseUrl
                }
            };
        }
        public void Save(string path)
        {
            var updatedSettings = GetUpdatedSettings();
            AppSettingsLoader.SaveToFile(path, updatedSettings);

            // Update original snapshot after saving
            _originalSettings.FieldMappings = updatedSettings.FieldMappings.ToDictionary(kv => kv.Key, kv => kv.Value);
            _originalSettings.ConceptMappings = updatedSettings.ConceptMappings.ToDictionary(kv => kv.Key, kv => kv.Value);
            _originalSettings.FedEx.ClientId = updatedSettings.FedEx.ClientId;
            _originalSettings.FedEx.ClientSecret = updatedSettings.FedEx.ClientSecret;
            _originalSettings.FedEx.AccountNumber = updatedSettings.FedEx.AccountNumber;
            _originalSettings.FedEx.MeterNumber = updatedSettings.FedEx.MeterNumber;
            _originalSettings.FedEx.ApiBaseUrl = updatedSettings.FedEx.ApiBaseUrl;

            OnPropertyChanged(nameof(IsDirty));
        }

        public bool IsDirty =>
        !FieldMappings.SequenceEqual(_originalSettings.FieldMappings.Select(kv => new MappedField(kv.Key, kv.Value))) ||
        !ConceptMappings.SequenceEqual(_originalSettings.ConceptMappings.Select(kv => new MappedField(kv.Key, kv.Value))) ||
        FedEx.ClientId != _originalSettings.FedEx.ClientId ||
        FedEx.ClientSecret != _originalSettings.FedEx.ClientSecret ||
        FedEx.AccountNumber != _originalSettings.FedEx.AccountNumber ||
        FedEx.MeterNumber != _originalSettings.FedEx.MeterNumber ||
        FedEx.ApiBaseUrl != _originalSettings.FedEx.ApiBaseUrl;

        public AppSettings GetUpdatedSettings()
        {
            return new AppSettings
            {
                FieldMappings = FieldMappings.ToDictionary(m => m.Key, m => m.Value),
                ConceptMappings = ConceptMappings.ToDictionary(m => m.Key, m => m.Value),
                FedEx = FedEx
            };
        }
    }
}
