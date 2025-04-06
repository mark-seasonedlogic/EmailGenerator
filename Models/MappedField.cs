using CommunityToolkit.Mvvm.ComponentModel;

namespace EmailGenerator.Models
{
    /// <summary>
    /// Represents a mutable key-value pair used for editable mappings in the UI.
    /// Unlike KeyValuePair, this class supports property change notifications.
    /// </summary>
    /// <remarks>
    /// This is necessary because KeyValuePair is a struct and immutable — editing
    /// the Key or Value won't trigger UI updates. By making the mapping editable
    /// and observable, we can support live change tracking (e.g., IsDirty updates)
    /// and enable binding in XAML.
    /// </remarks>
    public class MappedField : ObservableObject
    {
        private string _key;
        private string _value;

        /// <summary>
        /// Gets or sets the token used in templates (e.g., RSTRNT_LEGAL_NAME).
        /// </summary>
        public string Key
        {
            get => _key;
            set => SetProperty(ref _key, value);
        }

        /// <summary>
        /// Gets or sets the actual CSV column name this token maps to (e.g., Legal Name).
        /// </summary>
        public string Value
        {
            get => _value;
            set => SetProperty(ref _value, value);
        }

        /// <summary>
        /// Initializes a new instance of the MappedField class with a key and value.
        /// </summary>
        public MappedField(string key, string value)
        {
            _key = key;
            _value = value;
        }
    }
}
