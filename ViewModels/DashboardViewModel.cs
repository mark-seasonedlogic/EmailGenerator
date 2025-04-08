using Microsoft.Extensions.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Text;
using System;

namespace EmailGenerator.Views
{
    public class DashboardViewModel : INotifyPropertyChanged
    {
        private readonly IConfiguration _config;
        private string _subject;
        private string _to;
        private string _cc;
        private string _bcc;

        public string SubjectLine
        {
            get => _subject;
            set => SetProperty(ref _subject, value);
        }
        public string StaticToRecipients
        {
            get => _to;
            set => SetProperty(ref _to, value);
        }
        public string StaticCcRecipients
        {
            get => _cc;
            set => SetProperty(ref _cc, value);
        }
        public string StaticBccRecipients
        {
            get => _bcc;
            set => SetProperty(ref _bcc, value);
        }

        public string HtmlShellPath =>
    "file:///" + Path.Combine(AppContext.BaseDirectory, _config["Email:HtmlTemplateShellPath"]).Replace("\\", "/");

        public DashboardViewModel(IConfiguration config)
        {
            _config = config;
            _subject = config["Email:SubjectLine"] ?? "";
            _to = config["Email:StaticToRecipients"] ?? "";
            _cc = config["Email:StaticCcRecipients"] ?? "";
            _bcc = config["Email:StaticBccRecipients"] ?? "";
        }

        public void Save()
        {
            var path = Path.Combine(AppContext.BaseDirectory, "appsettings.json");
            var json = File.ReadAllText(path);
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement.Clone();

            using var stream = new MemoryStream();
            using var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true });

            writer.WriteStartObject();
            foreach (var prop in root.EnumerateObject())
            {
                if (prop.NameEquals("Email"))
                {
                    writer.WritePropertyName("Email");
                    writer.WriteStartObject();
                    writer.WriteString("SubjectLine", SubjectLine);
                    writer.WriteString("StaticToRecipients", StaticToRecipients);
                    writer.WriteString("StaticCcRecipients", StaticCcRecipients);
                    writer.WriteString("StaticBccRecipients", StaticBccRecipients);
                    writer.WriteString("HtmlTemplateShellPath", HtmlShellPath);
                    foreach (var subProp in prop.Value.EnumerateObject())
                    {
                        if (!subProp.NameEquals("SubjectLine") &&
                            !subProp.Name.StartsWith("Static") &&
                            !subProp.Name.Equals("HtmlTemplateShellPath"))
                        {
                            writer.WritePropertyName(subProp.Name);
                            subProp.WriteTo(writer);
                        }
                    }
                    writer.WriteEndObject();
                }
                else
                {
                    writer.WritePropertyName(prop.Name);
                    prop.Value.WriteTo(writer);
                }
            }
            writer.WriteEndObject();
            writer.Flush();
            File.WriteAllText(path, Encoding.UTF8.GetString(stream.ToArray()));
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void SetProperty<T>(ref T storage, T value, [CallerMemberName] string propertyName = null)
        {
            if (!EqualityComparer<T>.Default.Equals(storage, value))
            {
                storage = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}