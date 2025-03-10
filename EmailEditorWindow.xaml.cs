using Microsoft.UI.Xaml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using Microsoft.UI.Xaml.Controls;

namespace OutlookDeviceEmailer
{
    public sealed partial class EmailEditorWindow : Window
    {
        private const string TemplateFilePath = "email_template.txt"; // Saves email template
        private List<string> placeholders = new List<string>(); // Holds CSV column names

        public EmailEditorWindow(string deviceFile, string emailFile)
        {
            InitializeComponent();
            LoadTemplate();
            LoadColumnHeaders(deviceFile, emailFile);
        }

        private void LoadTemplate()
        {
            if (File.Exists(TemplateFilePath))
            {
                txtEmailTemplate.Text = File.ReadAllText(TemplateFilePath, Encoding.UTF8);
            }
        }

        private void LoadColumnHeaders(string deviceFile, string emailFile)
        {
            placeholders.Clear();

            if (!string.IsNullOrEmpty(deviceFile) && File.Exists(deviceFile))
                placeholders.AddRange(GetCsvHeaders(deviceFile));

            if (!string.IsNullOrEmpty(emailFile) && File.Exists(emailFile))
                placeholders.AddRange(GetCsvHeaders(emailFile));

            placeholders = placeholders.Distinct().OrderBy(p => p).ToList(); // Remove duplicates

            cmbPlaceholders.ItemsSource = placeholders; // Bind headers to dropdown
        }

        private List<string> GetCsvHeaders(string filePath)
        {
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
            {
                if (csv.Read())
                {
                    csv.ReadHeader();
                    return csv.HeaderRecord.ToList();
                }
            }
            return new List<string>();
        }

        private void InsertPlaceholder_Click(object sender, RoutedEventArgs e)
        {
            if (cmbPlaceholders.SelectedItem != null)
            {
                string placeholder = $"{{{cmbPlaceholders.SelectedItem.ToString().Trim()}}}"; // Ensure exact match
                int cursorPos = txtEmailTemplate.SelectionStart;
                txtEmailTemplate.Text = txtEmailTemplate.Text.Insert(cursorPos, placeholder);
                txtEmailTemplate.SelectionStart = cursorPos + placeholder.Length; // Move cursor
            }
        }

        private void SaveTemplate_Click(object sender, RoutedEventArgs e)
        {
            File.WriteAllText(TemplateFilePath, txtEmailTemplate.Text, Encoding.UTF8);
            this.Close();
        }
    }
}
