using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;

namespace EmailGenerator.Services
{
    /// <summary>
    /// Service responsible for loading and transforming CSV data into structured email lookup mappings.
    /// </summary>
    /// <remarks>
    /// Moving this logic out of the UI layer keeps EmailGeneratorView.xaml.cs clean and focused solely on view interaction.
    /// The service approach enables testability, modularity, and support for future data sources.
    /// </remarks>
    public class EmailDataService
    {
        /// <summary>
        /// Loads restaurant-level details (emails, contacts, addresses) from a CSV file.
        /// Returns a dictionary keyed by composite key (store# + concept).
        /// </summary>
        public Dictionary<string, Dictionary<string, string>> LoadEmailLookup(string emailCsvPath, Func<string, string> fieldResolver, Func<string, string> conceptMapper)
        {
            var lookup = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

            using var reader = new StreamReader(emailCsvPath);
            using var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture));

            if (!csv.Read() || !csv.ReadHeader())
                throw new InvalidOperationException("CSV file missing headers");

            var headers = csv.HeaderRecord.ToList();

            while (csv.Read())
            {
                string conceptCode = csv.GetField(fieldResolver("CONCEPT_CD"))?.Trim();
                string restaurantNumber = csv.GetField(fieldResolver("RSTRNT_NBR"))?.Trim();

                if (string.IsNullOrEmpty(conceptCode) || string.IsNullOrEmpty(restaurantNumber))
                    continue;

                string concept = conceptMapper(conceptCode);
                string paddedRestaurant = restaurantNumber.PadLeft(4, '0');
                string key = paddedRestaurant + concept;

                var dict = new Dictionary<string, string>();
                foreach (var h in headers)
                {
                    dict[h] = csv.GetField(h)?.Trim();
                }

                lookup[key] = dict;
            }

            return lookup;
        }

        /// <summary>
        /// Maps devices to their matching restaurant email based on structured identifiers from a second CSV.
        /// Returns a dictionary of recipient emails to base64-encoded JSON device data strings.
        /// </summary>
        public Dictionary<string, List<string>> GenerateEmailDictionary(
            string deviceCsvPath,
            Dictionary<string, Dictionary<string, string>> emailLookup,
            Func<string, string> fieldResolver)
        {
            var emailDict = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            using var reader = new StreamReader(deviceCsvPath);
            using var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture));

            if (!csv.Read() || !csv.ReadHeader())
                throw new InvalidOperationException("Device CSV missing headers");

            while (csv.Read())
            {
                string userName = csv.GetField("Username")?.Trim();
                if (string.IsNullOrEmpty(userName))
                    continue;

                string storeNumber = userName.Substring(3, 4);
                string concept = userName.Substring(0, 3);
                string lookupKey = storeNumber + concept;

                if (!emailLookup.ContainsKey(lookupKey))
                    continue;

                string recipientEmail = emailLookup[lookupKey][fieldResolver("STORE_EMAIL_ADDR")];

                if (!emailDict.ContainsKey(recipientEmail))
                    emailDict[recipientEmail] = new List<string>();

                var details = new Dictionary<string, string>();

                foreach (var kv in emailLookup[lookupKey])
                {
                    if (kv.Key != fieldResolver("STORE_EMAIL_ADDR"))
                        details[kv.Key] = kv.Value;
                }

                details["MD + Serial"] = csv.GetField("MD + Serial")?.Trim();
                details["Serial Number"] = csv.GetField("Serial Number")?.Trim();
                details["JVP EMAIL"] = csv.GetField("JVP EMAIL")?.Trim();
                details["MVP EMAIL"] = csv.GetField("MVP EMAIL")?.Trim();

                string json = JsonSerializer.Serialize(details);
                string encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));

                emailDict[recipientEmail].Add(encoded);
            }

            return emailDict;
        }
    }
}
