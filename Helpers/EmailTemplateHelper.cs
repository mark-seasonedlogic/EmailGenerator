using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailGenerator.Helpers;

public static class EmailTemplateHelper
{
    /// <summary>
    /// Replaces tokens in the form of {{TOKEN_NAME}} within the given template,
    /// using the FieldMappings dictionary to resolve actual CSV field names from restaurant/device data.
    /// </summary>
    public static string ApplyFieldMappingsToTemplate(
        string template,
        Dictionary<string, string> dataRow,
        Dictionary<string, string> fieldMappings)
    {
        foreach (var token in fieldMappings)
        {
            string placeholder = $"{{{{{token.Key}}}}}"; // e.g., {{STORE_EMAIL_ADDR}}
            if (dataRow.TryGetValue(token.Value, out var value))
            {
                template = template.Replace(placeholder, value);
            }
            else
            {
                template = template.Replace(placeholder, ""); // fallback to empty if not found
            }
        }

        return template;
    }

    /// <summary>
    /// Builds a device list HTML snippet (e.g., <ul><li>Device1</li><li>Device2</li></ul>)
    /// from a set of CSV-parsed device rows using field mappings.
    /// </summary>
    public static string GenerateDeviceListHtml(
        IEnumerable<Dictionary<string, string>> deviceRows,
        List<string> fieldsToDisplay,
        Dictionary<string, string> fieldMappings)
    {
        var builder = new System.Text.StringBuilder();
        builder.AppendLine("<ul>");

        foreach (var row in deviceRows)
        {
            var displayValues = new List<string>();

            foreach (var tokenKey in fieldsToDisplay)
            {
                if (fieldMappings.TryGetValue(tokenKey, out var actualField) &&
                    row.TryGetValue(actualField, out var value))
                {
                    displayValues.Add(value);
                }
            }

            builder.AppendLine($"<li>{string.Join(", ", displayValues)}</li>");
        }

        builder.AppendLine("</ul>");
        return builder.ToString();
    }
}
