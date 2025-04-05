using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailGenerator.Models.Settings
{
    public class AppSettings
    {
        public Dictionary<string, string> FieldMappings { get; set; } = new();
        public Dictionary<string, string> ConceptMappings { get; set; } = new();
        public FedExSettings FedEx { get; set; } = new();
    }

}
