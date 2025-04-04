using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailGenerator.Models.Settings
{
    public class FedExSettings
    {
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
        public string ApiBaseUrl { get; set; }
    }

}
