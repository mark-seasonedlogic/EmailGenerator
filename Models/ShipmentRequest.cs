using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailGenerator.Models
{

        public class ShipmentRequest
        {
            public RequestedShipment requestedShipment { get; set; }
            public AccountNumber accountNumber { get; set; }
            public string mergeLabelDocOption { get; set; } = "LABELS_AND_DOCS";
            public string labelResponseOptions { get; set; } = "LABEL";
    }

        public class RequestedShipment
        {
            public Shipper shipper { get; set; }
            public List<Recipient> recipients { get; set; } = new List<Recipient>
        {
            new Recipient
            {
                contact = new Contact
                {
                    personName = "MTECH MOBILITY",
                    phoneNumber = "8448642463"
                },
                address = new Address
                {
                    streetLines = new List<string> { "15827 GUILD COURT" },
                    city = "JUPITER",
                    stateOrProvinceCode = "FL",
                    postalCode = "33478",
                    countryCode = "US"
                }
            }
        };
            public string serviceType { get; set; } = "STANDARD_OVERNIGHT";
            public string packagingType { get; set; } = "YOUR_PACKAGING";
            public string pickupType { get; set; } = "DROPOFF_AT_FEDEX_LOCATION";
            public ShippingChargesPayment shippingChargesPayment { get; set; }
            public LabelSpecification labelSpecification { get; set; }
            public List<RequestedPackageLineItem> requestedPackageLineItems { get; set; }
        }

        public class Shipper
        {
            public Contact contact { get; set; }
            public Address address { get; set; }
        }

        public class Recipient
        {
            public Contact contact { get; set; }
            public Address address { get; set; }
        }

        public class Contact
        {
            public string personName { get; set; }
            public string phoneNumber { get; set; }
        }

        public class Address
        {
            public List<string> streetLines { get; set; }
            public string city { get; set; }
            public string stateOrProvinceCode { get; set; }
            public string postalCode { get; set; }
            public string countryCode { get; set; }
        }

        public class ShippingChargesPayment
        {
            public string paymentType { get; set; } = "THIRD_PARTY";
            public Payor payor { get; set; }
        }

        public class Payor
        {
            public ResponsibleParty responsibleParty { get; set; }
            public Address address { get; set; }
        }

        public class ResponsibleParty
        {
            public AccountNumber accountNumber { get; set; }
        }

        public class AccountNumber
        {
            public string value { get; set; }
            public string key { get; set; }
        }

        public class LabelSpecification
        {
            // Add properties as needed
        }

        public class RequestedPackageLineItem
        {
            public Weight weight { get; set; }
        }

        public class Weight
        {
            public string units { get; set; } = "LB";
            public string value { get; set; }
        }
    }

