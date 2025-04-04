using EmailGenerator.Interfaces;
using EmailGenerator.Models;
using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using EmailGenerator.Models.Settings;
using Microsoft.Extensions.Options;
using MsgKit;
using System.Reflection.Emit;
using System.Xml.Linq;

namespace EmailGenerator.Services
{
    public class FedExShippingService : IFedExShippingService
    {
        private readonly IFedExAuthProvider _authProvider;
        private readonly HttpClient _httpClient;
        private readonly FedExSettings _settings;

        public FedExShippingService(
            IFedExAuthProvider authProvider,
            HttpClient httpClient,
            IOptions<FedExSettings> options)
        {
            _authProvider = authProvider;
            _httpClient = httpClient;
            _settings = options.Value;
        }

        public async Task<byte[]> CreateShipmentLabelAsync(ShipmentRequest request)
        {
            var token = await _authProvider.GetAccessTokenAsync();
            var requestUrl = $"{_settings.ApiBaseUrl}/ship/v1/shipments";

            var httpRequest = new HttpRequestMessage(HttpMethod.Post, requestUrl);
            httpRequest.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
            httpRequest.Content = new StringContent(JsonConvert.SerializeObject(request), Encoding.UTF8, "application/json");
            //httpRequest.Content = new StringContent("{\"mergeLabelDocOption\":\"LABELS_AND_DOCS\",\"labelResponseOptions\":\"LABEL\",\"requestedShipment\":{\"shipper\":{\"contact\":{\"personName\":\"SENDER NAME\",\"phoneNumber\":\"9018328595\"},\"address\":{\"streetLines\":[\"SENDER ADDRESS 1\"],\"city\":\"MEMPHIS\",\"stateOrProvinceCode\":\"TN\",\"postalCode\":\"38116\",\"countryCode\":\"US\"}},\"recipients\":[{\"contact\":{\"personName\":\"RECIPIENT NAME\",\"phoneNumber\":\"9018328595\"}},{\"address\":{\"streetLines\":[\"RECIPIENT ADDRESS 1\"],\"city\":\"MEMPHIS\",\"stateOrProvinceCode\":\"TN\",\"postalCode\":\"38116\",\"countryCode\":\"US\"}}],\"serviceType\":\"STANDARD_OVERNIGHT\",\"packagingType\":\"YOUR_PACKAGING\",\"pickupType\":\"DROPOFF_AT_FEDEX_LOCATION\",\"shippingChargesPayment\":{\"paymentType\":\"SENDER\",\"payor\":{\"responsibleParty\":{\"accountNumber\":{\"value\":\"XXXX\",\"key\":\"\"}},\"address\":{\"streetLines\":[\"SENDER ADDRESS 1\"],\"city\":\"MEMPHIS\",\"stateOrProvinceCode\":\"TN\",\"postalCode\":\"38116\",\"countryCode\":\"US\"}}},\"labelSpecification\":{},\"requestedPackageLineItems\":[{\"weight\":{\"units\":\"LB\",\"value\":\"20\"}}]},\"accountNumber\":{\"value\":\"XXXXX2842\"}}", Encoding.UTF8, "application/json");

            var response = await _httpClient.SendAsync(httpRequest);
            response.EnsureSuccessStatusCode();
            var content = await response.Content.ReadAsByteArrayAsync();
            return content;
        }
    }

}
