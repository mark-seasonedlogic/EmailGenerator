using EmailGenerator.Models;
using System.Threading.Tasks;

namespace EmailGenerator.Interfaces
{
    public interface IFedExShippingService
    {
        Task<byte[]> CreateShipmentLabelAsync(ShipmentRequest request);
    }

}
