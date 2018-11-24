using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Web.Http.Description;
 using GVWebapi.RemoteData;
using GVWebapi.Models;
using GoogleMaps.LocationServices;
using GuigleAPI;
using GuigleAPI.Model;
using System.Threading.Tasks;

namespace GVWebapi.Controllers
{
    public class CustomersController : ApiController
    {
        private readonly CoFreedomEntities _coFreedomEntities = new CoFreedomEntities();

        public IQueryable<v_ARCustomers> Getv_ARCustomers()
        {
            return _coFreedomEntities.v_ARCustomers;
        }

        [ResponseType(typeof(v_ARCustomers))]
        public IHttpActionResult Getv_ARCustomers(int id)
        {
            var vArCustomers = _coFreedomEntities.v_ARCustomers.Find(id);
            if (vArCustomers == null)
            {
                return NotFound();
            }

            return Ok(vArCustomers);
        }
         
        [HttpGet,Route("api/locations/{customerID}")]
        public IHttpActionResult GetLocations(int customerId)
        {
            var i = 0;
            var locations = (from c in _coFreedomEntities.vw_admin_EquipmentList_MeterGroup
                             where c.Active && c.CustomerID == customerId
                             select new
                             {
                                 value = c.LocName,
                                 text = c.LocName,
                             }).Distinct().ToList();

            var equipmentLocations = new List<EquipmentLocation>();
            foreach (var location in locations)
            {
                equipmentLocations.Add(new EquipmentLocation { text = location.text, value = i++ });
            }

            return Ok(equipmentLocations);

        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _coFreedomEntities.Dispose();
            }
            base.Dispose(disposing);
        }

        [System.Web.Http.Route("api/SetCustomerLongLat/{LocationID}")]
        [System.Web.Http.HttpGet]
        public async Task<IHttpActionResult> SetCustomerLongLat(int LocationID)
        {
            IEnumerable<ARCustomer> Locations = _coFreedomEntities.ARCustomers.Where(c => c.LocationID == LocationID).ToList();
            GoogleGeocodingAPI.GoogleAPIKey = "AIzaSyDoh1lOgTIiabZX7R0PPx_iq353YRfER2c";
            

            foreach (var Location in Locations)
            {
                string Address = $"{Location.Address}, {Location.City} {Location.State} {Location.Zip}";
                AddressResponse result =  await  GoogleGeocodingAPI.SearchAddressAsync(Address);

                if (result.Results.Count  > 0)
                {
                    Location.Latitude = (decimal)result.Results.First().Geometry.Location.Lat;
                    Location.Longitude = (decimal)result.Results.First().Geometry.Location.Lng;
                    _coFreedomEntities.SaveChanges();
                }
              

            }


            // Save lat/long values to DB...
            return  Ok();
        }
    }
}