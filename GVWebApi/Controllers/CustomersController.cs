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

        
        [HttpGet, Route("api/customers/{id}")]
        [ResponseType(typeof(v_ARCustomers))]
        public IHttpActionResult Getv_ARCustomers(int id)
        {
            var vArCustomers = _coFreedomEntities.v_ARCustomers.Where(x => x.LocationID == id && x.Active == true)
                    .Select(customer => new Customer
                    {
                                                    CustomerID = customer.CustomerID,
                                                    CustomerNumber = customer.CustomerNumber,
                                                    CustomerName = customer.CustomerNumber,
                                                    Address = customer.Address,
                                                    City = customer.City,
                                                    State = customer.State,
                                                    County = customer.Country,
                                                    Zip = customer.Zip,
                                                    Country = customer.Country,
                                                    Longitude = customer.Longitude,
                                                    Latitude = customer.Latitude,
                                                    Phone1 = customer.Phone1,
                                                    Phone2 = customer.Phone2,
                                                    
                                                    Email = customer.Email,
                                                    WebSite = customer.WebSite
                                                })
                    .ToList();
            if (vArCustomers == null)
            {
                return NotFound();
            }
            foreach(var customer in vArCustomers)
            {
                customer.Address = customer.Address.Split(',')[0];

            }
            return Json(vArCustomers);
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
            var Locations = _coFreedomEntities.ARCustomers.Where(c => c.LocationID == LocationID && c.Active == true)
                                                .Select(x => new Customer
                                                {
                                                    CustomerID = x.CustomerID,
                                                    CustomerNumber = x.CustomerNumber,
                                                    CustomerName = x.CustomerNumber,
                                                    Address = x.Address,
                                                    City = x.City,
                                                    State = x.State,
                                                    County = x.Country,
                                                    Zip = x.Zip,
                                                    Country = x.Country,
                                                    Longitude = x.Longitude,
                                                    Latitude = x.Latitude,
                                                    Phone1 = x.Phone1,
                                                    Phone2 = x.Phone2,
                                                    Email = x.Email,
                                                    WebSite = x.WebSite,
                                                })
                                                .ToList();
            GoogleGeocodingAPI.GoogleAPIKey = "AIzaSyDoh1lOgTIiabZX7R0PPx_iq353YRfER2c";
            

            foreach (var Location in Locations)
            {
                string Address = string.Empty;
                 if (Location.Address.Contains(","))
                {
                    Address = $"{Location.Address.Split(',')[0]}, {Location.City} {Location.State} {Location.Zip}";
                } else
                {
                    Address = $"{Location.Address}, {Location.City} {Location.State} {Location.Zip}";
                }
                  

                AddressResponse result =  await  GoogleGeocodingAPI.SearchAddressAsync(Address);

                if (result.Results.Count  > 0)
                {
                   Location.Latitude = (decimal)result.Results.First().Geometry.Location.Lat;
                    Location.Longitude = (decimal)result.Results.First().Geometry.Location.Lng;
                    _coFreedomEntities.SaveChanges();
                }
              

            }


            
            return Json(Locations);
        }
    }
}