using System.Web.Http;
using GV.Domain;
using GVWebapi.Services;

namespace GVWebapi.Controllers
{
    public class LocationsApiController : ApiController
    {
        private readonly ILocationsService _locationsService;
        private readonly IUnitOfWork _unitOfWork;

        public LocationsApiController(ILocationsService locationsService, IUnitOfWork unitOfWork)
        {
            _locationsService = locationsService;
            _unitOfWork = unitOfWork;
        }

        [HttpGet,Route("api/adminlocations/{customerId}")]
        public IHttpActionResult GetAll(long customerId)
        {
            var locations = _locationsService.LoadAll(customerId);
            _unitOfWork.Commit();
            return Ok(locations);
        }

        [HttpGet, Route("api/adminlocation/setcorporate/{locationId}/{newValue}")]
        public IHttpActionResult UpdateCorporate(long locationId, bool newValue)
        {
            _locationsService.SetCorporate(locationId, newValue);
            _unitOfWork.Commit();
            return Ok();
        }

        [HttpPost,Route("api/adminlocation/updatetaxrate")]
        public IHttpActionResult UpdateTaxRate(LocationTaxRateSaveModel model)
        {
            _locationsService.UpdateTaxRate(model.LocationId, model.TaxRate);
            _unitOfWork.Commit();
            return Ok();
        }

        [HttpPost,Route("api/adminlocation/updateallrates")]
        public IHttpActionResult UpdateAllTaxRate(LocationAllTaxRateUpdateModel model)
        {
            _locationsService.UpdateAllRates(model.TaxRate);
            _unitOfWork.Commit();
            return Ok();
        }
    }

    public class LocationAllTaxRateUpdateModel
    {
        public decimal TaxRate { get; set; }
    }

    public class LocationTaxRateSaveModel
    {
        public long LocationId { get; set; }
        public decimal TaxRate { get; set; }
    }
}