using System.Web.Http;
using System.Collections.Generic;
using System.Linq;
using GV.Domain;
using GVWebapi.Services;
using GVWebapi.RemoteData;
using GVWebapi.Models.CostAllocation;
using GVWebapi.Models.Schedules;

namespace GVWebapi.Controllers
{
    public class CycleHistoryApiController : ApiController
    {
        private readonly ICycleHistoryService _cycleHistoryService;
        private readonly ICyclePeriodService _cyclePeriodService;
        private readonly CostAllocationService _costAllocationService;
        private readonly IUnitOfWork _unitOfWork;

        public CycleHistoryApiController(ICycleHistoryService cycleHistoryService, IUnitOfWork unitOfWork, ICyclePeriodService cyclePeriodService, CostAllocationService costAllocationService)
        {
            _cycleHistoryService = cycleHistoryService;
            _unitOfWork = unitOfWork;
            _cyclePeriodService = cyclePeriodService;
            _costAllocationService = costAllocationService;
        }

        [HttpGet, Route("api/cyclehistory/availablecycles/{customerId}")]
        public IHttpActionResult GetAvailablesCycles(long customerId)
        {
            return Ok(_cycleHistoryService.GetAvailableCycles(customerId));
        }

        [HttpPost, Route("api/cyclehistory/new")]
        public IHttpActionResult SaveNewCycle(NewCycleModel model)
        {
            _cycleHistoryService.AddNewCycle(model);
            _unitOfWork.Commit();
            return Ok();
        }

        [HttpPost, Route("api/cyclehistory/edit")]
        public IHttpActionResult UpdateCycle(NewCycleModel model)
        {
            _cycleHistoryService.UpdateCycle(model);
            
            return Ok();
        }

        [HttpGet, Route("api/cyclehistory/getall/{customerId}")]
        public IHttpActionResult GetCurrentCycles(long customerId)
        {
            return Ok(_cycleHistoryService.GetActiveCycles(customerId));
        }

        [HttpPost, Route("api/cyclehistory/toggle/")]
        public IHttpActionResult ToggleAvailability(ToggleSaveModel model)
        {
            _cycleHistoryService.ToggleAvailability(model);
            _unitOfWork.Commit();
            return Ok();
        }

        [HttpGet, Route("api/cyclehistory/period/add/{cycleId}")]
        public IHttpActionResult AddPeriod(long cycleId)
        {
            var endDate = _cyclePeriodService.AddPeriodToCycle(cycleId);
            _unitOfWork.Commit();
            var periods = _cyclePeriodService.GetCyclePeriods(cycleId);
            var model = new
            {
                EndDate = endDate,
                Periods = periods
            };
            return Ok(model);
        }

        [HttpGet, Route("api/cyclehistory/period/delete/{cyclePeriodId}")]
        public IHttpActionResult DeletePeriod(long cyclePeriodId)
        {
            var endDate = _cyclePeriodService.RemovePeriod(cyclePeriodId);
            _unitOfWork.Commit();
            var periods = _cyclePeriodService.GetCyclePeriodsByPeriodId(cyclePeriodId);
            var model = new
            {
                EndDate = endDate,
                Periods = periods
            };
            return Ok(model);
        }

        [HttpGet, Route("api/cyclehistory/delete/{cycleId}")]
        public IHttpActionResult DeleteCycle(long cycleId)
        {
            _cycleHistoryService.DeleteCycle(cycleId);
            _unitOfWork.Commit();
            return Ok();
        }

        [HttpGet, Route("api/cyclehistory/period/summary/{custid}/{cyclePeriodId}")]
        public IHttpActionResult GetCyclePeriodSummary(int custid, long cyclePeriodId)
        {
            
            var model = _cyclePeriodService.GetCyclePeriodSummary(cyclePeriodId);
            _unitOfWork.Commit();

            return Ok( new { model , allocatedService = model.AllocatedServices, costCenterSummaries = model.CostcenterSummary, model.Metergroups });
        }

        [HttpGet, Route("api/cyclehistory/period/devices/{cyclePeriodId}")]
        public IHttpActionResult GetCyclePeriodDevices(long cyclePeriodId)
        {
            var model = _cyclePeriodService.RefreshCyclePeriodDevices(cyclePeriodId);
            _unitOfWork.Commit();
            return Ok(model);
        }

        [HttpPost, Route("api/cyclehistory/period/instancesinvoiced")]
        public IHttpActionResult UpdateScheduleInstancesInvoiced(InvoiceInstanceSaveModel imodel)
        {
            _cyclePeriodService.SaveInstancesInvoiced(imodel);
            _unitOfWork.Commit();
  
            var model = _cyclePeriodService.GetCyclePeriodSummary(imodel.CyclePeriodId);
            return Ok(new { model, allocatedService = model.AllocatedServices, costCenterSummaries = model.CostcenterSummary, model.Metergroups });
        }

        [HttpPost, Route("api/cyclehistory/period/updateinvoice")]
        public IHttpActionResult UpdateInvoiceNumber(InvoiceNumberSaveModel model)
        {
            _cyclePeriodService.UpdateInvoiceNumber(model);
            _unitOfWork.Commit();
            return Ok();
        }

        [HttpPost,Route("api/cyclehistory/updatereconcile")]
        public IHttpActionResult UpdateIsReconciled(CycleReconcileSaveModel model)
        {
            _cycleHistoryService.ToggleReconcile(model);
            _unitOfWork.Commit();
            return Ok();
        }
    }
}