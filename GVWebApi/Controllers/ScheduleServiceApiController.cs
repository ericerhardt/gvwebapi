using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Web.Http.Description;
using GV.Domain;
using GVWebapi.Services;
using GVWebapi.RemoteData;
using GVWebapi.Models.Schedules;

namespace GVWebapi.Controllers
{
    public class ScheduleServiceApiController : ApiController
    {
        private readonly IScheduleServicesService _scheduleServicesService;
        private readonly IScheduleService _scheduleService;
        private readonly IUnitOfWork _unitOfWork;
        

        public ScheduleServiceApiController(IScheduleServicesService scheduleServicesService, IUnitOfWork unitOfWork, IScheduleService scheduleService)
        {
            _scheduleServicesService = scheduleServicesService;
            _scheduleService = scheduleService;
            _unitOfWork = unitOfWork;
        }
        [HttpGet, Route("api/services/costcenterservices/{customerId}")]
        public IHttpActionResult GetCostCenterServices(int CustomerId)
        {
            var costcenterServices = _scheduleService.GetCostCenterSchedule(CustomerId).ToList();
            var ccServicesSummary = _scheduleService.GetCCSummary(CustomerId).ToList();
            return Json(new {summary = ccServicesSummary,services = costcenterServices });
        }

        
        [HttpGet, Route("api/editschedule/services/get/{scheduleId}")]
        public IHttpActionResult GetScheduleServices(long scheduleId)
        {
            var metergroups = _scheduleServicesService.GetMeterGroups(scheduleId);
            var serviceadjustment = _scheduleServicesService.GetServiceAdjustments(scheduleId);
          
            return Json(new { MeterGroups = metergroups, ServiceAdjustments = serviceadjustment } );
        }

        [HttpPost, Route("api/editschedule/services/save")]
        public IHttpActionResult SaveScheduleService(ScheduleServiceSave model)
        {
            _scheduleServicesService.SaveSchedules(model.ScheduleServices,model.ScheduleId,model.ServiceAdjustment);
            _unitOfWork.Commit();
            return Ok();
        }

        [HttpPost,Route("api/editschedule/services/checkchanged")]
        public IHttpActionResult SetRemovedFromSchedule(ScheduleServiceCheckChangedModel model)
        {
            _scheduleServicesService.UpdateRemovedFromSchedule(model);
            _unitOfWork.Commit();
            return Ok();
        }
    }

    public class ScheduleServiceCheckChangedModel
    {
        public long ScheduleServiceId {get; set; }
        public bool RemovedFromSchedule { get; set; }
    }

    public class ScheduleServiceSaveModel
    {
        public long ScheduleServiceId { get; set; }
        public int ContractedPages { get; set; }
        public  decimal BaseCpp { get; set; }
        public decimal OverageCpp { get; set; }
    }
    public class ScheduleServiceSave
    {
        public IList<ScheduleServiceSaveModel> ScheduleServices { get; set; }
        public long ScheduleId { get; set; }
        public decimal ServiceAdjustment { get; set; }
    }
}