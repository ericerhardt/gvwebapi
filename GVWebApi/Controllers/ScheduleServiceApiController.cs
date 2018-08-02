using System.Collections.Generic;
using System.Web.Http;
using GV.Domain;
using GVWebapi.Services;

namespace GVWebapi.Controllers
{
    public class ScheduleServiceApiController : ApiController
    {
        private readonly IScheduleServicesService _scheduleServicesService;
        private readonly IUnitOfWork _unitOfWork;

        public ScheduleServiceApiController(IScheduleServicesService scheduleServicesService, IUnitOfWork unitOfWork)
        {
            _scheduleServicesService = scheduleServicesService;
            _unitOfWork = unitOfWork;
        }

        [HttpGet, Route("api/editschedule/services/get/{scheduleId}")]
        public IHttpActionResult GetScheduleServices(long scheduleId)
        {
            return Ok(_scheduleServicesService.GetMeterGroups(scheduleId));
        }

        [HttpPost, Route("api/editschedule/services/save/")]
        public IHttpActionResult SaveScheduleService(IList<ScheduleServiceSaveModel>  model)
        {
            _scheduleServicesService.SaveSchedules(model);
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
}