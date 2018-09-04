using System.Web.Http;
using GV.Domain;
using GVWebapi.Models.Schedules;
using GVWebapi.Services;

namespace GVWebapi.Controllers
{
    public class ScheduleApiController : ApiController
    {
        private readonly IScheduleService _scheduleService;
        private readonly IUnitOfWork _unitOfWork;

        public ScheduleApiController(IScheduleService scheduleService, IUnitOfWork unitOfWork)
        {
            _scheduleService = scheduleService;
            _unitOfWork = unitOfWork;
        }

        [HttpPost,Route("api/schedules/add")]
        public IHttpActionResult AddSchedule(ScheduleSaveModel model)
        {
            var hasSchedule = _scheduleService.ScheduleExists(model);
            if (hasSchedule)
            {
                return Content(System.Net.HttpStatusCode.BadRequest, "Duplicate Schedule");
            }
            else
            {
               _scheduleService.AddSchedule(model);
                _unitOfWork.Commit();
                return Ok();
            }
           
        }

        [HttpPost,Route("api/schedules/update")]
        public IHttpActionResult UpdateSchedule(ScheduleSaveModel model)
        {
            _scheduleService.UpdateSchedule(model);
            _unitOfWork.Commit();
            return Ok();
        }

        [HttpGet,Route("api/schedules/all/{customerId}")]
        public IHttpActionResult GetAll(long customerId)
        {
            return Ok(_scheduleService.GetAll(customerId));
        }

        [HttpGet,Route("api/schedules/delete/{scheduleId}")]
        public IHttpActionResult DeleteSchedule(long scheduleId)
        {
            _scheduleService.DeleteSchedule(scheduleId);
            _unitOfWork.Commit();
            return Ok();
        }

        [HttpGet,Route("api/schedules/coterminous/getall/{customerId}")]
        public IHttpActionResult GetCoterminous(long customerId)
        {
            return Ok(_scheduleService.GetCoterminous(customerId));
        }

        [HttpGet,Route("api/schedules/delete/candelete/{scheduleId}")]
        public IHttpActionResult CanDeleteSchedule(long scheduleId)
        {
            return Ok(_scheduleService.CanDeleteSchedule(scheduleId));
        }

        [HttpGet,Route("api/schedules/edit/{scheduleId}")]
        public IHttpActionResult GetEditSchedule(long scheduleId)
        {
            return Ok(_scheduleService.GetExistingSchedule(scheduleId));
        }

        [HttpGet,Route("api/schedules/get/{scheduleId}")]
        public IHttpActionResult GetSchedule(long scheduleId)
        {
            return Ok(_scheduleService.GetSchedule(scheduleId));
        }
    }
}