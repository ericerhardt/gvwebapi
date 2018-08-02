using System.Web.Http;
using GV.Domain;
using GVWebapi.Services;

namespace GVWebapi.Controllers
{
    public class EditScheduleApiController : ApiController
    {
        private readonly IEditScheduleService _editScheduleService;
        private readonly IUnitOfWork _unitOfWork;
        private readonly ICoFreedomDeviceService _coFreedomDeviceService;

        public EditScheduleApiController(IEditScheduleService editScheduleService, IUnitOfWork unitOfWork, ICoFreedomDeviceService coFreedomDeviceService)
        {
            _editScheduleService = editScheduleService;
            _unitOfWork = unitOfWork;
            _coFreedomDeviceService = coFreedomDeviceService;
        }

        [HttpGet,Route("api/editschedule/gettop/{scheduleId}")]
        public IHttpActionResult GetTopSchedule(long scheduleId)
        {
            var topModel = _editScheduleService.GetScheduleTopModel(scheduleId);
            _unitOfWork.Commit();
            return Ok(topModel);
        }

        [HttpGet,Route("api/editschedule/device/gettop/{scheduleId}")]
        public IHttpActionResult GetDeviceTopModel(long scheduleId)
        {
            var deviceModel = _coFreedomDeviceService.GetTabCounts(scheduleId);
            _unitOfWork.Commit();
            return Ok(deviceModel);
        }
    }
}