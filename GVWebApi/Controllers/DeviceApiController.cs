using System.Web.Http;
using GV.CoFreedomDomain;
using GV.Domain;
using GV.Services;
using GVWebapi.Models.Devices;
using GVWebapi.Models.Schedules;
using GVWebapi.Services;

namespace GVWebapi.Controllers
{
    public class DeviceApiController : ApiController
    {
        private readonly IDeviceService _deviceService;
        private readonly IUnitOfWork _unitOfWork;
        private readonly ICoFreedomUnitOfWork _coFreedomUnitOfWork;
        private readonly IScheduleService _scheduleService;
        private readonly ILocationsService _locationSerivce;

        public DeviceApiController(IDeviceService deviceService, IUnitOfWork unitOfWork, ICoFreedomUnitOfWork coFreedomUnitOfWork, IScheduleService scheduleService, ILocationsService locationSerivce)
        {
            _deviceService = deviceService;
            _unitOfWork = unitOfWork;
            _coFreedomUnitOfWork = coFreedomUnitOfWork;
            _scheduleService = scheduleService;
            _locationSerivce = locationSerivce;
        }

        [HttpGet, Route("api/editschedule/devices/active/{scheduleId}")]
        public IHttpActionResult GetActiveDevices(long scheduleId)
        {
            return Json(_deviceService.GetActiveDevices(scheduleId));
        }

        [HttpGet, Route("api/editschedule/devices/active/delete/{deviceId}")]
        public IHttpActionResult DeleteActiveDevice(long deviceId)
        {
            _deviceService.DeleteDevice(deviceId);
            _unitOfWork.Commit();
            _coFreedomUnitOfWork.Commit();
            return Ok();
        }

        [HttpGet, Route("api/editschedule/devices/active/getdetails/{deviceId}")]
        public IHttpActionResult GetEditDetails(long deviceId)
        {
            var model = new
            {
                ActiveSchedules = _scheduleService.GetActiveSchedules(deviceId),
                Locations = _locationSerivce.LoadAllByDeviceId(deviceId),
                DeviceItem = _deviceService.GetDeviceByID(deviceId)
            };

            _unitOfWork.Commit();
            return Ok(model);
        }

        [HttpPost, Route("api/editschedule/devices/active/save")]
        public IHttpActionResult SaveDeviceDetails(DeviceSaveModel model)
        {
            _deviceService.SaveDevice(model);
            _unitOfWork.Commit();
            var TotalCost = _deviceService.DeviceTotalCost(model.ScheduleId);
            _scheduleService.UpdateMonthyHardwareCost(TotalCost, model.ScheduleId);
            _unitOfWork.Commit();
            _coFreedomUnitOfWork.Commit();
            return Ok();
        }

        [HttpGet, Route("api/editschedule/devices/unallocated/{scheduleId}")]
        public IHttpActionResult GetUnallocatedDevices(long scheduleId)
        {
            var model = new
            {
                Devices = _deviceService.GetUnallocatedDevices(scheduleId),
                Schedules = _scheduleService.GetAcitveSchedulesByScheduleId(scheduleId)
            };

            return Ok(model);
        }

        [HttpPost, Route("api/editschedule/devices/setschedulefordevice")]
        public IHttpActionResult AddDevicesToSchedule(SetScheduleSaveModel model)
        {
            _deviceService.AddDevicesToSchedule(model);
            _coFreedomUnitOfWork.Commit();
            _unitOfWork.Commit();
            return Ok();
        }

        [HttpPost, Route("api/editschedule/devices/confirmremove")]
        public IHttpActionResult ConfirmDeviceRemoval(DeviceRemoveModel model)
        {
            _deviceService.ConfirmDeviceRemove(model);
            _unitOfWork.Commit();
            return Ok();
        }

        [HttpPost, Route("api/editschedule/devices/confirmreplacement")]
        public IHttpActionResult ConfirmFormatterReplaced(FormatterReplacedModel model)
        {
            _deviceService.ConfirmFormatterReplacement(model);
            _unitOfWork.Commit();
            return Ok();
        }

        [HttpGet, Route("api/editschedule/devices/removed/{scheduleId}")]
        public IHttpActionResult GetRemovedDevices(long scheduleId)
        {
            return Ok(_deviceService.GetRemovedDevices(scheduleId));
        }

        [HttpGet, Route("api/editschedule/devices/replacementinfo/{scheduleId}")]
        public IHttpActionResult GetReplacementDeviceInfo(long scheduleId)
        {
            var model = new
            {
                Devices = _deviceService.GetDevicesToSearch(scheduleId),
                Schedules = _scheduleService.GetAcitveSchedulesByScheduleId(scheduleId)
            };
            return Ok(model);
        }

        [HttpPost,Route("api/editschedule/devices/replacementdevice/save")]
        public IHttpActionResult SaveReplacementDevice(DeviceReplacementSaveModel model)
        {
            _deviceService.AddReplacementDevice(model);
            _unitOfWork.Commit();
            return Ok();
        }
    }
}