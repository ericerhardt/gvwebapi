using System.Web.Http;
using GV.CoFreedomDomain;
using GV.Domain;
using GV.Services;
using GVWebapi.Models.Devices;
using GVWebapi.Models.Schedules;
using GVWebapi.Services;
using GVWebapi.RemoteData;

namespace GVWebapi.Controllers
{
    public class DeviceApiController : ApiController
    {
        private readonly IScheduleDevicesService _deviceService;
        private readonly IUnitOfWork _unitOfWork;
        private readonly ICoFreedomUnitOfWork _coFreedomUnitOfWork;
        private readonly IScheduleService _scheduleService;
        private readonly ILocationsService _locationSerivce;
        private readonly GlobalViewEntities _globalViewEntities = new GlobalViewEntities();

        public DeviceApiController(IScheduleDevicesService deviceService, IUnitOfWork unitOfWork, ICoFreedomUnitOfWork coFreedomUnitOfWork, IScheduleService scheduleService, ILocationsService locationSerivce)
        {
            _deviceService = deviceService;
            _unitOfWork = unitOfWork;
            _coFreedomUnitOfWork = coFreedomUnitOfWork;
            _scheduleService = scheduleService;
            _locationSerivce = locationSerivce;
        }
 
        [HttpGet, Route("api/editschedule/devices/active/getdetails/{scheduleDeviceID}")]
        public IHttpActionResult GetEditDetails(long scheduleDeviceID)
        {
            var device = _globalViewEntities.ScheduleDevices.Find(scheduleDeviceID);
            var model = new
            {
                DeviceItem = _deviceService.GetDeviceByID(device.ScheduleDeviceID),
                ActiveSchedules = _scheduleService.GetActiveSchedules(device.CustomerID),
                Locations = _locationSerivce.LoadAllByDeviceId(device.EquipmentID),
               
            };

            _unitOfWork.Commit();
            return Ok(model);
        }

        [HttpPost, Route("api/editschedule/devices/active/save")]
        public IHttpActionResult SaveDeviceDetails(DeviceSaveModel model)
        {
            _deviceService.SaveDevice(model);
            _coFreedomUnitOfWork.Commit();
            var TotalCost = _deviceService.DeviceTotalCost(model.ScheduleId);
            _scheduleService.UpdateMonthyHardwareCost(TotalCost, model.ScheduleId);
            _unitOfWork.Commit();
           
            return Ok();
        }

        [HttpGet, Route("api/editschedule/devices/unallocated/{customerId}")]
        public IHttpActionResult GetUnallocatedDevices(long customerId)
        {
            var model = new
            {
                Devices = _deviceService.GetUnallocatedDevices(customerId),
                Schedules = _scheduleService.GetAcitveSchedulesByCustomer(customerId)
            };

            return Ok(model);
        }

        [HttpPost, Route("api/editschedule/devices/setschedulefordevice")]
        public IHttpActionResult AddDevicesToSchedule(SetScheduleSaveModel model)
        {
            _deviceService.AddDevicesToSchedule(model);
            _coFreedomUnitOfWork.Commit();

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