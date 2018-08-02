namespace GVWebapi.Models.Devices
{
    public class DeviceReplacementSaveModel
    {
        public long OldDeviceId { get; set; }
        public string Location { get; set; }
        public string NewSerialNumber { get; set; }
        public string NewModel { get; set; }
        public decimal ReplacementValue { get; set; }
        public string ScheduleNumber { get; set; }
    }
}