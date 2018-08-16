namespace GVWebapi.Models.Devices
{
    public class DeviceSaveModel
    {
        public long DeviceId { get; set; }
        public long ScheduleId { get; set; }
        public int LocationId { get; set; }
        public decimal MonthlyCost { get; set; }
        public string Exhibit { get; set; }
        public string User { get; set; }
        public string CostCenter { get; set; }
    }
}