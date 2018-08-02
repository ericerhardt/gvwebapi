namespace GVWebapi.Models.Devices
{
    public class DeviceSearchModel
    {
        public string SerialNumber { get; set; }
        public int EquipmentId { get; set; }
        public string Model { get; set; }
        public string Location { get; set; }
        public string Status { get; set; }
        public string EquipmentNumber { get; set; }
    }
}