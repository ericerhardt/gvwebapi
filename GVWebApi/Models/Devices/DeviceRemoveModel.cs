using System;

namespace GVWebapi.Models.Devices
{
    public class DeviceRemoveModel
    {
        public long DeviceId { get; set; }
        public DateTime RemovalDate { get; set; }
    }
}