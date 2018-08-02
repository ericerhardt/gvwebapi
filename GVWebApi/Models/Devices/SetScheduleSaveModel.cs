using System.Collections.Generic;

namespace GVWebapi.Models.Devices
{
    public class SetScheduleSaveModel
    {
        public long ScheduleId { get; set; }
        public IList<long> DeviceIds { get; set; } = new List<long>();
    }
}