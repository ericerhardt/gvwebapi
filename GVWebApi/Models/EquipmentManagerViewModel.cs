using System.Collections.Generic;
using GVWebapi.RemoteData;

namespace GVWebapi.Models
{
    public class EquipmentManagerViewModel
    {
        public IEnumerable<EquipmentManagerList> equimpmentList { get; set; }
        public int draw { get; set; }
        public int totalRecords{ get; set; }
        public int filteredRecords { get; set; }
        public int inactive { get; set; }
        public int active { get; set; }
        public int incomplete { get; set; }
    }
}