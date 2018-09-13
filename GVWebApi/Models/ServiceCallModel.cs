using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models
{
    public class ServiceCallModel
    {
        public int EquipmentID { get; set; }
        public string Name { get; set; }
        public string Description {get;set;}
        public int CallTypeID { get; set; }
        public string UserID { get; set; }
        public int Black { get; set; }
        public int Cyan { get; set; }
        public int Magenta { get; set; }
        public int Yellow { get; set; }
        public bool isWorking { get; set; }

    }
}