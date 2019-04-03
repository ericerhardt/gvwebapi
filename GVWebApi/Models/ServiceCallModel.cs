using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GVWebapi.Models
{
    public class ServiceCallModel
    {
        [Key]
        public int EquipmentID { get; set; }
        public string EquipmentNumber { get; set; }
        public string Name { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public string Description {get;set;}
        public int CallTypeID { get; set; }
        public string UserID { get; set; }
        public int Black { get; set; }
        public int Cyan { get; set; }
        public int Magenta { get; set; }
        public int Yellow { get; set; }
        public bool isWorking { get; set; }
        public string CallID { get; set; }
        public string CallNumber { get; set; }
        public Nullable<int> CallType { get; set; }
    }
}