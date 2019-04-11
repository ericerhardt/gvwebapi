using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models.Devices
{
    public class EquipmentPropertiesModel
    {
        public Nullable<int> Id { get; set; }
        public Nullable<int> Property { get; set;}
        public string Value { get; set; }
    }
}