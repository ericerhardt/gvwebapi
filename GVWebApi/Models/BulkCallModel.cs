using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models
{
    public class BulkCallModel
    {
    
       public string DeviceId { get; set; }
       public string CallId { get; set; }
       public int Blk { get; set; }
       public string Comment { get; set; }
       public string Contact { get; set; }
       public int Cyan { get; set; }
       public string Email { get; set; }
       public string HasSupplies { get; set; }
       public string IsWorking { get; set; }
       public bool  Isservice { get; set; }
       public string Issue { get; set; }
       public bool Issupply { get; set; }
       public int Magenta { get; set; }
       public string Phone { get; set; }
       public int Yellow { get; set; }
    }
}