using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using GVWebapi.RemoteData;
using System.ComponentModel.DataAnnotations;

namespace GVWebapi.Models
{
    public class VolumeTrends
    {
      public string Label { get; set; }
      public string Color { get; set; }
      public List<csDeviceVolumes_Result> Data { get; set; }
    }
    
}