﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models
{
    public class SccMeterGroupSumsViewModel
    {
        public long MeterGroupID {get;set;}
        public string Name { get; set; }
        public int? Total { get; set; }
    }
}