﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using GVWebapi.Models;
namespace GVWebapi.RemoteData
{
    public class RevisionHistoryModel 
    {
    public DateTime? peroid { get; set; }
  public IEnumerable<PeriodHistoryView> detail {get;set;}
    }
}