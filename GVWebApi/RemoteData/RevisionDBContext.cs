using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;
using GVWebapi.Models;

namespace GVWebapi.RemoteData
{
    public class RevisionDBContext : DbContext
    {
        public RevisionDBContext() :base("PeriodHistoryConnection")
        {
            
        } 
        public DbSet<PeriodHistoryView> PeriodHistoryView { get;set;}
    }
}