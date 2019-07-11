using GVWebapi.Models.CostAllocation;
using GVWebapi.RemoteData;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Web;

namespace GVWebapi.Entities
{
    public partial class FprEntities : DbContext
    {
        public FprEntities()
             : base("name=FprEntities")
        {
        }
        public virtual DbSet<vw_EquipmentInvoiceHistory> vw_EquipmentInvoiceHistory { get; set; }
        public virtual DbSet<vw_GV_InvoicedEquipmentHistory> vw_GV_InvoicedEquipmentHistory { get; set; }
        public virtual DbSet<v_SCContractMeterGroups> v_SCContractMeterGroups { get; set; }
    }
}