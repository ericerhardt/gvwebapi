using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GVWebapi.Models.CostAllocation
{
    public partial class v_SCContractMeterGroups
    {
       [Key]
       public virtual int ContractID { get; set; }
       public virtual int ContractMeterGroupID { get; set; }
       public virtual string ContractMeterGroup { get; set; }
       public virtual decimal  CoveredCopies { get; set; }
       public virtual bool ApplyToExpiration { get; set; }
       public virtual bool UseOverages { get; set; }
       public virtual decimal Rate { get; set; }
    }
}