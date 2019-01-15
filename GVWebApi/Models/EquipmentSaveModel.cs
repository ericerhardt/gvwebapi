using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models
{
    public class EquipmentSaveModel
    {
        public long Id { get; set; }
        public string Model { get; set; }
        public Nullable<System.DateTime> IntroDate { get; set; }
        public Nullable<int> MFRMoVol { get; set; }
        public Nullable<decimal> PurchasePrice { get; set; }
        public string Image { get; set; }
        public HttpPostedFileBase Attachment { get; set; }
        public Nullable<System.DateTime> LastUpdatedOn { get; set; }
        public Nullable<int> Status { get; set; }
        public string LastUpdatedBy { get; set; }
    }
}