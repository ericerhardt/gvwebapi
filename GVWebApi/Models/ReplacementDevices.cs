namespace GVWebapi.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("AssetReplacement")]
    public partial class ReplacementDevices
    {
        [Key]
        public int ReplacementID { get; set; }

        public int? CustomerID { get; set; }

        [StringLength(250)]
        public string Location { get; set; }

        public DateTime? ReplacementDate { get; set; }

        [StringLength(50)]
        public string OldSerialNumber { get; set; }

        [StringLength(50)]
        public string OldModel { get; set; }

        [StringLength(50)]
        public string NewSerialNumber { get; set; }

        [StringLength(50)]
        public string NewModel { get; set; }

        public decimal? ReplacementValue { get; set; }

        [StringLength(1250)]
        public string Comments { get; set; }
    }
}
