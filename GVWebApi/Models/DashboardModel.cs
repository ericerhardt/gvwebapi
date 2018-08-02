using System.Collections.Generic;

namespace GVWebapi.Models
{
    public class DashboardModel
    {
        public decimal? ReplacementTotals { get; set; }
        public decimal? RevisionTotals { get; set; }
        public decimal? CostAvoidanceTotals { get; set; }
        public decimal? RollOversTotals { get; set; }
        public IEnumerable<LocationsModel> Locations { get; set; }
    }
}