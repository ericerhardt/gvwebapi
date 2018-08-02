namespace GVWebapi.Models.Reconciliation
{
    public class InvoicedServiceModel
    {
        public string MeterGroup { get; set; }
        public decimal ActualPages { get; set; }
        public decimal ContractedPages { get; set; }
        public decimal BaseServiceForCycle { get; set; }
        public decimal OverageCost { get; set; }
        public decimal Credit { get; set; }
        public long CycleReconciliationServiceId {get; set; }
    }
}