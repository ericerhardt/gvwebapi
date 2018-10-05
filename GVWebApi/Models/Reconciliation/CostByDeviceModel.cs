namespace GVWebapi.Models.Reconciliation
{
    public class CostByDeviceModel
    {
        public string SerialNumber { get; set; }
        public string Model { get; set; }
        public string Schedule { get; set; }
        public string Location { get; set; }
        public string User { get; set; }
        public string CostCenter { get; set; }
        public decimal BWVolume { get; set; }
        public decimal ColorVolume { get; set; }
        public decimal BWCopies { get; set; }
        public decimal BWPrints { get; set; }
        public decimal ColorPrints { get; set; }
        public decimal ColorCopies { get; set; }
        public decimal Credit { get; set; }
        public decimal Total => BWPrints + BWCopies + ColorCopies + ColorPrints + Credit;
    }
}