using GV.Domain.Entities;

namespace GVWebapi.Models.Locations
{
    public class LocationViewModel
    {
        public static LocationViewModel For(LocationEntity location, CoFreedomLocationModel coFreedomModel)
        {
            var model = new LocationViewModel();
            model.LocationId = location.LocationId;
            model.Name = location.Name;
            model.TaxRate = location.TaxRate;
            model.IsCorporate = location.IsCorporate;
            model.Address = coFreedomModel.Address;
            model.City = coFreedomModel.City;
            model.State = coFreedomModel.State;
            model.Zip = coFreedomModel.Zip;
            model.CoFreedomLocationId = location.CoFreedomLocationId;
            return model;
        }

        private LocationViewModel()
        {
        }

        public long LocationId { get; set; }
        public string Name { get; set; }
        public decimal TaxRate { get; set; }
        public bool IsCorporate { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Zip { get; set; }
        public string FullAddress => $"{City}, {State} {Zip}";
        public int CoFreedomLocationId { get; set; }
    }
}