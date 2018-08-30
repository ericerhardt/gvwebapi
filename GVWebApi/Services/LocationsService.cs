using System;
using System.Collections.Generic;
using System.Linq;
using GV.CoFreedomDomain;
using GV.CoFreedomDomain.Entities;
using GV.Domain;
using GV.Domain.Entities;
using GV.ExtensionMethods;
using GVWebapi.Models.Locations;

namespace GVWebapi.Services
{
    public interface ILocationsService
    {
        IList<LocationViewModel> LoadAll(long customerId);
        void SetCorporate(long locationId, bool newValue);
        void UpdateTaxRate(long locationId, decimal newTaxRate);
        IList<LocationViewModel> LoadAllByDeviceId(long deviceId);
        decimal GetTaxRate(string location);
        void UpdateAllRates(decimal rate);
    }

    public class LocationsService : ILocationsService
    {
        private readonly IRepository _repository;
        private readonly ICoFreedomRepository _coFreedomRepository;

        public LocationsService(IRepository repository, ICoFreedomRepository coFreedomRepository)
        {
            _repository = repository;
            _coFreedomRepository = coFreedomRepository;
        }

        public IList<LocationViewModel> LoadAll(long customerId)
        {
            var coFreedomLocations = LoadCoFreedomLocations(customerId);
            var globalViewEntities = LoadGlobalViewLocations(customerId);

            //new added items
            foreach (var coFreedomLocationModel in coFreedomLocations)
            {
                var existingLocation = globalViewEntities.FirstOrDefault(x => x.Name.EqualsIgnore(coFreedomLocationModel.Name));
                if (existingLocation != null) continue;
                var newLocation = new LocationEntity(customerId, coFreedomLocationModel.Name, coFreedomLocationModel.LocationId);
                _repository.Add(newLocation);
            }

            //set deleted items
            foreach (var globalViewEntity in globalViewEntities)
            {
                var existingLocation = coFreedomLocations.FirstOrDefault(x => x.Name.EqualsIgnore(globalViewEntity.Name));
                if(existingLocation != null) continue;
                globalViewEntity.IsDeleted = true;
            }

            return MergeCoFreedomAndGlobalView(coFreedomLocations, LoadGlobalViewLocations(customerId));
        }

        public void SetCorporate(long locationId, bool newValue)
        {
            var originalLocation = _repository.Get<LocationEntity>(locationId);
            if (originalLocation == null) return;
            var locations = _repository.Find<LocationEntity>()
                .Where(x => x.CustomerId == originalLocation.CustomerId)
                .Where(x => x.LocationId != locationId)
                .ToList();

            foreach (var location in locations)
            {
                location.IsCorporate = false;
            }

            originalLocation.IsCorporate = newValue;
        }

        public void UpdateTaxRate(long locationId, decimal newTaxRate)
        {
            var location = _repository.Get<LocationEntity>(locationId);
            location.ModifiedDateTime = DateTimeOffset.Now;
            location.TaxRate = newTaxRate;
        }

        public IList<LocationViewModel> LoadAllByDeviceId(long deviceId)
        {
            var device = _coFreedomRepository.Get<ScEquipmentEntity>(deviceId);
            return LoadAll(device.CustomerId);
        }

        public decimal GetTaxRate(string location)
        {
            var locationEntity = _repository
                .Find<LocationEntity>()
                .Where(x => x.IsDeleted == false)
                .FirstOrDefault(x => x.Name.ToLower() == location.ToLower());

            return locationEntity?.TaxRate ?? 0M;
        }

        public void UpdateAllRates(decimal rate)
        {
            var allLocations = _repository.Find<LocationEntity>()
                .Where(x => x.IsDeleted == false)
                .ToList();

            foreach (var location in allLocations)
            {
                location.TaxRate = rate;
            }
        }

        private static List<LocationViewModel> MergeCoFreedomAndGlobalView(IList<CoFreedomLocationModel> coFreedomLocations, IEnumerable<LocationEntity> globalViewEntities)
        {
            var viewModels = new List<LocationViewModel>();

            foreach (var globalViewEntity in globalViewEntities)
            {
                var existingItem = coFreedomLocations.FirstOrDefault(x => x.Name.EqualsIgnore(globalViewEntity.Name));
                if (existingItem == null) continue;
                viewModels.Add(LocationViewModel.For(globalViewEntity, existingItem));
            }

            return viewModels;
        }

        private List<LocationEntity> LoadGlobalViewLocations(long customerId)
        {
            return _repository.Find<LocationEntity>()
                .Where(x => x.CustomerId == customerId)
                .ToList();
        }

        private  IList<CoFreedomLocationModel> LoadCoFreedomLocations(long customerId)
        {
            var subQuery = _coFreedomRepository.Find<ScEquipmentEntity>()
                .Where(x => x.CustomerId == customerId)
                .Where(x => x.Active)
                .Select(x => x.Location.CustomerId)
                .Distinct();

            return _coFreedomRepository.Find<ArCustomersEntity>()
                .Where(x => subQuery.Contains(x.CustomerId))
                .OrderBy(x => x.CustomerName)
                .Select(x => new CoFreedomLocationModel
                {
                    Name = x.CustomerName,
                    Address = x.Address,
                    City = x.City,
                    State = x.State,
                    Zip = x.Zip,
                    LocationId = x.CustomerId
                })
                .ToList();
        }
    }
}