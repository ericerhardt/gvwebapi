using System;
using System.Collections.Generic;
using System.Linq;
using GV.CoFreedomDomain;
using GV.CoFreedomDomain.Entities;
using GV.Domain;
using GV.Domain.Entities;
using GV.Lookup;

namespace GV.Services
{
    public interface IEasyLinkChildManagerService
    {
        IList<LookupInfo> GetAllCustomers(long? customerToLeaveIn = null);
        IList<LookupInfo> GetAllEasyLinkChildIds(int? childToLeaveIn = null);
        void AddEasyLinkChildMatch(long customerId, int childId, bool isEasyLinkOnly);
        IList<EasyLinkChildManagerModel> GetChildLinks(bool hideEasyLinkOnly);
        void RemoveLink(long easyLinkChildMatchId);
        void SwitchChildLink(EasyLinkChildSwitchModel model);
        IList<EasyLinkUnMappedChildModel> GetUnMappedChildIds();
    }

    public class EasyLinkChildManagerService : IEasyLinkChildManagerService
    {
        private readonly ICoFreedomRepository _coFreedomRepository;
        private readonly IRepository _repository;

        public EasyLinkChildManagerService(ICoFreedomRepository coFreedomRepository, IRepository repository)
        {
            _coFreedomRepository = coFreedomRepository;
            _repository = repository;
        }

        public IList<LookupInfo> GetAllCustomers(long? customerToLeaveIn = null)
        {
            var customersAlreadyAdded = _repository.Find<EasyLinkChildMatchEntity>()
                    .Where(x => x.IsDeleted == false)
                    .Select(x => x.CustomerId)
                    .Distinct()
                    .ToList();

            if (customerToLeaveIn.HasValue)
                customersAlreadyAdded.Remove(customerToLeaveIn.Value);

            return _coFreedomRepository.Find<ArCustomersEntity>()
                .Where(x => customersAlreadyAdded.Contains(x.CustomerId) == false)
                .OrderBy(x => x.CustomerName)
                .Select(x => new LookupInfo(x.CustomerId, x.CustomerName))
                .ToList();
        }

        public IList<LookupInfo> GetAllEasyLinkChildIds(int? childToLeaveIn = null)
        {
            var subQuery = _repository.Find<EasyLinkChildMatchEntity>()
                .Where(x => x.IsDeleted == false);

            if (childToLeaveIn.HasValue)
                subQuery = subQuery.Where(x => x.ChildId != childToLeaveIn.Value);

            var finalSubQuery = subQuery.Select(x => x.ChildId).Distinct();

            var items = _repository.Find<EasyLinkItemEntity>()
                .Where(x => finalSubQuery.Contains(x.Child) == false)
                .Select(x => x.Child)
                .Select(x => new LookupInfo(x, x.ToString()))
                .Distinct()
                .ToList();

            return items.OrderBy(x => x.Id).ToList();
        }

        public void AddEasyLinkChildMatch(long customerId, int childId, bool isEasyLinkOnly)
        {
            var matchEntity = new EasyLinkChildMatchEntity();
            matchEntity.CustomerId = customerId;
            matchEntity.ChildId = childId;
            matchEntity.IsEasyLinkOnly = isEasyLinkOnly;
            matchEntity.CreatedDateTime = DateTimeOffset.Now;
            _repository.Add(matchEntity);
        }

        public IList<EasyLinkChildManagerModel> GetChildLinks(bool hideEasyLinkOnly)
        {
            var childLinksToShow = _repository.Find<EasyLinkChildMatchEntity>()
                .Where(x => x.IsDeleted == false)
                .ToList();

            var customersToGet = childLinksToShow
                .Select(x => x.CustomerId)
                .Distinct()
                .ToList();

            var customers = _coFreedomRepository.Find<ArCustomersEntity>()
                .Where(x => customersToGet.Contains(x.CustomerId))
                .ToList();

            var modelsToView = new List<EasyLinkChildManagerModel>();

            foreach (var childLink in childLinksToShow)
            {
                if(hideEasyLinkOnly && childLink.IsEasyLinkOnly) continue;
                var model = new EasyLinkChildManagerModel();
                model.CustomerId = childLink.CustomerId;
                model.ChildId = childLink.ChildId;
                model.EasyLinkChildMatchId = childLink.EasyLinkChildMatchId;
                model.IsEasyLinkOnly = childLink.IsEasyLinkOnly;

                var customer = customers.FirstOrDefault(x => x.CustomerId == childLink.CustomerId);
                if (customer != null)
                    model.CustomerName = customer.CustomerName;

                modelsToView.Add(model);
            }

            return modelsToView;
        }

        public void RemoveLink(long easyLinkChildMatchId)
        {
            var childLink = _repository.Get<EasyLinkChildMatchEntity>(easyLinkChildMatchId);
            childLink.IsDeleted = true;
            childLink.ModifiedDateTime = DateTimeOffset.Now;
        }

        public void SwitchChildLink(EasyLinkChildSwitchModel model)
        {
            var existingItem = _repository.Find<EasyLinkChildMatchEntity>()
                .Where(x => x.CustomerId == model.ExistingCustomerId)
                .Where(x => x.ChildId == model.ExistingChildId)
                .FirstOrDefault(x => x.IsDeleted == false);

            if (existingItem != null)
            {
                existingItem.IsDeleted = true;
                existingItem.ModifiedDateTime = DateTimeOffset.Now;
            }

            var newMatch = new EasyLinkChildMatchEntity();
            newMatch.CustomerId = model.NewCustomerId;
            newMatch.ChildId = model.NewChildId;
            newMatch.CreatedDateTime = DateTimeOffset.Now;
            newMatch.IsEasyLinkOnly = model.IsEasyLinkOnly;
            _repository.Add(newMatch);
        }

        public IList<EasyLinkUnMappedChildModel> GetUnMappedChildIds()
        {
            var mappedIds = _repository
                .Find<EasyLinkChildMatchEntity>()
                .Where(x => x.IsDeleted == false)
                .Select(x => x.ChildId)
                .Distinct();

            return _repository.Find<EasyLinkItemEntity>()
                .Where(x => mappedIds.Contains(x.Child) == false)
                .GroupBy(x => x.Child)
                .Select(x => new EasyLinkUnMappedChildModel
                {
                    ChildId = x.Key,
                    Count = x.Count(),
                    TotalPages =  x.Sum(page => page.Pages),
                    TotalCharge = x.Sum(charge => charge.Charge)
                }).ToList();
        }
    }
}