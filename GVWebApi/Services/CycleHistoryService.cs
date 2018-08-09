﻿using System;
using System.Collections.Generic;
using System.Linq;
using FluentDateTime;
using GV.Domain;
using GV.Domain.Entities;

namespace GVWebapi.Services
{
    public interface ICycleHistoryService
    {
        List<DateTime> GetAvailableCycles(long customerId);
        void AddNewCycle(NewCycleModel model);
        IList<CycleHistoryViewModel> GetActiveCycles(long customerId);
        void ToggleAvailability(ToggleSaveModel model);
        void DeleteCycle(long cycleId);
        void ToggleReconcile(CycleReconcileSaveModel model);
    }

    public class CycleHistoryService : ICycleHistoryService
    {
        private readonly IRepository _repository;
        private readonly ICyclePeriodService _cyclePeriodService;

        public CycleHistoryService(IRepository repository, ICyclePeriodService cyclePeriodService)
        {
            _repository = repository;
            _cyclePeriodService = cyclePeriodService;
        }

        public List<DateTime> GetAvailableCycles(long customerId)
        {
            var schedules = _repository.Find<SchedulesEntity>()
                .Where(x => x.CustomerId == customerId)
                .Where(x => x.IsDeleted == false)
                .Where(x => x.ExpiredDateTime >= DateTimeOffset.Now)
                .Where(x => x.EffectiveDateTime <= DateTimeOffset.Now)
                .ToList();

            var availableDates = new List<DateTime>();
            foreach (var schedule in schedules)
            {
                var schedulesDates = GetScheduleDates(schedule);
                var startDateTime = schedulesDates.StartDate.FirstDayOfMonth();
                while (startDateTime <= schedulesDates.EndDate)
                {
                    if (availableDates.Contains(startDateTime) == false)
                        availableDates.Add(startDateTime);
                    startDateTime = startDateTime.AddMonths(1);
                }

                availableDates = availableDates
                    .Where(x => x.IsBefore(DateTime.Now.Date))
                    .Where(x => x.IsAfter(schedulesDates.StartDate.Date))
                    .ToList();
            }

            var activeCycles = GetActiveCycles(customerId);

            var finalDates = new List<DateTime>();

            foreach (var availableDate in availableDates)
            {
                var alreadyContainsCycle = activeCycles.Any(x => x.CycleStartDate == availableDate);
                if(alreadyContainsCycle) continue;

                var cycleHasPeriodWithThisDate = activeCycles.Any(cycle => cycle.Periods.Any(x => x.Period == availableDate));
                if(cycleHasPeriodWithThisDate) continue;

                finalDates.Add(availableDate);
            }

            return finalDates;
        }

        public void AddNewCycle(NewCycleModel model)
        {
            var newCycle = new CyclesEntity(model.CustomerId, model.CycleDate);
            newCycle.InvisibleToClient = true;
            _repository.Add(newCycle);
        }

        public IList<CycleHistoryViewModel> GetActiveCycles(long customerId)
        {
            var items = _repository.Find<CyclesEntity>()
                .Where(x => x.InActive == false)
                .Where(x => x.IsDeleted == false)
                .Where(x => x.CustomerId == customerId)
                .OrderByDescending(x => x.StartDate)
                .Select(x => new CycleHistoryViewModel
                {
                    CycleHistoryId = x.CycleId,
                    InvisibleToClient = x.InvisibleToClient,
                    CycleStartDate = x.StartDate,
                    EndDate = x.EndDate,
                    IsReconciled = x.IsReconciled
                    
                }).ToList();

            
            var maxItem = items.Count;
            foreach (var item in items)
            {
                item.CycleNumber = maxItem;
                item.Periods = _cyclePeriodService.GetCyclePeriods(item.CycleHistoryId);
                maxItem--;
            }

            return items;
        }

        public void ToggleAvailability(ToggleSaveModel model)
        {
            var cycle = _repository.Get<CyclesEntity>(model.CycleId);
            cycle.InvisibleToClient = model.InvisibleToClient;
        }

        public void DeleteCycle(long cycleId)
        {
            var cycle = _repository.Get<CyclesEntity>(cycleId);
            cycle.IsDeleted = true;
        }

        public void ToggleReconcile(CycleReconcileSaveModel model)
        {
            var cycle = _repository.Get<CyclesEntity>(model.CycleId);
            cycle.IsReconciled = model.IsReconciled;
        }

        private static InternalDate GetScheduleDates(SchedulesEntity schedule)
        {
            if (schedule.CoterminousSchedule != null)
            {
                GetScheduleDates(schedule.CoterminousSchedule);
            }

            if (schedule.EffectiveDateTime.HasValue == false || schedule.ExpiredDateTime.HasValue == false)
                throw new ApplicationException("Invalid Schedule Dates");

            return new InternalDate(schedule.EffectiveDateTime.Value.Date, schedule.ExpiredDateTime.Value.Date);
        }

        private class InternalDate
        {
            public InternalDate(DateTime startDate, DateTime endDate)
            {
                EndDate = endDate;
                StartDate = startDate;
            }

            public DateTime StartDate { get; set; }
            public DateTime EndDate { get; set; }
        }
    }

    public class CycleHistoryViewModel
    {
        public  long CycleHistoryId { get; set; }
        public bool InvisibleToClient { get; set; }
        public DateTime CycleStartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public int CycleNumber { get; set; }
        public decimal Billed {get; set; }
        public decimal Allocated { get; set; }
        public IList<CyclePeriodModel> Periods { get; set; } = new List<CyclePeriodModel>();
        public bool IsReconciled { get; set; }
    }

    public class NewCycleModel
    {
        public long CustomerId {get; set; }
        public DateTime CycleDate { get; set; }
    }

    public class ToggleSaveModel
    {
        public long CycleId { get; set; }
        public bool InvisibleToClient {get; set; }
    }

    public class CycleReconcileSaveModel
    {
        public bool IsReconciled { get; set; }
        public long CycleId { get; set; }
    }
}