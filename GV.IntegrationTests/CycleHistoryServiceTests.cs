using System;
using System.Collections.Generic;
using System.Linq;
using FakeItEasy;
using GV.Domain;
using GV.Domain.Entities;
using GVWebapi.Services;
using NUnit.Framework;

namespace GV.IntegrationTests
{
    [TestFixture]
    public class CycleHistoryServiceTests
    {
        private ICycleHistoryService _cycleHistoryService;
        private IRepository _repository;

        [SetUp]
        public void Setup()
        {
            _repository = A.Fake<IRepository>();
            _cycleHistoryService = new CycleHistoryService(_repository, null);
        }

        [Test]
        public void should_be_able_to_get_available_cycles()
        {
            var returnValue = new List<SchedulesEntity>
            {
                new SchedulesEntity
                {
                    ScheduleId = 1,
                    CustomerId = 1650,
                    Name = "N20141030-100",
                    EffectiveDateTime = new DateTimeOffset(new DateTime(2012,03,25)),
                    ExpiredDateTime = new DateTimeOffset(new DateTime(2020,01,01)),
                    Term = 1,
                    MonthlyHwCost = 1,
                    MonthlySvcCost = 1,
                    IsDeleted = false
                }
            };

            A.CallTo(() => _repository.Find<SchedulesEntity>()).Returns(returnValue.AsQueryable());

            _cycleHistoryService.GetAvailableCycles(1650);
        }
    }
}
