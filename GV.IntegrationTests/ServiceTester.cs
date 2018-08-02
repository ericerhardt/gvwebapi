using GV.Ninject;
using GVWebapi.Services;
using NUnit.Framework;

namespace GV.IntegrationTests
{
    [TestFixture]
    public class ServiceTester
    {
        [Test]
        public void should_be_able_to_resolve_tests()
        {
            Ioc.Do<IScheduleServicesService>(service => {});
        }
    }
}