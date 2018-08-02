using GV.Ninject;
using GVWebapi.Ninject;
using Ninject;
using NUnit.Framework;

namespace GV.IntegrationTests
{
    [SetUpFixture]
    public class BeforeAllTests
    {
        [OneTimeSetUp]
        public void OneTimeSetup()
        {
            var kernel = new StandardKernel();
            kernel.Load(new GvNinjectLoadModule());
            kernel.Load(new NinjectApiLoadModule());
            Ioc.Initialize(kernel);
        }
    }
}