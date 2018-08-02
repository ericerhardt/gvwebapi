using FluentNHibernate.Cfg;
using FluentNHibernate.Cfg.Db;
using GV.CoFreedomDomain.Mappings;
using GV.Domain;
using NHibernate;
using NHibernate.Cfg;

namespace GV.CoFreedomDomain
{
    public interface ICoFreedomSessionFactory : ISessionFactoryHelper
    {

    }

    public class CoFreedomSessionFactory : ICoFreedomSessionFactory
    {
        private readonly string _connectionString;
        private ISessionFactory _currentSessionFactory;

        public CoFreedomSessionFactory(string connectionString)
        {
            _connectionString = connectionString;
        }

        public ISessionFactory CurrentSessionFactory => _currentSessionFactory ?? (_currentSessionFactory = CreateSessionFactory());

        private ISessionFactory CreateSessionFactory()
        {
            return GetFluentConfiguration(_connectionString).ExposeConfiguration(config =>
                {
                    config.SetProperty(Environment.Hbm2ddlKeyWords, "none");
                    config.SetProperty(Environment.CommandTimeout, "60");
                })
                .BuildSessionFactory();
        }

        private static FluentConfiguration GetFluentConfiguration(string connectionString)
        {
            return Fluently.Configure()
                .Database(MsSqlConfiguration.MsSql2012.ConnectionString(connectionString))
                .Mappings(x =>
                {
                    x.FluentMappings.Add<ScEquipmentCustomPropertiesMap>();
                    x.FluentMappings.Add<ScEquipmentMap>();
                    x.FluentMappings.Add<ArCustomersMap>();
                    x.FluentMappings.Add<IcModelMap>();
                    x.FluentMappings.Add<ScContractsMap>();
                    x.FluentMappings.Add<ScContractMeterGroupsMap>();
                    x.FluentMappings.Add<ScContractDetailsMap>();
                    x.FluentMappings.Add<ViewEquipmentAndRateMap>();
                });
        }
    }
}