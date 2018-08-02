using FluentNHibernate.Cfg;
using FluentNHibernate.Cfg.Db;
using GV.Domain.Entities;
using GV.Domain.Mappings;
using NHibernate;
using NHibernate.Cfg;

namespace GV.Domain
{
    public interface ISessionFactoryHelper
    {
        ISessionFactory CurrentSessionFactory { get; }
    }

    public class SessionFactoryHelper : ISessionFactoryHelper
    {
        private readonly string _connectionString;
        private ISessionFactory _currentSessionFactory;

        public SessionFactoryHelper(string connectionString)
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
                .Database(MsSqlConfiguration.MsSql2008.ConnectionString(connectionString))
                .Mappings(x =>
                {
                    x.FluentMappings.Add<DevicesMap>();
                    x.FluentMappings.Add<LocationMap>();
                    x.FluentMappings.Add<SchedulesMap>();
                    x.FluentMappings.Add<AssetReplacementMap>();
                    x.FluentMappings.Add<ScheduleServiceMap>();
                    x.FluentMappings.Add<CyclesMap>();
                    x.FluentMappings.Add<CyclePeriodMap>();
                    x.FluentMappings.Add<CyclePeriodScheduleMap>();
                    x.FluentMappings.Add<CycleReconciliationServicesMap>();
                    x.FluentMappings.Add<EasyLinkItemMap>();
                    x.FluentMappings.Add<EasyLinkMap>();
                    x.FluentMappings.Add<EasyLinkChildMatchMap>();
                });
        }
    }
}