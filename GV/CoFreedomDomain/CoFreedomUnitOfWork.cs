using GV.Domain;

namespace GV.CoFreedomDomain
{
    public interface ICoFreedomUnitOfWork : IUnitOfWork
    {

    }

    public class CoFreedomUnitOfWork : UnitOfWork
    {
        public CoFreedomUnitOfWork(ICoFreedomSessionFactory sessionFactoryHelper) : base(sessionFactoryHelper)
        {
        }
    }
}