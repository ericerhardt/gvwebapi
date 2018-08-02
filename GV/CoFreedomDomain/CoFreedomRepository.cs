using GV.Domain;

namespace GV.CoFreedomDomain
{
    public interface ICoFreedomRepository : IRepository
    {
        
    }

    public class CoFreedomRepository : Repository
    {
        public CoFreedomRepository(ICoFreedomUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }
    }
}