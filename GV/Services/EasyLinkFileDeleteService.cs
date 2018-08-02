using System.IO;
using GV.Domain;
using GV.Domain.Entities;

namespace GV.Services
{
    public interface IEasyLinkFileDeleteService
    {
        void DeleteFile(long easyLinkId);
    };

    public class EasyLinkFileDeleteService : IEasyLinkFileDeleteService
    {
        private readonly IRepository _repository;

        public EasyLinkFileDeleteService(IRepository repository)
        {
            _repository = repository;
        }

        public void DeleteFile(long easyLinkId)
        {
            var easyLink = _repository.Get<EasyLinkEntity>(easyLinkId);

            var pathToFile = Path.Combine(easyLink.FileLocation, easyLink.SavedFileName);
            File.Delete(pathToFile);

            _repository.Remove(easyLink);
        }
    }
}