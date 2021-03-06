using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using GV.Configuration;
using GV.Domain;
using GV.Domain.Entities;
using GV.ExtensionMethods;

namespace GV.Services
{
    public interface IEasyLinkService
    {
        void LoadFile(EasyLinkImportHistoryEntity easyLink);
        IList<EasyLinkViewModel> GetAll();
    }

    public class EasyLinkService : IEasyLinkService
    {
        private readonly IRepository _repository;
        private readonly IUnitOfWork _unitOfWork;
        private readonly IGlobalViewConfiguration _globalViewConfiguration;

        public EasyLinkService(IRepository repository, IUnitOfWork unitOfWork, IGlobalViewConfiguration globalViewConfiguration)
        {
            _repository = repository;
            _unitOfWork = unitOfWork;
            _globalViewConfiguration = globalViewConfiguration;
        }

        public void LoadFile(EasyLinkImportHistoryEntity easyLink)
        {
            var fullFilePath = Path.Combine(easyLink.FileLocation, easyLink.FileName);
            var allLines = LoadFileFromDisk(fullFilePath).Skip(1).ToList();
            easyLink.ImportedRecords = allLines.Count;

            _repository.Add(easyLink);
            _unitOfWork.Commit();

            var itemsToInsert = new List<EasyLinkDataInsertModel>();

            foreach (var lineItem in allLines)
            {
                var array = lineItem.Split(',');
                var model = new EasyLinkDataInsertModel();
                model.ImportID = easyLink.ImportID;
                model.Child = Convert.ToInt32(array[0]);
                model.EmailAddress = array[1];
                model.FaxNumber = array[2];
                model.TransDate = Convert.ToDateTime($"{array[3]} {array[4]}");
                model.Description = array[5];
                model.Duration = Convert.ToDecimal(array[6]);
                model.Pages = Convert.ToInt32(array[7]);
                model.Charge = Convert.ToDecimal(array[8]);
                model.Messages = array[9];

                itemsToInsert.Add(model);
            }

            using (var bulkCopy = new SqlBulkCopy(_globalViewConfiguration.GlobalViewConnectionString))
            {
                bulkCopy.DestinationTableName = "EasyLinkData";
                bulkCopy.WriteToServer(itemsToInsert.ToDataTable());
            }
        }

        public IList<EasyLinkViewModel> GetAll()
        {
            return _repository.Find<EasyLinkEntity>()
                .Select(x => EasyLinkViewModel.For(x))
                .ToList();
        }

        private static IEnumerable<string> LoadFileFromDisk(string fileLocation)
        {
            return File.ReadAllLines(fileLocation);
        }

        private class EasyLinkItemInsertModel
        {
            //this column has to be here for bulk insert to work
            public long EasyLinkItemId { get; set; }
            public long EasyLinkId { get; set; }
            public int Child { get; set; }
            public string Email { get; set; }
            public string Fax { get; set; }
            public DateTime DateTime { get; set; }
            public string Description { get; set; }
            public decimal Duration { get; set; }
            public int Pages { get; set; }
            public decimal Charge { get; set; }
            public string MessageNumber { get; set; }
        }
        private class EasyLinkDataInsertModel
        {
            //this column has to be here for bulk insert to work
            public long EasylinkImportID { get; set; }
            public DateTime PeriodDate { get; set; }
            public long ImportID { get; set; }
            public int Child { get; set; }
            public string EmailAddress { get; set; }
            public string FaxNumber { get; set; }
            public DateTime TransDate { get; set; }
            public string Description { get; set; }
            public decimal Duration { get; set; }
            public int Pages { get; set; }
            public decimal Charge { get; set; }
            public string Messages { get; set; }
        }
    }
}