using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using GVWebapi.RemoteData;
using GV.Configuration;
using GVWebapi.Models;
using GVWebapi.Models.Easylink;
using System.Data.SqlClient;
using GV.ExtensionMethods;
using System.Net.Http;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using GV.Lookup;
namespace GVWebapi.Services
{
    public interface IEasyLinkFileSaveService
    {
        EasylinkImportHistory SaveFile(EasyLinkUploadModel model, string NewFilePath);
        void ImportData(EasylinkImportHistory model);
        EasyLinkUploadModel GetFormData(MultipartFormDataStreamProvider result);
        CustomMultipartFormDataStreamProvider GetMultipartProvider();
        string GetDeserializedFileName(MultipartFileData fileData);
        string GetFileName(MultipartFileData fileData);
        void AddEasyLinkChildMatch(int customerId, int childId, bool isEasyLinkOnly);
        void RemoveLink(int childId, int clientId);
        IList<EasyLinkUnMappedChildModel> GetUnMappedChildIds();
        IList<EasyLinkChildManagerModel> GetChildLinks(bool hideEasyLinkOnly);
        IList<LookupInfo> GetAllCustomers(int? customerToLeaveIn = null);
        IList<LookupInfo> GetAllEasyLinkChildIds(int? childToLeaveIn = null);
    }

    public class EasyLinkServices:  IEasyLinkFileSaveService
    {
         
        private readonly GlobalViewEntities _globalView = new GlobalViewEntities();
        private readonly CoFreedomEntities _coFreedomRepository = new CoFreedomEntities();

        public EasylinkImportHistory SaveFile(EasyLinkUploadModel model, string NewFilePath)
        {
            var fullPath = GetFileAndPath(NewFilePath);
            return CreateAndReturnEasyLinkEntity(fullPath, model);
        }

        private EasylinkImportHistory CreateAndReturnEasyLinkEntity(string newFileAndPath, EasyLinkUploadModel model)
        {
            var fileInfo = new FileInfo(newFileAndPath);
            var easyLinkEntity = new EasylinkImportHistory();
            //easyLinkEntity.FileName = model.FileName.Replace("\"", string.Empty);
            easyLinkEntity.FileName = model.EasyLinkFile.FileName;
            easyLinkEntity.FileLocation = fileInfo.DirectoryName;
            easyLinkEntity.PeriodDate = DateTime.Parse(model.PeriodDate);
            easyLinkEntity.ImportedOn =   DateTime.Now;
            easyLinkEntity.ImportedBy = model.UserName;
            return easyLinkEntity;
        }

        private static void SaveFileAndContents(string newFileAndPath, byte[] contents)
        {
            File.WriteAllBytes(newFileAndPath, contents);
        }

        private string GetFileAndPath(string fileName)
        {
            var basePath = System.Configuration.ConfigurationManager.AppSettings["EasyLink.FileSavePath"];
            if (Directory.Exists(basePath) == false)
                Directory.CreateDirectory(basePath);
            fileName = fileName.Replace("\"", string.Empty);
            var newFileName = $"{fileName.Substring(0, fileName.IndexOf('.'))}_{Guid.NewGuid()}.csv";
            return Path.Combine(basePath, newFileName);
        }
        private static IEnumerable<string> LoadFileFromDisk(string fileLocation)
        {
            return File.ReadAllLines(fileLocation);
        }
        public void ImportData(EasylinkImportHistory easyLink)
        {
            var fullFilePath = Path.Combine(easyLink.FileLocation, easyLink.FileName);
            var allLines = LoadFileFromDisk(fullFilePath).Skip(1).ToList();
            easyLink.ImportRecords = allLines.Count;

            _globalView.EasylinkImportHistories.Add(easyLink);
            _globalView.SaveChanges();

            var itemsToInsert = new List<EasylinkData>();

            foreach (var lineItem in allLines)
            {
                var array = lineItem.Split(',');
                var model = new EasylinkData();
                model.ImportID = easyLink.ImportID;
                model.Child = Convert.ToInt32(array[0]);
                model.emailaddress = array[1];
                model.FaxNumber = array[2];
                model.TransDate = Convert.ToDateTime($"{array[3]} {array[4]}");
                model.Description = array[5];
                model.PeriodDate = easyLink.PeriodDate;
                model.Duration = Convert.ToDecimal(array[6]);
                model.Pages = Convert.ToInt32(array[7]);
                model.Charges = Convert.ToDecimal(array[8]);
                model.Messages = array[9];

                itemsToInsert.Add(model);
            }

            _globalView.EasylinkDatas.AddRange(itemsToInsert);
            _globalView.SaveChanges();
           
         }

        public EasyLinkUploadModel GetFormData(MultipartFormDataStreamProvider result)
        {
            if (result.FormData.HasKeys())
            {
                var model = new EasyLinkUploadModel();
                model.PeriodDate = result.FormData["periodDate"];
                model.UserName = result.FormData["userName"];
                return model;
            }

            return null;
        }
        public CustomMultipartFormDataStreamProvider GetMultipartProvider()
        {
            
            var root = System.Configuration.ConfigurationManager.AppSettings["EasyLink.FileSavePath"];  
            return new CustomMultipartFormDataStreamProvider(root);
        }
        public string GetDeserializedFileName(MultipartFileData fileData)
        {
            var fileName = GetFileName(fileData);
            return JsonConvert.DeserializeObject(fileName).ToString();
        }

        public string GetFileName(MultipartFileData fileData)
        {
            return fileData.Headers.ContentDisposition.FileName;
        }
        public void RemoveLink(int childId, int clientId)
        {
            var childLink = _globalView.EasyLinkMappings.Where(x => x.ChildId == childId && x.ClientId == clientId).FirstOrDefault();
            if(childLink != null){
                _globalView.EasyLinkMappings.Remove(childLink);
                _globalView.SaveChanges();
            }
          
        }

        public IList<EasyLinkUnMappedChildModel> GetUnMappedChildIds()
        {
            var mappedIds = _globalView.EasyLinkMappings
                .Select(x => x.ChildId)
                .Distinct();

            return _globalView.EasylinkDatas
                .Where(x => mappedIds.Contains(x.Child) == false)
                .GroupBy(x => x.Child)
                .Select(x => new EasyLinkUnMappedChildModel
                {
                    ChildId = x.Key.Value,
                    Count = x.Count(),
                    TotalPages = x.Sum(page => page.Pages.Value),
                    TotalCharge = x.Sum(charge => charge.Charges.Value)
                }).ToList();
        }

        public IList<EasyLinkChildManagerModel> GetChildLinks(bool hideEasyLinkOnly)
        {
            var childLinksToShow = _globalView.EasyLinkMappings.ToList();

            var customersToGet = childLinksToShow
                .Select(x => x.ClientId)
                .Distinct()
                .ToList();

            var customers = _coFreedomRepository.ARCustomers
                .Where(x => customersToGet.Contains(x.CustomerID))
                .ToList();

            var modelsToView = new List<EasyLinkChildManagerModel>();

            foreach (var childLink in childLinksToShow)
            {
                if (hideEasyLinkOnly && childLink.IsEasyLinkOnly.Value) continue;
                var model = new EasyLinkChildManagerModel();
                model.CustomerId = childLink.ClientId.Value;
                model.ChildId = childLink.ChildId.Value;
                model.EasyLinkChildMatchId = childLink.MappingID;
                model.IsEasyLinkOnly = childLink.IsEasyLinkOnly.Value;

                var customer = customers.FirstOrDefault(x => x.CustomerID == childLink.ClientId);
                if (customer != null)
                    model.CustomerName = customer.CustomerName;

                modelsToView.Add(model);
            }

            return modelsToView.OrderBy( x => x.CustomerName).ToList();
        }
        public IList<LookupInfo> GetAllCustomers(int? customerToLeaveIn = null)
        {
            var customersAlreadyAdded = _globalView.EasyLinkMappings
                    .Select(x => x.ClientId)
                    .Distinct()
                    .ToList();

            if (customerToLeaveIn.HasValue)
                customersAlreadyAdded.Remove(customerToLeaveIn.Value);

            return _coFreedomRepository.vw_ClientsOnContract
                .Where(x => customersAlreadyAdded.Contains(x.CustomerID) == false)
                .OrderBy(x => x.CustomerName)
                .Select(x => new LookupInfo {
                 Id   =    x.CustomerID,
                 Display =    x.CustomerName
                 })
                .ToList();
        }

        public IList<LookupInfo> GetAllEasyLinkChildIds(int? childToLeaveIn = null)
        {
            var subQuery = _globalView.EasyLinkMappings.AsEnumerable();


            if (childToLeaveIn.HasValue)
                subQuery = subQuery.Where(x => x.ChildId != childToLeaveIn.Value);

            var finalSubQuery = subQuery.Select(x => x.ChildId).Distinct();

            var items = _globalView.EasylinkDatas
                .Where(x => finalSubQuery.Contains(x.Child) == false)
                .Select(x => x.Child)
                 .Select(x => new LookupInfo
                 {
                     Id = x.Value,
                     Display = x.ToString()
                 })
                .Distinct()
                .ToList();

            return items.OrderBy(x => x.Id).ToList();
        }
        public void AddEasyLinkChildMatch(int customerId, int childId, bool isEasyLinkOnly)
        {
            var matchEntity = new EasyLinkMapping();
            matchEntity.ClientId = customerId;
            matchEntity.ChildId = childId;
            matchEntity.IsEasyLinkOnly = isEasyLinkOnly;
            _globalView.EasyLinkMappings.Add(matchEntity);
            _globalView.SaveChanges();
        }

    }


}