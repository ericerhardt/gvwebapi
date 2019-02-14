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

namespace GVWebapi.Services
{
    public interface IEasyLinkFileSaveService
    {
        EasylinkImportHistory SaveFile(EasyLinkFileSaveModel model);
        void ImportData(EasylinkImportHistory model);
    }

    public class EasyLinkServices  
    {
        private readonly IGlobalViewConfiguration _globalViewConfiguration;
        private readonly GlobalViewEntities _globalView;
        public EasyLinkServices(IGlobalViewConfiguration globalViewConfiguration,GlobalViewEntities globalView)
        {
            _globalViewConfiguration = globalViewConfiguration;
            _globalView = globalView;
        }

        public EasylinkImportHistory SaveFile(EasyLinkFileSaveModel model)
        {
            var newFileAndPath = GetFileAndPath(model.FileName);
            SaveFileAndContents(newFileAndPath, model.Contents);
            return CreateAndReturnEasyLinkEntity(newFileAndPath, model);
        }

        private EasylinkImportHistory CreateAndReturnEasyLinkEntity(string newFileAndPath, EasyLinkFileSaveModel model)
        {
            var fileInfo = new FileInfo(newFileAndPath);
            var easyLinkEntity = new EasylinkImportHistory();
            easyLinkEntity.FileName = model.FileName.Replace("\"", string.Empty);
            easyLinkEntity.FileName = fileInfo.Name;
            easyLinkEntity.FileLocation = fileInfo.DirectoryName;
            easyLinkEntity.ImportedOn =   DateTime.Now;
          
            return easyLinkEntity;
        }

        private static void SaveFileAndContents(string newFileAndPath, byte[] contents)
        {
            File.WriteAllBytes(newFileAndPath, contents);
        }

        private string GetFileAndPath(string fileName)
        {
            var basePath = _globalViewConfiguration.EasyLinkFileSavePath;
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
                model.Duration = Convert.ToDecimal(array[6]);
                model.Pages = Convert.ToInt32(array[7]);
                model.Charges = Convert.ToDecimal(array[8]);
                model.Messages = array[9];

                itemsToInsert.Add(model);
            }

            using (var bulkCopy = new SqlBulkCopy(_globalViewConfiguration.GlobalViewConnectionString))
            {
                bulkCopy.DestinationTableName = "EasyLinkData";
                bulkCopy.WriteToServer(itemsToInsert.ToDataTable());
            }
         }

        }
}