using System;
using System.IO;
using GV.Configuration;
using GV.Domain.Entities;

namespace GV.Services
{
    public interface IEasyLinkFileSaveService
    {
        EasyLinkEntity SaveFile(EasyLinkFileSaveModel model);
    }

    public class EasyLinkFileSaveService : IEasyLinkFileSaveService
    {
        private readonly IGlobalViewConfiguration _globalViewConfiguration;

        public EasyLinkFileSaveService(IGlobalViewConfiguration globalViewConfiguration)
        {
            _globalViewConfiguration = globalViewConfiguration;
        }

        public EasyLinkEntity SaveFile(EasyLinkFileSaveModel model)
        {
            var newFileAndPath = GetFileAndPath(model.FileName);
            SaveFileAndContents(newFileAndPath, model.Contents);
            return CreateAndReturnEasyLinkEntity(newFileAndPath, model);
        }

        private EasyLinkEntity CreateAndReturnEasyLinkEntity(string newFileAndPath, EasyLinkFileSaveModel model)
        {
            var fileInfo = new FileInfo(newFileAndPath);
            var easyLinkEntity = new EasyLinkEntity();
            easyLinkEntity.FileName = model.FileName.Replace("\"", string.Empty);
            easyLinkEntity.SavedFileName = fileInfo.Name;
            easyLinkEntity.FileLocation = fileInfo.DirectoryName;
            easyLinkEntity.CreatedDateTime = DateTimeOffset.Now;
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
    }
}