using System;
using GV.Domain.Entities;

namespace GV.Services
{
    public class EasyLinkViewModel
    {
        public static EasyLinkViewModel For(EasyLinkEntity entity)
        {
            var model = new EasyLinkViewModel();
            model.EasyLinkId = entity.EasyLinkId;
            model.PeriodDate = entity.FileName.Replace(".csv", string.Empty);
            model.FileName = entity.FileName;
            model.NumberOfLines = entity.NumberOfLines;
            model.ImportedOn = entity.CreatedDateTime.LocalDateTime;
            return model;
        }

        public long EasyLinkId { get; set; }
        public string PeriodDate { get; set; }
        public string FileName { get; set; }
        public int NumberOfLines { get; set; }
        public DateTime ImportedOn { get; set; }
    }
}