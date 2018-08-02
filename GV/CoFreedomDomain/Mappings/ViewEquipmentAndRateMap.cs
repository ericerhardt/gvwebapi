using FluentNHibernate.Mapping;
using GV.CoFreedomDomain.Entities;

namespace GV.CoFreedomDomain.Mappings
{
    public class ViewEquipmentAndRateMap : ClassMap<ViewEquipmentAndRate>
    {
        public ViewEquipmentAndRateMap()
        {
            ReadOnly();
            Schema("dbo");
            Table("view_equipment_and_rate");

            Id(x => x.ViewKey).GeneratedBy.Assigned();

            Map(x => x.InvoiceId);
            Map(x => x.EquipmentNumber);
            Map(x => x.EquipmentSerialNumber);
            Map(x => x.DifferenceCopies);
            Map(x => x.ContractMeterGroup);
            Map(x => x.EffectiveRate);
            Map(x => x.StartDate);
            Map(x => x.EndDate);
        }
    }
}