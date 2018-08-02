namespace GVWebapi.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class InitialCreate : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.PeriodHistoryViews",
                c => new
                    {
                        RevisionDataID = c.Int(nullable: false, identity: true),
                        ERPMeterGroupDesc = c.String(),
                        ContractedVolume = c.Decimal(precision: 18, scale: 2),
                        ActualVolume = c.Decimal(precision: 18, scale: 2),
                        CPPRate = c.Decimal(precision: 18, scale: 2),
                        VolumeOffset = c.Decimal(precision: 18, scale: 2),
                        PeriodDate = c.DateTime(),
                        Credits = c.Decimal(precision: 18, scale: 2),
                        FBRContractBase = c.Decimal(precision: 18, scale: 2),
                        ClientContractBase = c.Decimal(precision: 18, scale: 2),
                        ClientCPP = c.Decimal(precision: 18, scale: 2),
                        ERPContractID = c.Int(),
                        Rollover = c.Decimal(precision: 18, scale: 2),
                        OverageExpense = c.Decimal(precision: 18, scale: 2),
                        ERPMeterGroupID = c.Int(),
                    })
                .PrimaryKey(t => t.RevisionDataID);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.PeriodHistoryViews");
        }
    }
}
