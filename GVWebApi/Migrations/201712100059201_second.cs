namespace GVWebapi.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class second : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.PeriodHistoryViews", "ContractedVolume", c => c.Long());
        }
        
        public override void Down()
        {
            AlterColumn("dbo.PeriodHistoryViews", "ContractedVolume", c => c.Decimal(precision: 18, scale: 2));
        }
    }
}
