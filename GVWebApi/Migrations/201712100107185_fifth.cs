namespace GVWebapi.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class fifth : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.PeriodHistoryViews", "Rollover", c => c.Int());
        }
        
        public override void Down()
        {
            AlterColumn("dbo.PeriodHistoryViews", "Rollover", c => c.Decimal(precision: 18, scale: 2));
        }
    }
}
