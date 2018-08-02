namespace GVWebapi.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class forth : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.PeriodHistoryViews", "ERPContractID", c => c.Int());
        }
        
        public override void Down()
        {
            AlterColumn("dbo.PeriodHistoryViews", "ERPContractID", c => c.Long());
        }
    }
}
