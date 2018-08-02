namespace GVWebapi.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class first : DbMigration
    {
        public override void Up()
        {
            DropPrimaryKey("dbo.PeriodHistoryViews");
            AlterColumn("dbo.PeriodHistoryViews", "RevisionDataID", c => c.Long(nullable: false, identity: true));
            AlterColumn("dbo.PeriodHistoryViews", "ERPContractID", c => c.Long());
            AddPrimaryKey("dbo.PeriodHistoryViews", "RevisionDataID");
        }
        
        public override void Down()
        {
            DropPrimaryKey("dbo.PeriodHistoryViews");
            AlterColumn("dbo.PeriodHistoryViews", "ERPContractID", c => c.Int());
            AlterColumn("dbo.PeriodHistoryViews", "RevisionDataID", c => c.Int(nullable: false, identity: true));
            AddPrimaryKey("dbo.PeriodHistoryViews", "RevisionDataID");
        }
    }
}
