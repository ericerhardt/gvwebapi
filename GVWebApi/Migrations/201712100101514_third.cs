namespace GVWebapi.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class third : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.PeriodHistoryViews", "ActualVolume", c => c.Long());
            AlterColumn("dbo.PeriodHistoryViews", "VolumeOffset", c => c.Long());
        }
        
        public override void Down()
        {
            AlterColumn("dbo.PeriodHistoryViews", "VolumeOffset", c => c.Decimal(precision: 18, scale: 2));
            AlterColumn("dbo.PeriodHistoryViews", "ActualVolume", c => c.Decimal(precision: 18, scale: 2));
        }
    }
}
