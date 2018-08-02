using System.Collections.Generic;

namespace GVWebapi.Models.Schedules
{
    public class ScheduleEditModel
    {
        public SchedulesModel Schedule { get; set; } = new SchedulesModel();
        public IList<CoterminousModel> CoterminousModels { get; set; } = new List<CoterminousModel>();
    }
}