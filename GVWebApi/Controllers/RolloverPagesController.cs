using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using GVWebapi.RemoteData;
using GVWebapi.Models;
using GVWebapi.Helpers;
namespace GVWebapi.Controllers
{
    public class RolloverPagesController : ApiController
    {
        private readonly ExcelRevisionExport db = new ExcelRevisionExport();

        private readonly CoFreedomEntities ea = new CoFreedomEntities();

        [HttpGet, Route("api/RolloverPages/{id}")]
        public IHttpActionResult GetRolloverView(int id)
        {
            var ContractID = db.GetContractID(id);
            var periods = db.GetRolloverHistory(ContractID).OrderByDescending(r => r.Period).ToList();
            var rollovers = ea.SCContractMeterGroups.Where(o => o.ContractID == ContractID).Select(o =>
              new RolloverUsageModel
              {
                  ContractID = o.ContractID,
                  ContractMeterGroupID = o.ContractMeterGroupID,
                  ContractMeterGroup = o.ContractMeterGroup,
                  RolloverUsage = o.RolloverUsage
              }).ToList();
            var totalSavings = periods.Sum(p=> p.TotalSavings);

            return Json(new { rollovers = periods, contractrollovers = rollovers, savings = totalSavings });
        }

        [HttpGet, Route("api/getmetergroups/{id}")]
        public IHttpActionResult getmetergrpups(int id)
        {

            var ContractID = db.GetContractID(id);
            var rollovers = ea.SCContractMeterGroups.Where(o=> o.ContractID == ContractID).Select(o=>
                new RolloverUsageModel
                {
                    ContractID = o.ContractID,
                    ContractMeterGroupID = o.ContractMeterGroupID,
                    ContractMeterGroup = o.ContractMeterGroup,
                    RolloverUsage = o.RolloverUsage
                }).ToList();
 

            return Json(rollovers);
        }
        [HttpPost, Route("api/updatemetergrouprollovers/")]
        public IHttpActionResult UpdateMeterGroupRollovers(IEnumerable<RolloverUsageModel> model)
        {
            foreach(var rollover in model)
            {
                var _rollover = ea.SCContractMeterGroups.Find(rollover.ContractMeterGroupID);
                if (_rollover != null)
                {
                    
                    _rollover.RolloverUsage = rollover.RolloverUsage;
                    ea.SaveChanges();

                }

            }


            return Ok();
        }

    }
}