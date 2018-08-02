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

namespace GVWebapi.Controllers
{
    public class RolloverPagesController : ApiController
    {
        private RevisionDataEntities db = new RevisionDataEntities();

        public IQueryable<RolloverView> GetRolloverViews()
        {
            return db.RolloverViews;
        }

        [ResponseType(typeof(RolloverView))]
        public IHttpActionResult GetRolloverView(int id)
        {
            var periodDates = db.RolloverViews.OrderByDescending(r => r.PeriodDate).Where(r => r.CustomerID == id).Select(r => r.PeriodDate).Distinct().ToList();
            var values = new List<RolloverPagesModel>();
            foreach (var period in periodDates)
            {
                var ret = new RolloverPagesModel()
                {
                    Period = period,
                    Data = db.RolloverViews.Where(c => c.PeriodDate == period && c.CustomerID == id).Select(c => new Rollovers { ERPMeterGroupDesc = c.ERPMeterGroupDesc, Rollover = c.Rollover, CPP = c.CPP, Savings = (c.Rollover * c.CPP) }).ToList()
                };
                values.Add(ret);
            }
            var totalSavings = db.RolloverViews.Where(c => c.CustomerID == id).Sum(r => r.Rollover * r.CPP);

            return Json(new { data = values, savings = totalSavings });
        }

        [ResponseType(typeof(void))]
        public async Task<IHttpActionResult> PutRolloverView(string id, RolloverView rolloverView)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != rolloverView.ERPCustomerNumber)
            {
                return BadRequest();
            }

            db.Entry(rolloverView).State = EntityState.Modified;

            try
            {
                await db.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!RolloverViewExists(id))
                {
                    return NotFound();
                }

                throw;
            }

            return StatusCode(HttpStatusCode.NoContent);
        }

        [ResponseType(typeof(RolloverView))]
        public async Task<IHttpActionResult> PostRolloverView(RolloverView rolloverView)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.RolloverViews.Add(rolloverView);

            try
            {
                await db.SaveChangesAsync();
            }
            catch (DbUpdateException)
            {
                if (RolloverViewExists(rolloverView.ERPCustomerNumber))
                {
                    return Conflict();
                }

                throw;
            }

            return CreatedAtRoute("DefaultApi", new { id = rolloverView.ERPCustomerNumber }, rolloverView);
        }

        [ResponseType(typeof(RolloverView))]
        public async Task<IHttpActionResult> DeleteRolloverView(string id)
        {
            var rolloverView = await db.RolloverViews.FindAsync(id);
            if (rolloverView == null)
            {
                return NotFound();
            }

            db.RolloverViews.Remove(rolloverView);
            await db.SaveChangesAsync();

            return Ok(rolloverView);
        }

        private bool RolloverViewExists(string id)
        {
            return db.RolloverViews.Count(e => e.ERPCustomerNumber == id) > 0;
        }
    }
}