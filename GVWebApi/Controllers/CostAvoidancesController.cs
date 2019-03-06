using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using GVWebapi.RemoteData;

namespace GVWebapi.Controllers
{
    public class CostAvoidancesController : ApiController
    {
        private readonly GlobalViewEntities _customerPortalEntities = new GlobalViewEntities();

        public IQueryable<CostAvoidance> GetCostAvoidances()
        {
            return _customerPortalEntities.CostAvoidances;
        }

        [HttpGet,Route("api/costavoidancebyclient/{idclient}")]
        public IQueryable<CostAvoidance> CostAvoidancesByClient(int idclient)
        {
            return _customerPortalEntities.CostAvoidances.Where(c => c.CustomerID == idclient);
        }

        // GET: api/CostAvoidances/5
        [ResponseType(typeof(CostAvoidance))]
        public async Task<IHttpActionResult> GetCostAvoidance(int id)
        {
            var costAvoidance = await _customerPortalEntities.CostAvoidances.FindAsync(id);
            if (costAvoidance == null)
            {
                return NotFound();
            }

            return Ok(costAvoidance);
        }

        // PUT: api/CostAvoidances/5
        [ResponseType(typeof(void))]
        public async Task<IHttpActionResult> PutCostAvoidance(int id, CostAvoidance costAvoidance)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != costAvoidance.CostAvoidanceID)
            {
                return BadRequest();
            }

            _customerPortalEntities.Entry(costAvoidance).State = System.Data.Entity.EntityState.Modified;

            try
            {
                await _customerPortalEntities.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!CostAvoidanceExists(id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return StatusCode(HttpStatusCode.NoContent);
        }

        // POST: api/CostAvoidances
        [ResponseType(typeof(CostAvoidance))]
        public async Task<IHttpActionResult> PostCostAvoidance(CostAvoidance costAvoidance)
        {
           

            if (CostAvoidanceExists(costAvoidance.CostAvoidanceID))
            {

                _customerPortalEntities.Entry(costAvoidance).State = System.Data.Entity.EntityState.Modified;
 
            } else{

                _customerPortalEntities.CostAvoidances.Add(costAvoidance);
            }
            
                await _customerPortalEntities.SaveChangesAsync();
             
            
            return CreatedAtRoute("DefaultApi", new { id = costAvoidance.CostAvoidanceID }, costAvoidance);
        }

         
        [HttpPost, Route("api/removecostavoidance/{id}")]
        public async Task<IHttpActionResult> DeleteCostAvoidance(int id)
        {
            var costAvoidance = await _customerPortalEntities.CostAvoidances.FindAsync(id);
            if (costAvoidance == null)
            {
                return NotFound();
            }

            _customerPortalEntities.CostAvoidances.Remove(costAvoidance);
            await _customerPortalEntities.SaveChangesAsync();

            return Ok(costAvoidance);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _customerPortalEntities.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool CostAvoidanceExists(int id)
        {
            return _customerPortalEntities.CostAvoidances.Count(e => e.CostAvoidanceID == id) > 0;
        }
    }
}