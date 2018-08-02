using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using GVWebapi.RemoteData;
using GVWebapi.Models;

namespace GVWebapi.Controllers
{
    public class EquipmentManagersController : ApiController
    {
        private readonly GlobalViewEntities _globalViewEntities = new GlobalViewEntities();

        public IHttpActionResult GetEquipmentManagers()
        {
            var viewModel = new EquipmentManagerViewModel();
            viewModel.equimpmentList = _globalViewEntities.EquipmentManagerLists.ToList();
            viewModel.active = _globalViewEntities.EquipmentManagerLists.Count(c => c.Status == "Active");
            viewModel.inactive = _globalViewEntities.EquipmentManagerLists.Count(c => c.Status == "Inactive");
            viewModel.incomplete = _globalViewEntities.EquipmentManagerLists.Count(c => c.Status == "Incomplete");
            return Json(viewModel);
        }

        [ResponseType(typeof(EquipmentManager))]
        public async Task<IHttpActionResult> GetEquipmentManager(long id)
        {
            var equipmentManager = await _globalViewEntities.EquipmentManagers.FindAsync(id);
            if (equipmentManager == null)
            {
                return NotFound();
            }

            return Ok(equipmentManager);
        }

        [HttpGet,Route("api/equipmentmangermodel/{model}")]
        public async Task<IHttpActionResult> GetEquipmentManagerByModel(string model)
        {
            var equipmentManager = await _globalViewEntities.EquipmentManagers.Where(m => m.Model.Contains(model)).FirstOrDefaultAsync();
            if (equipmentManager == null)
            {
                return NotFound();
            }

            return Ok(equipmentManager);
        }

        [ResponseType(typeof(EquipmentManager))]
        public async Task<IHttpActionResult> PutEquipmentManager(long id, EquipmentManager equipmentManager)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != equipmentManager.Id)
            {
                return BadRequest();
            }

            _globalViewEntities.Entry(equipmentManager).State = EntityState.Modified;

            try
            {
                await _globalViewEntities.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!EquipmentManagerExists(id))
                {
                    return NotFound();
                }

                throw;
            }

            return Ok(equipmentManager);
        }

        [ResponseType(typeof(EquipmentManager))]
        public async Task<IHttpActionResult> PostEquipmentManager(EquipmentManager equipmentManager)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            _globalViewEntities.EquipmentManagers.Add(equipmentManager);
            await _globalViewEntities.SaveChangesAsync();

            return CreatedAtRoute("DefaultApi", new { id = equipmentManager.Id }, equipmentManager);
        }

        [ResponseType(typeof(EquipmentManager))]
        public async Task<IHttpActionResult> DeleteEquipmentManager(long id)
        {
            var equipmentManager = await _globalViewEntities.EquipmentManagers.FindAsync(id);
            if (equipmentManager == null)
            {
                return NotFound();
            }

            _globalViewEntities.EquipmentManagers.Remove(equipmentManager);
            await _globalViewEntities.SaveChangesAsync();

            return Ok(equipmentManager);
        }

        private bool EquipmentManagerExists(long id)
        {
            return _globalViewEntities.EquipmentManagers.Count(e => e.Id == id) > 0;
        }
    }
}