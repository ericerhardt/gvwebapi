using System.Linq;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using GVWebapi.RemoteData;

namespace GVWebapi.Controllers
{
    public class GlobalViewUsersController : ApiController
    {
        private readonly GlobalViewEntities _globalViewEntities = new GlobalViewEntities();

        public IQueryable<GlobalViewUser> GetGlobalViewUsers()
        {
            return _globalViewEntities.GlobalViewUsers;
        }

        [HttpGet, Route("api/getglobalviewusers/getusersbyclientid/{idClient}")]
        public IQueryable<GlobalViewUser> GetUsersbyClientId(int idClient)
        {
            return _globalViewEntities.GlobalViewUsers.Where(g => g.idClient == idClient).AsQueryable();
        }

        [ResponseType(typeof(GlobalViewUser))]
        public async Task<IHttpActionResult> GetGlobalViewUser(int id)
        {
            GlobalViewUser globalViewUser = await _globalViewEntities.GlobalViewUsers.FindAsync(id);
            if (globalViewUser == null)
            {
                return NotFound();
            }

            return Ok(globalViewUser);
        }

        [HttpPost, Route("api/updateuserphone")]
        public async Task<IHttpActionResult> UpdateUserPhone(int id, string phone)
        {
            GlobalViewUser globalViewUser = await _globalViewEntities.GlobalViewUsers.FindAsync(id);
            if (globalViewUser == null)
            {
                return NotFound();
            }
            globalViewUser.Phone = phone;
            _globalViewEntities.SaveChanges();

            return Ok();
        }

        [ResponseType(typeof(GlobalViewUser))]
        public async Task<IHttpActionResult> PostGlobalViewUser(GlobalViewUser globalViewUser)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }
            if (!GlobalViewUserExists(globalViewUser.idUser))
            {
                _globalViewEntities.GlobalViewUsers.Add(globalViewUser);
                await _globalViewEntities.SaveChangesAsync();

                return CreatedAtRoute("DefaultApi", new { id = globalViewUser.idUser }, globalViewUser);
            }

            _globalViewEntities.Entry(globalViewUser).State = System.Data.Entity.EntityState.Modified;
            await _globalViewEntities.SaveChangesAsync();
            return Ok(globalViewUser);

        }

        [HttpPost, Route("api/DeleteGlobalViewUser/{id}"), ResponseType(typeof(GlobalViewUser))]
        public async Task<IHttpActionResult> DeleteGlobalViewUser(int id)
        {
            GlobalViewUser globalViewUser = await _globalViewEntities.GlobalViewUsers.FindAsync(id);
            if (globalViewUser == null)
            {
                return NotFound();
            }

            _globalViewEntities.GlobalViewUsers.Remove(globalViewUser);
            await _globalViewEntities.SaveChangesAsync();

            return Ok(globalViewUser);
        }

        private bool GlobalViewUserExists(int id)
        {
            return _globalViewEntities.GlobalViewUsers.Count(e => e.idUser == id) > 0;
        }
    }
}