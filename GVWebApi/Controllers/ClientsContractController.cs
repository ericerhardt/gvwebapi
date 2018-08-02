using System.Linq;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using GVWebapi.RemoteData;

namespace GVWebapi.Controllers
{
    public class ClientsContractController : ApiController
    {
        private readonly CoFreedomEntities _coFreedomEntities = new CoFreedomEntities();

        public  IHttpActionResult  Getvw_ClientsOnContract()
        {
            var clients =   _coFreedomEntities.vw_ClientsOnContract.OrderBy(c => c.CustomerName).AsEnumerable();
            return Json(clients);
        }

        [ResponseType(typeof(vw_ClientsOnContract))]
        public async Task<IHttpActionResult> Getvw_ClientsOnContract(string id)
        {
            var vwClientsOnContract = await _coFreedomEntities.vw_ClientsOnContract.FindAsync(id);
            if (vwClientsOnContract == null)
            {
                return NotFound();
            }

            return Json(vwClientsOnContract);
        }
 
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _coFreedomEntities.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}