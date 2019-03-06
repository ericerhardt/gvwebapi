using System;
 
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Description;
using GVWebapi.RemoteData;
using GVWebapi.Services;
using GVWebapi.Models.Easylink;
using System.Threading.Tasks;
using System.Web;
using Newtonsoft.Json;
using System.IO;
using System.Net.Http.Headers;
using GV.Lookup;
 

namespace GVWebapi.Controllers
{
    public class EasylinkController : ApiController
    {


        private readonly GlobalViewEntities db = new GlobalViewEntities();
        IEasyLinkFileSaveService _easyLinkService = new EasyLinkServices();

        [HttpGet, Route("api/easylink")]
        public IHttpActionResult GetEasylinkImportHistories()
        {
            var results = db.EasylinkImportHistories.OrderByDescending(p => p.PeriodDate).AsEnumerable();
            return Json(results);
        }

        
        [HttpGet, Route("api/easylinkdetail/{id}")]
        [ResponseType(typeof(EasylinkImportHistory))]
        public async Task<IHttpActionResult> GetEasylinkImportHistory(int id)
        {
            EasylinkImportHistory easylinkImportHistory = await db.EasylinkImportHistories.FindAsync(id);
            if (easylinkImportHistory == null)
            {
                return NotFound();
            }
            var easylinkdetail = await db.EasylinkDatas.Where(x => x.ImportID == id).ToListAsync();
            
            return Ok( new { header = easylinkImportHistory, detail = easylinkdetail });
        }

        // PUT: api/Easylink/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutEasylinkImportHistory(int id, EasylinkImportHistory easylinkImportHistory)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != easylinkImportHistory.ImportID)
            {
                return BadRequest();
            }

            db.Entry(easylinkImportHistory).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!EasylinkImportHistoryExists(id))
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

        // POST: api/Easylink
        [ResponseType(typeof(EasylinkImportHistory))]
        public IHttpActionResult PostEasylinkImportHistory(EasylinkImportHistory easylinkImportHistory)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.EasylinkImportHistories.Add(easylinkImportHistory);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = easylinkImportHistory.ImportID }, easylinkImportHistory);
        }

        [HttpPost, Route("api/easylinkdelete/{id}")]
        [ResponseType(typeof(EasylinkImportHistory))]
        public async Task<IHttpActionResult> DeleteEasylinkImportHistory(int id)
        {
            EasylinkImportHistory  easylinkImportHistory = await db.EasylinkImportHistories.FindAsync(id);
            if (easylinkImportHistory == null)
            {
                return NotFound();
            }
            var easyLinkDatas = await db.EasylinkDatas.Where(x => x.ImportID == id).ToListAsync();
            if (easyLinkDatas != null)
            {
                db.EasylinkDatas.RemoveRange(easyLinkDatas);
            }
            db.EasylinkImportHistories.Remove(easylinkImportHistory);
            await db.SaveChangesAsync();
            var results = db.EasylinkImportHistories.OrderByDescending(p => p.PeriodDate).AsEnumerable();
            return Json(results);
         
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool EasylinkImportHistoryExists(int id)
        {
            return db.EasylinkImportHistories.Count(e => e.ImportID == id) > 0;
        }

        [HttpPost, Route("api/easylink/uploadfile")]
        public async Task<IHttpActionResult> UploadFile()
        {
            if (Request.Content.IsMimeMultipartContent() == false)
                throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);

            var provider = _easyLinkService.GetMultipartProvider();

            var result = await Request.Content.ReadAsMultipartAsync(provider);
            var EasyLinkObj = _easyLinkService.GetFormData(result);
            var originalFileName = _easyLinkService.GetDeserializedFileName(result.FileData.First());
            var uploadedFileInfo = new FileInfo(result.FileData.First().LocalFileName);
            EasyLinkObj.EasyLinkFile = new EasyLinkFileSaveModel(null, originalFileName);
            var easyLinkEntity = _easyLinkService.SaveFile(EasyLinkObj, originalFileName);
            _easyLinkService.ImportData(easyLinkEntity);

            return Ok();
        }

        [HttpGet, Route("api/easylink/getunmappedchild")]
        public IHttpActionResult GetUnMappedChild()
        {
            return Ok(_easyLinkService.GetUnMappedChildIds());
        }
        [HttpGet, Route("api/easylink/getallchild")]
        public IHttpActionResult GetChildActions(bool hideEasyLinkOnly)
        {
            return Ok(_easyLinkService.GetChildLinks(hideEasyLinkOnly));
        }
        [HttpPost, Route("api/easylink/loadaddchildmanagerforedit")]
        public IHttpActionResult GetChildManagerAddEditData(ChildManagerExistingData model)
        {
            var customers = _easyLinkService.GetAllCustomers(model.CustomerId);
            customers.Insert(0, new LookupInfo(0, "-- Select a Client --"));

            var childIds = _easyLinkService.GetAllEasyLinkChildIds(model.ChildId);
            childIds.Insert(0, new LookupInfo(0, "-- Select an Easylink ID --"));

            var viewModel = new
            {
                Customers = customers,
                ChildIds = childIds
            };
            return Ok(viewModel);
        }
        [HttpPost, Route("api/easylink/switchchildid")]
        public IHttpActionResult SwitchChild(EasyLinkChildSwitchModel model)
        {
            var mapping = db.EasyLinkMappings
                .Where(x => x.ChildId == model.ExistingChildId && x.ClientId == model.ExistingCustomerId)
                .FirstOrDefault();

            mapping.ChildId = model.NewChildId;
            mapping.ClientId = model.NewCustomerId;
            mapping.IsEasyLinkOnly = model.IsEasyLinkOnly;
            db.SaveChanges();
            return Ok();
        }

        [HttpPost, Route("api/easylink/addchildmatch")]
        public IHttpActionResult AddChildMatch(EasyLinkChildMatchSaveModel model)
        {
            _easyLinkService.AddEasyLinkChildMatch(model.CustomerId, model.ChildId, model.IsEasyLinkOnly);
           
            return Ok();
        }
        [HttpPost, Route("api/easylink/removelink")]
        public IHttpActionResult DeleteChildAction(int childId, int clientId)
        {
            _easyLinkService.RemoveLink(childId,clientId);
          
            return Ok();
        }
        [HttpGet, Route("api/easylink/loadaddchildmanager")]
        public IHttpActionResult GetChildManagerAddData()
        {
            var customers = _easyLinkService.GetAllCustomers();
            customers.Insert(0, new LookupInfo(0, "-- Select a Client --"));
            customers.Insert(1, new LookupInfo(999999, "EasyLink Only Customer"));

            var childIds = _easyLinkService.GetAllEasyLinkChildIds();
            childIds.Insert(0, new LookupInfo(0, "-- Select an Easylink ID --"));

            var viewModel = new
            {
                Customers = customers,
                ChildIds = childIds
            };
            return Ok(viewModel);
        }
    }    
}