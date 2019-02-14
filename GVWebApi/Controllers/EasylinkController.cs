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

namespace GVWebapi.Controllers
{
    public class EasylinkController : ApiController
    {
         

          private readonly GlobalViewEntities db = new GlobalViewEntities();
        private readonly EasyLinkServices _easyLinkService;

        // GET: api/Easylink
        public IHttpActionResult GetEasylinkImportHistories()
        {
            var results = db.EasylinkImportHistories.OrderByDescending(p => p.PeriodDate).AsEnumerable();
            return Json(results);
        }

        // GET: api/Easylink/5
        [ResponseType(typeof(EasylinkImportHistory))]
        public IHttpActionResult GetEasylinkImportHistory(int id)
        {
            EasylinkImportHistory easylinkImportHistory = db.EasylinkImportHistories.Find(id);
            if (easylinkImportHistory == null)
            {
                return NotFound();
            }

            return Ok(easylinkImportHistory);
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

        // DELETE: api/Easylink/5
        [ResponseType(typeof(EasylinkImportHistory))]
        public IHttpActionResult DeleteEasylinkImportHistory(int id)
        {
            EasylinkImportHistory easylinkImportHistory = db.EasylinkImportHistories.Find(id);
            if (easylinkImportHistory == null)
            {
                return NotFound();
            }

            db.EasylinkImportHistories.Remove(easylinkImportHistory);
            db.SaveChanges();

            return Ok(easylinkImportHistory);
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

            var provider = new MultipartMemoryStreamProvider();
            await Request.Content.ReadAsMultipartAsync(provider);

            foreach (var file in provider.Contents)
            {
                var fileName = file.Headers.ContentDisposition.FileName;
                var filesBytes = await file.ReadAsByteArrayAsync();
                var easyLinkEntity = _easyLinkService.SaveFile(new EasyLinkFileSaveModel(filesBytes, fileName));
                _easyLinkService.ImportData(easyLinkEntity);
            }

            return Ok();
        }
    }
}