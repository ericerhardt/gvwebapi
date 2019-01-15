using System;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using GVWebapi.RemoteData;
using GVWebapi.Models;
using System.Web;
using System.Net.Http;
using System.Net;
using System.Diagnostics;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.IO;
using GVWebapi.Helpers;

namespace GVWebapi.Controllers
{
    public class EquipmentManagersController : ApiController
    {
        private readonly GlobalViewEntities _globalViewEntities = new GlobalViewEntities();
        [HttpGet, Route("api/listequipment/")]
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
        [HttpGet, Route("api/getequipment/{id}")]
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
       
        public async Task<IHttpActionResult> EditEquipmentManager(long id, EquipmentManager equipmentManager)
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
        [HttpPost, Route("api/editequipment")]
        public async Task<IHttpActionResult> EditEquipment(EquipmentManager model)
        {
 

            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }
          
            if (model.Id == 0)
            {

                _globalViewEntities.EquipmentManagers.Add(model);
            }
            else
            {
              
              _globalViewEntities.Entry(model).State = EntityState.Modified;
            }
            await _globalViewEntities.SaveChangesAsync();
            return Ok(model);
        }

        [HttpPost, Route("api/editequipmentfile/")]
        public async Task<IHttpActionResult> EditEquipmentFile()
        {

            if (!Request.Content.IsMimeMultipartContent())
            {
                throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);
            }

            string root = HttpContext.Current.Server.MapPath("~/App_Data");
            var provider = await Request.Content.ReadAsMultipartAsync(new InMemoryMultipartFormDataStreamProvider());
            NameValueCollection formData = provider.FormData;
            // Read the form data.
           // await Request.Content.ReadAsMultipartAsync(provider);

            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }
            EquipmentManager model = new EquipmentManager();
            model.Model = String.IsNullOrEmpty(formData["Model"]) ? "" : formData["Model"];
            model.IntroDate = Convert.ToDateTime(formData["IntroDate"]);
            model.MFRMoVol = String.IsNullOrEmpty(formData["MFPMoVol"]) ? 0 : Convert.ToInt32(formData["MFPMoVol"]);
            model.PurchasePrice = String.IsNullOrEmpty(formData["PurchasePrice"]) ? 0 : Convert.ToDecimal(formData["PurchasePrice"]); 
            model.Image = formData["Image"];
            model.Status = Convert.ToInt32(formData["Status"]);
            model.LastUpdatedBy = formData["LastUpdatedBy"];
            model.LastUpdatedOn = DateTime.Now;
            if (formData["Id"] == "0")
            {

                _globalViewEntities.EquipmentManagers.Add(model);
            } else
            {
                model.Id = Convert.ToInt32(formData["Id"]);
                _globalViewEntities.Entry(model).State = EntityState.Modified;
            }
 
            //access files  
            IList<HttpContent> files = provider.Files;
            if (files.Count() > 0)
            {
                HttpContent file1 = files[0];
                var fileName = file1.Headers.ContentDisposition.FileName.Trim('\"');
                Stream input = await file1.ReadAsStreamAsync();


                var path = Path.Combine(HttpContext.Current.Server.MapPath("~/uploads/equipment/models"), fileName);
                //Deletion exists file  
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                using (Stream file = File.Create(path))
                {
                    input.CopyTo(file);
                    //close file  
                    file.Close();
                    var equipment = _globalViewEntities.EquipmentManagers.Find(model.Id);
                    if (equipment != null)
                    {
                        equipment.Image = fileName;
                       
                    }
                }

            }
            await _globalViewEntities.SaveChangesAsync();
            return Ok(model);
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