using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using GVWebapi.RemoteData;
namespace GVWebapi.Controllers
{
    public class ClientSettingsController : ApiController
    {
        private readonly GlobalViewEntities _globalViewEntities = new GlobalViewEntities();
        [HttpGet, Route("api/getclientsettings/{id}")]
        public IHttpActionResult getclientsettings(int id)
        {
            var settings = _globalViewEntities.ClientSettings.Where(x => x.ClientID == id).Select(x => new { key = x.SettingsType, value = x.SettingsValue }).ToList();
            return Json(settings);
        }
        [HttpPost, Route("api/setclientsetting")]
        public IHttpActionResult setclientsetting(ClientSetting model)
        {
            if (model == null) return NotFound();
            var settings = _globalViewEntities.ClientSettings.Where(x => x.ClientID == model.ClientID && x.SettingsType == model.SettingsType).FirstOrDefault();
            if(settings != null)
            {
                settings.ClientID = model.ClientID;
                settings.SettingsValue = model.SettingsValue;
                _globalViewEntities.SaveChanges();
            }
            else if(model.ClientID != 0)
            {
                ClientSetting setting = new ClientSetting();
                setting.ClientID = model.ClientID;
                setting.SettingsValue = model.SettingsValue;
                setting.SettingsType = model.SettingsType;
                _globalViewEntities.ClientSettings.Add(setting);
                _globalViewEntities.SaveChanges();
              
            }
           
            return Json(settings);
        }
    }
}
