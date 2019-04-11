using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using GVWebapi.Models;
using GVWebapi.RemoteData;
 
using GVWebapi.Services;

namespace GVWebapi.Controllers
{
    public class ServiceCallController : ApiController
    {
        private CoFreedomEntities db = new CoFreedomEntities();

        

        // POST: api/ServiceCall
        [HttpPost, Route("api/bulkservice")]
        public async Task<IHttpActionResult> PostServiceCallModel(BulkCallModel[] serviceCallModel)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            List<ServiceCallModel> clientCallModel = new List<ServiceCallModel>();
            foreach (var call in serviceCallModel)
            {
             var Equipment =   db.vw_admin_EquipmentList_MeterGroup.Where(x => x.EquipmentNumber.Equals(call.DeviceId)).FirstOrDefault();
             if(Equipment != null)
                {
                    if (call.Isservice)
                    {
                     
                        ServiceCallModel oServiceInfo = new ServiceCallModel();
                        oServiceInfo.Name = call.Contact;
                        oServiceInfo.Phone = call.Phone;
                        oServiceInfo.Email = call.Email;
                        oServiceInfo.EquipmentID = Equipment.EquipmentID;
                        oServiceInfo.EquipmentNumber = call.DeviceId;
                        oServiceInfo.isWorking = call.IsWorking == "true" ? true : false;
                        oServiceInfo.Description = call.Issue;
                        oServiceInfo.Address = Equipment.Address;
                        oServiceInfo.User = Equipment.AssetUser;
                        oServiceInfo.Floor = Equipment.Floor;

                        var Id = "69874";//InsertServiceCall(oServiceInfo);
                       
                        oServiceInfo.CallNumber = Id;
                        oServiceInfo.CallType = 1;
                        clientCallModel.Add(oServiceInfo);
                        var success = BulkMailParser.EmailSupportServiceCall(Id, oServiceInfo, 1);
                    }
                    if (call.Issupply)
                    {
                        ServiceCallModel oSupplyInfo = new ServiceCallModel();
                        oSupplyInfo.Name = call.Contact;
                        oSupplyInfo.Phone = call.Phone;
                        oSupplyInfo.Email = call.Email;
                        oSupplyInfo.EquipmentID = Equipment.EquipmentID;
                        oSupplyInfo.EquipmentNumber = call.DeviceId;
                        oSupplyInfo.isWorking = call.HasSupplies == "true" ? true : false;
                        oSupplyInfo.Magenta = call.Magenta;
                        oSupplyInfo.Cyan = call.Cyan;
                        oSupplyInfo.Black = call.Blk;
                        oSupplyInfo.Yellow = call.Yellow;
                        oSupplyInfo.Description = call.Comment;
                        oSupplyInfo.Address = Equipment.Address;
                        oSupplyInfo.User = Equipment.AssetUser;
                        oSupplyInfo.Floor = Equipment.Floor;
                        var Id = InsertSupplyCall(oSupplyInfo);
                        oSupplyInfo.CallNumber = Id;
                        oSupplyInfo.CallType = 2;
                        clientCallModel.Add(oSupplyInfo);
                         var success = BulkMailParser.EmailSupportServiceCall(Id, oSupplyInfo, 2);
                    }
                }
            }
            var clientsuccess = BulkMailParser.EmailClientServiceCall(clientCallModel);

            return Ok("success");//CreatedAtRoute("DefaultApi", new { id = serviceCallModel.EquipmentID }, serviceCallModel);
        }

        

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

       
        public static string InsertSupplyCall(ServiceCallModel oSupplyInfo)
        {
            var db = new CoFreedomEntities();
            DbConnection con = db.Database.Connection;
            DbCommand cmd = con.CreateCommand();
            try
            {
                con.Open();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "Web_SCInsertServiceCall";

                DbParameter paramCaller = cmd.CreateParameter();
                paramCaller.ParameterName = "Caller";
                paramCaller.Value = oSupplyInfo.Name + "(" + oSupplyInfo.Phone + "/" + oSupplyInfo.Email + ")";
                cmd.Parameters.Add(paramCaller);

                DbParameter paramEquipid = cmd.CreateParameter();
                paramEquipid.ParameterName = "EquipmentID";
                paramEquipid.Value = oSupplyInfo.EquipmentID;
                cmd.Parameters.Add(paramEquipid);

                DbParameter paramDescription = cmd.CreateParameter();
                paramDescription.ParameterName = "Description";
                paramDescription.Value = "Supply Order For Device " + oSupplyInfo.EquipmentNumber + "\r\n Black =" + oSupplyInfo.Black + "\r\n Cyan=" + oSupplyInfo.Cyan + "\r\n Magenta=" + oSupplyInfo.Magenta + "\r\n Yellow=" + oSupplyInfo.Yellow + "\r\n Do they have supplies?:" + oSupplyInfo.isWorking.ToString() + "\r\n Comments:\r\n " + oSupplyInfo.Description;
                cmd.Parameters.Add(paramDescription);

                DbParameter paramCallType = cmd.CreateParameter();
                paramCallType.ParameterName = "CallTypeID";
                paramCallType.Value = 4;
                cmd.Parameters.Add(paramCallType);

                DbParameter paramCallId = cmd.CreateParameter();
                paramCallId.ParameterName = "CallID";
                paramCallId.DbType = DbType.Int32;
                paramCallId.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(paramCallId);

                cmd.ExecuteNonQuery();

                return cmd.Parameters["CallID"].Value.ToString();
            }
            catch (Exception)
            {
                return "!UNKNOWN!";
            }
            finally
            {
                cmd.Dispose();
                con.Dispose();
            }
        }
        public static string InsertServiceCall(ServiceCallModel oSupplyInfo)
        {
            var db = new CoFreedomEntities();
            DbConnection con = db.Database.Connection;
            DbCommand cmd = con.CreateCommand();
            try
            {
                con.Open();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "Web_SCInsertServiceCall";

                DbParameter paramCaller = cmd.CreateParameter();
                paramCaller.ParameterName = "Caller";

                paramCaller.Value = oSupplyInfo.Name + "(" + oSupplyInfo.Phone + "/" + oSupplyInfo.Email + ")";
                cmd.Parameters.Add(paramCaller);

                DbParameter paramEquipid = cmd.CreateParameter();
                paramEquipid.ParameterName = "EquipmentID";
                paramEquipid.Value = oSupplyInfo.EquipmentID;
                cmd.Parameters.Add(paramEquipid);

                DbParameter paramDescription = cmd.CreateParameter();
                paramDescription.ParameterName = "Description";
                paramDescription.Value = oSupplyInfo.Description;
                cmd.Parameters.Add(paramDescription);

                DbParameter paramCallType = cmd.CreateParameter();
                paramCallType.ParameterName = "CallTypeID";
                paramCallType.Value = 4;
                cmd.Parameters.Add(paramCallType);

                DbParameter paramCallId = cmd.CreateParameter();
                paramCallId.ParameterName = "CallID";
                paramCallId.DbType = DbType.Int32;
                paramCallId.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(paramCallId);

                cmd.ExecuteNonQuery();

                return cmd.Parameters["CallID"].Value.ToString();
            }
            catch (Exception ex)
            {
                return "!UNKNOWN!";
            }
            finally
            {
                cmd.Dispose();
                con.Dispose();
            }

        }
    }
}