using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using GVWebapi.Models;
using GVWebapi.RemoteData;
using System.Linq;
using System.Net.Mail;
using System.Configuration;
using System.IO;

namespace GVWebapi.Services
{
    public class MailParser
    {
 
 
 
        private static string ParseSupplyAssignment(String email,ServiceCallModel oSupplyInfo)
        {
            var freedomEntities = new CoFreedomEntities();
     
            try
            {



                if (oSupplyInfo.CallID != "!UNKNOWN!")
                {
                    
                    var oSupply = freedomEntities.vw_admin_EquipmentList_MeterGroup.Where(x => x.EquipmentID == oSupplyInfo.EquipmentID).FirstOrDefault();
                    var sActive = (oSupply.Active ? "Active" : "In-Active");
                    string oEmail = email.Replace("_#SERVICESUPPLY#_", "Supply ");
                    email = oEmail.Replace("_#ID#_", oSupplyInfo.CallID);
                    oEmail = email.Replace("_#Client Name#_", oSupply.CustomerName);
                    email = oEmail.Replace("_#SUPPLYPROVIDER#_", oSupply.SupplyProviderPrefFullName);
                    oEmail = email.Replace("_#REQUESTOR#_", oSupplyInfo.Name);
                    email = oEmail.Replace("_#DATETIME#_", DateTime.Now.ToShortDateString());
                    oEmail = email.Replace("_#EMAIL#_", oSupplyInfo.Email);
                    email = oEmail.Replace("_#TELEPHONE#_", oSupplyInfo.Phone);
                    oEmail = email.Replace("_#OPERATIONAL#_", oSupplyInfo.isWorking.ToString());
                    email = oEmail.Replace("_#HAVESUPPLIES#_", oSupplyInfo.isWorking.ToString());
                    oEmail = email.Replace("_#BLK#_", oSupplyInfo.Black.ToString());
                    email = oEmail.Replace("_#CYAN#_", oSupplyInfo.Cyan.ToString());
                    oEmail = email.Replace("_#YEL#_", oSupplyInfo.Yellow.ToString());
                    email = oEmail.Replace("_#MAG#_", oSupplyInfo.Magenta.ToString());
                    oEmail = email.Replace("_#ID#_", oSupplyInfo.CallID);
                    email = oEmail.Replace("_#DEVICEID#_", oSupply.EquipmentNumber);
                    oEmail = email.Replace("_#ISSUEID#_", oSupplyInfo.CallID);
                    email = oEmail.Replace("_#MODEL#_", oSupply.Model);
                    oEmail = email.Replace("_#SERIAL#_", oSupply.SerialNumber);
                    email = oEmail.Replace("_#DEVICETYPE#_", oSupply.ModelCategory);
                    oEmail = email.Replace("_#DEVICESTATUS#_", sActive);
                    email = oEmail.Replace("_#LOCATION#_", oSupply.LocName);
                    oEmail = email.Replace("_#ADDRESS#_", oSupply.Address);
                    email = oEmail.Replace("_#CITYSTATEZIP#_", oSupply.City + ", " + oSupply.State + " " + oSupply.Zip);
                    oEmail = email.Replace("_#DEPARTMENT#_", "");

                    email = oEmail.Replace("_#FLOORUSER#_", oSupply.Location);
                    oEmail = email.Replace("_#ISSUE#_", oSupplyInfo.CallID);
                    email = oEmail.Replace("_#COMMENTS#_", oSupplyInfo.Description);
                  
                    return email;

                }
                else
                {

                    string oEmail = email.Replace("_#SERVICESUPPLY#_", "Supply ");
                    email = oEmail.Replace("_#ID#_", oSupplyInfo.CallID);
                    oEmail = email.Replace("_#REQUESTOR#_", oSupplyInfo.Name);
                    email = oEmail.Replace("_#DATETIME#_", DateTime.Now.ToShortDateString());
                    oEmail = email.Replace("_#EMAIL#_", oSupplyInfo.Email);
                    email = oEmail.Replace("_#TELEPHONE#_", oSupplyInfo.Phone);
                    oEmail = email.Replace("_#OPERATIONAL#_", oSupplyInfo.isWorking.ToString());
                    email = oEmail.Replace("_#HAVESUPPLIES#_", oSupplyInfo.isWorking.ToString());
                    oEmail = email.Replace("_#BLK#_", oSupplyInfo.Black.ToString());
                    email = oEmail.Replace("_#CYAN#_", oSupplyInfo.Cyan.ToString());
                    oEmail = email.Replace("_#YEL#_", oSupplyInfo.Yellow.ToString());
                    email = oEmail.Replace("_#MAG#_", oSupplyInfo.Magenta.ToString());
                    oEmail = email.Replace("_#DEVICEID#_", oSupplyInfo.EquipmentNumber);
                    email = oEmail.Replace("_#ISSUEID#_", oSupplyInfo.CallID);
                    oEmail = email.Replace("_#DEPARTMENT#_", "");
                    email = oEmail.Replace("_#COMMENTS#_", oSupplyInfo.Description);
 
                    return email;
                }

            }
            finally
            {
                freedomEntities.Dispose();
            }
        }
        private static string ParseServiceAssignment(String email, ServiceCallModel oSupplyInfo)
        {
            var freedomEntities = new CoFreedomEntities();
           
            try
            {


                if (oSupplyInfo.CallID != "!UNKNOWN!")
                {

                   var equipQuery = freedomEntities.vw_admin_EquipmentList_MeterGroup.Where(x => x.EquipmentID == oSupplyInfo.EquipmentID).FirstOrDefault();
                   
                 


                    var sActive = (equipQuery.Active ? "Active" : "In-Active");
                    string oEmail = email.Replace("_#SERVICESUPPLY#_", "Service ");
                    email = oEmail.Replace("_#ID#_", oSupplyInfo.CallID);
                    oEmail = email.Replace("_#Client Name#_", equipQuery.CustomerName);
                    email = oEmail.Replace("_#SUPPLYPROVIDER#_", equipQuery.SupplyProviderPrefFullName);
                    oEmail = email.Replace("_#REQUESTOR#_", oSupplyInfo.Name);
                    email = oEmail.Replace("_#DATETIME#_", DateTime.Now.ToShortDateString());
                    oEmail = email.Replace("_#EMAIL#_", oSupplyInfo.Email);
                    email = oEmail.Replace("_#TELEPHONE#_", oSupplyInfo.Phone);
                    oEmail = email.Replace("_#OPERATIONAL#_", oSupplyInfo.isWorking.ToString());
                    email = oEmail.Replace("_#COMMENTS#_", oSupplyInfo.Description);           
                    oEmail = email.Replace("_#ISSUEID#_", oSupplyInfo.CallID);
                    email = oEmail.Replace("_#MODEL#_", equipQuery.Model);
                    oEmail = email.Replace("_#SERIAL#_", equipQuery.SerialNumber);
                    email = oEmail.Replace("_#DEVICETYPE#_", equipQuery.ModelCategory);
                    oEmail = email.Replace("_#DEVICESTATUS#_", sActive);
                    email = oEmail.Replace("_#LOCATION#_", equipQuery.LocName);
                    oEmail = email.Replace("_#ADDRESS#_", equipQuery.Address);
                    email = oEmail.Replace("_#CITYSTATEZIP#_", equipQuery.City + ", " + equipQuery.State + " " + equipQuery.Zip);
                    oEmail = email.Replace("_#DEPARTMENT#_", "");
                    email = oEmail.Replace("_#FLOORUSER#_", equipQuery.Location);
                    oEmail = email.Replace("_#ISSUEID#_", oSupplyInfo.CallID);
                    email = oEmail.Replace("_#IPADDRESS#_", equipQuery.IPAddress);
                    oEmail = email.Replace("_#DEVICEID#_", equipQuery.EquipmentNumber);
                    
                    return oEmail;
                }
                else
                {

                    string oEmail = email.Replace("_#SERVICESUPPLY#_", "Service ");
                    email = oEmail.Replace("_#ID#_", oSupplyInfo.CallID);
                    oEmail = email.Replace("_#REQUESTOR#_", oSupplyInfo.Name);
                    email = oEmail.Replace("_#DATETIME#_", DateTime.Now.ToShortDateString());
                    oEmail = email.Replace("_#EMAIL#_", oSupplyInfo.Email);
                    email = oEmail.Replace("_#TELEPHONE#_", oSupplyInfo.Phone);
                    oEmail = email.Replace("_#OPERATIONAL#_", oSupplyInfo.isWorking.ToString());
                    email = oEmail.Replace("_#COMMENTS#_", oSupplyInfo.Description);
                    oEmail = email.Replace("_#DEVICEID#_", oSupplyInfo.EquipmentNumber);
                    email = oEmail.Replace("_#ISSUEID#_", oSupplyInfo.CallID);
                    email = oEmail.Replace("_#DEPARTMENT#_", "");
                   

                    return oEmail;
                }
            }
            finally
            {
                freedomEntities.Dispose();
            }
        }
       
       public static bool EmailServiceCall(string id,ServiceCallModel model,int type)
        {
            var body = string.Empty;
            model.CallID = GetCallNumber(id);
            if(type == 1)
            {
                string filename = @"c:\inetpub\wwwroot\Service\views\ServiceRequest.htm";
                //string filename = @"d:\dev\repos\ServiceRequest.htm";
                //Get a StreamReader class that can be used to read the file 
                StreamReader objStreamReader = default(StreamReader);
                objStreamReader = File.OpenText(filename);

                //Now, read the entire file into a string 
                string contents = objStreamReader.ReadToEnd();
               body = ParseServiceAssignment(contents, model);
            }
            if (type == 2)
            {
                string filename = @"c:\inetpub\wwwroot\Supply\Views\SupplyRequest.htm";
               // string filename = @"d:\Dev\repos\SupplyRequest.htm";
                //Get a StreamReader class that can be used to read the file 
                StreamReader objStreamReader = default(StreamReader);
                objStreamReader = File.OpenText(filename);

                //Now, read the entire file into a string 
                string contents = objStreamReader.ReadToEnd();
               body =  ParseSupplyAssignment(contents, model);
            }

            MailMessage supplymail = new MailMessage();
            supplymail.From = new MailAddress(ConfigurationManager.AppSettings["FromAddress"]);
            supplymail.To.Add(new MailAddress(model.Email));
            supplymail.To.Add(new MailAddress(ConfigurationManager.AppSettings["SupportAddress"]));
            supplymail.Body = body;
            supplymail.IsBodyHtml = true;
            supplymail.Subject = "Freedom Profit Recovery Service Call # " + model.CallID + " Reported - " +
                                     DateTime.Now.ToShortDateString() + " @ " + DateTime.Now.ToShortTimeString();
            var client = new SmtpClient();
            client.EnableSsl = true;
            client.Send(supplymail);
            return true;
        }
        public static string GetCallNumber(string callId)
        {
            using (var freedomEntities = new CoFreedomEntities())
            {
               

                try
                {


                   var CallID = Int32.Parse(callId);
                   var oCall = freedomEntities.vw_ServiceCallandContact.Where(x => x.CallID == CallID).FirstOrDefault();
 
                    return oCall.CallNumber;

                }
                finally
                {
                    freedomEntities.Dispose();

                }
            }

        }
    }
 }
