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



                if (oSupplyInfo.CallID != "Unlisted Device")
                {
                    
                    var oSupply = freedomEntities.vw_admin_EquipmentList_MeterGroup.Where(x => x.EquipmentID == oSupplyInfo.EquipmentID).FirstOrDefault();
                    var sActive = (oSupply.Active ? "Active" : "In-Active");
                    string oEmail = email.Replace("_#SERVICESUPPLY#_", "Supply ");
                    email = oEmail.Replace("_#ID#_", oSupplyInfo.CallID);
                    oEmail = email.Replace("_#CLIENTNAME#_", oSupply.CustomerName);
                    email = oEmail.Replace("_#SUPPLYPROVIDER#_", oSupply.SupplyProviderPrefFullName);
                    oEmail = email.Replace("_#REQUESTOR#_", oSupplyInfo.Name);
                    email = oEmail.Replace("_#DATETIME#_", DateTime.Now.ToShortDateString());
                    oEmail = email.Replace("_#EMAIL#_", oSupplyInfo.Email);
                    email = oEmail.Replace("_#TELEPHONE#_", oSupplyInfo.Phone);
                    oEmail = email.Replace("_#OPERATIONAL#_", oSupplyInfo.isWorking ? "Yes" : "No");
                    email = oEmail.Replace("_#HAVESUPPLIES#_", oSupplyInfo.isWorking ? "Yes" : "No");
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
                    email = oEmail.Replace("_#ID#_", "Unlisted Device");
                    oEmail = email.Replace("_#REQUESTOR#_", oSupplyInfo.Name);
                    email = oEmail.Replace("_#DATETIME#_", DateTime.Now.ToShortDateString());
                    oEmail = email.Replace("_#EMAIL#_", oSupplyInfo.Email);
                    email = oEmail.Replace("_#TELEPHONE#_", oSupplyInfo.Phone);
                    oEmail = email.Replace("_#OPERATIONAL#_", oSupplyInfo.isWorking ? "Yes" : "No");
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

        private static string ParseIntegrisSupplyAssignment(String email, IntegrisServiceCallModel oSupplyInfo)
        {
            var freedomEntities = new CoFreedomEntities();

            try
            {



                if (oSupplyInfo.CallID != "Unlisted Device")
                {

                    var oSupply = freedomEntities.vw_admin_EquipmentList_MeterGroup.Where(x => x.EquipmentID == oSupplyInfo.EquipmentID).FirstOrDefault();
                    var sActive = (oSupply.Active ? "Active" : "In-Active");
                    string oEmail = email.Replace("_#SERVICESUPPLY#_", "Supply ");
                    email = oEmail.Replace("_#ID#_", oSupplyInfo.CallID);
                    oEmail = email.Replace("_#CLIENTNAME#_", oSupply.CustomerName);
                    email = oEmail.Replace("_#SUPPLYPROVIDER#_", oSupply.SupplyProviderPrefFullName);
                    oEmail = email.Replace("_#REQUESTOR#_", oSupplyInfo.Name);
                    email = oEmail.Replace("_#DATETIME#_", DateTime.Now.ToShortDateString());
                    oEmail = email.Replace("_#EMAIL#_", oSupplyInfo.Email);
                    email = oEmail.Replace("_#TELEPHONE#_", oSupplyInfo.Phone);
                    oEmail = email.Replace("_#OPERATIONAL#_", oSupplyInfo.IsWorking ? "Yes" : "No");
                    email = oEmail.Replace("_#HAVESUPPLIES#_", oSupplyInfo.IsWorking ? "Yes" : "No");
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
                    email = oEmail.Replace("_#ID#_", "Unlisted Device");
                    oEmail = email.Replace("_#REQUESTOR#_", oSupplyInfo.Name);
                    email = oEmail.Replace("_#DATETIME#_", DateTime.Now.ToShortDateString());
                    oEmail = email.Replace("_#EMAIL#_", oSupplyInfo.Email);
                    email = oEmail.Replace("_#TELEPHONE#_", oSupplyInfo.Phone);
                    oEmail = email.Replace("_#OPERATIONAL#_", oSupplyInfo.IsWorking ? "Yes" : "No");
                    email = oEmail.Replace("_#HAVESUPPLIES#_", oSupplyInfo.IsWorking ? "Yes" : "No");
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


                if (oSupplyInfo.CallID != "Unlisted Device")
                {

                   var equipQuery = freedomEntities.vw_admin_EquipmentList_MeterGroup.Where(x => x.EquipmentID == oSupplyInfo.EquipmentID).FirstOrDefault();
                   
                 


                    var sActive = (equipQuery.Active ? "Active" : "In-Active");
                    string oEmail = email.Replace("_#SERVICESUPPLY#_", "Service ");
                           email = oEmail.Replace("_#ID#_", oSupplyInfo.CallID);
                           oEmail = email.Replace("_#CLIENTNAME#_", equipQuery.CustomerName);
                           email = oEmail.Replace("_#SUPPLYPROVIDER#_", equipQuery.SupplyProviderPrefFullName);
                           oEmail = email.Replace("_#REQUESTOR#_", oSupplyInfo.Name);
                           email = oEmail.Replace("_#DATETIME#_", DateTime.Now.ToShortDateString());
                    oEmail = email.Replace("_#EMAIL#_", oSupplyInfo.Email);
                    email = oEmail.Replace("_#TELEPHONE#_", oSupplyInfo.Phone);
                    oEmail = email.Replace("_#OPERATIONAL#_", oSupplyInfo.isWorking ? "Yes" : "No");
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
                    email = oEmail.Replace("_#ID#_", "Unlisted Device");
                    oEmail = email.Replace("_#REQUESTOR#_", oSupplyInfo.Name);
                    email = oEmail.Replace("_#DATETIME#_", DateTime.Now.ToShortDateString());
                    oEmail = email.Replace("_#EMAIL#_", oSupplyInfo.Email);
                    email = oEmail.Replace("_#TELEPHONE#_", oSupplyInfo.Phone);
                    oEmail = email.Replace("_#OPERATIONAL#_", oSupplyInfo.isWorking ? "Yes" : "No");
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
       private static string ParseIntegrisServiceAssignment(String email, IntegrisServiceCallModel oSupplyInfo)
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
                    oEmail = email.Replace("_#CLIENTNAME#_", equipQuery.CustomerName);                   
                    email = oEmail.Replace("_#DATETIME#_", DateTime.Now.ToShortDateString());
                    oEmail = email.Replace("_#EMAIL#_", oSupplyInfo.Email);
                    email = oEmail.Replace("_#TELEPHONE#_", oSupplyInfo.Phone);
                    oEmail = email.Replace("_#OPERATIONAL#_", oSupplyInfo.IsWorking ? "Yes" : "No");
                    oEmail = email.Replace("_#CLIENTCARE#_", oSupplyInfo.PatientCare ? "Yes" : "No");
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
                    email = oEmail.Replace("_#REQUESTOR#_", oSupplyInfo.Name);
                    return email;
                }
                else
                {

                    string oEmail = email.Replace("_#SERVICESUPPLY#_", "Service ");
                    email = oEmail.Replace("_#ID#_", "Unlisted Device");
                    oEmail = email.Replace("_#REQUESTOR#_", oSupplyInfo.Name);
                    email = oEmail.Replace("_#DATETIME#_", DateTime.Now.ToShortDateString());
                    oEmail = email.Replace("_#EMAIL#_", oSupplyInfo.Email);
                    email = oEmail.Replace("_#TELEPHONE#_", oSupplyInfo.Phone);
                    oEmail = email.Replace("_#OPERATIONAL#_", oSupplyInfo.IsWorking ? "Yes" : "No");
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
            MailMessage supplymail = new MailMessage();
            supplymail.From = new MailAddress(ConfigurationManager.AppSettings["FromAddress"]);
            supplymail.To.Add(new MailAddress(model.Email));
            supplymail.To.Add(new MailAddress(ConfigurationManager.AppSettings["SupportAddress"]));

            if (type == 1)
            {
                var CallNumber = GetCallNumber(id);
                model.CallID = $"# {CallNumber}";
                string filename = @"c:\inetpub\wwwroot\gvadmin\service\ServiceRequest.htm";
                // string filename = @"d:\Dev\repos\ServiceRequest.htm";
                if (!File.Exists(filename)) return false;
                StreamReader objStreamReader = default(StreamReader);
              
                objStreamReader = File.OpenText(filename);

                //Now, read the entire file into a string 
                string contents = objStreamReader.ReadToEnd();
               body = ParseServiceAssignment(contents, model);
               supplymail.Subject = "Freedom Profit Recovery Service Call " + model.CallID + " Reported - " +
                                   DateTime.Now.ToShortDateString() + " @ " + DateTime.Now.ToShortTimeString();
            }
            if (type == 2)
            {
                var CallNumber = GetCallNumber(id);
                model.CallID = $"# {CallNumber}";
                string filename = @"c:\inetpub\wwwroot\gvadmin\service\SupplyRequest.htm";
                // string filename = @"d:\Dev\repos\SupplyRequest.htm";
                if (!File.Exists(filename)) return false;
                StreamReader objStreamReader = default(StreamReader);
                objStreamReader = File.OpenText(filename);

                //Now, read the entire file into a string 
                string contents = objStreamReader.ReadToEnd();
               body =  ParseSupplyAssignment(contents, model);
               supplymail.Subject = "Freedom Profit Recovery Supply Call " + model.CallID + " Reported - " +
                                   DateTime.Now.ToShortDateString() + " @ " + DateTime.Now.ToShortTimeString();
            }
            if (type == 3)
            {
                model.CallID = "Unlisted Device";
                string filename = @"c:\inetpub\wwwroot\gvadmin\service\ServiceRequest-ServiceNotFound.htm";
                //string filename = @"D:\Dev\Projects\ServiceAndSupply\Supplies.EA.Websolution\Views\ServiceRequest-ServiceNotFound.htm";
                if (!File.Exists(filename)) return false;
                // string filename = @"d:\Dev\repos\SupplyRequest.htm";
                //Get a StreamReader class that can be used to read the file 
                StreamReader objStreamReader = default(StreamReader);
                objStreamReader = File.OpenText(filename);
                //Now, read the entire file into a string 
                string contents = objStreamReader.ReadToEnd();
                body = ParseServiceAssignment(contents, model);
                supplymail.Subject = "**Alert** Freedom Profit Recovery Service Call " + model.CallID + " Reported - " +
                                   DateTime.Now.ToShortDateString() + " @ " + DateTime.Now.ToShortTimeString();
            }
            if (type == 4)
            {
                model.CallID = "Unlisted Device";
                string filename = @"c:\inetpub\wwwroot\gvadmin\service\SupplyRequest-SupplyNotFound.htm";
                //string filename = @"D:\Dev\Projects\ServiceAndSupply\Supplies.EA.Websolution\Views\SupplyRequest-SupplyNotFound.htm";
                if (!File.Exists(filename)) return false;
                // string filename = @"d:\Dev\repos\SupplyRequest.htm";
                //Get a StreamReader class that can be used to read the file 
                StreamReader objStreamReader = default(StreamReader);
                objStreamReader = File.OpenText(filename);

                //Now, read the entire file into a string 
                string contents = objStreamReader.ReadToEnd();
                body = ParseSupplyAssignment(contents, model);
                supplymail.Subject = "**Alert** Freedom Profit Recovery Supply Call " + model.CallID + " Reported - " +
                                   DateTime.Now.ToShortDateString() + " @ " + DateTime.Now.ToShortTimeString();
            }

           
            supplymail.Body = body;
            supplymail.IsBodyHtml = true;
          
            var client = new SmtpClient();
            client.EnableSsl = true;
            client.Send(supplymail);
            return true;
        }
       public static bool EmailIntegrisServiceCall(string id, IntegrisServiceCallModel model, int type)
        {
            var body = string.Empty;
            MailMessage supplymail = new MailMessage();
            supplymail.From = new MailAddress(ConfigurationManager.AppSettings["FromAddress"]);
            supplymail.To.Add(new MailAddress(model.Email));
            supplymail.To.Add(new MailAddress(ConfigurationManager.AppSettings["SupportAddress"]));

            if (type == 1)
            {
                model.CallID = "# " + GetCallNumber(id);
                string filename = @"c:\inetpub\wwwroot\gvadmin\integris\ServiceRequest.htm";
                //string filename = @"d:\dev\repos\IntegrisServiceRequest.htm";
                //Get a StreamReader class that can be used to read the file 
                StreamReader objStreamReader = default(StreamReader);
                objStreamReader = File.OpenText(filename);
                if (!File.Exists(filename)) return false;
                objStreamReader = File.OpenText(filename);
                //Now, read the entire file into a string 
                string contents = objStreamReader.ReadToEnd();
                body = ParseIntegrisServiceAssignment(contents, model);
                supplymail.Subject = "Freedom Profit Recovery Service Call " + model.CallID + " Reported - " +
                                   DateTime.Now.ToShortDateString() + " @ " + DateTime.Now.ToShortTimeString();
            }
            if (type == 2)
            {
                model.CallID = "# " + GetCallNumber(id);
                string filename = @"c:\inetpub\wwwroot\gvadmin\integris\SupplyRequest.htm";
               // string filename = @"d:\Dev\repos\IntegrisSupplyRequest.htm";
                //Get a StreamReader class that can be used to read the file 
                StreamReader objStreamReader = default(StreamReader);
                if (!File.Exists(filename)) return false;
                objStreamReader = File.OpenText(filename);

                //Now, read the entire file into a string 
                string contents = objStreamReader.ReadToEnd();
                body = ParseIntegrisSupplyAssignment(contents, model);
                supplymail.Subject = "Freedom Profit Recovery Supply Call " + model.CallID + " Reported - " +
                                   DateTime.Now.ToShortDateString() + " @ " + DateTime.Now.ToShortTimeString();
            }
            if (type == 3)
            {
                model.CallID = "Unlisted Device";
                string filename = @"c:\inetpub\wwwroot\gvadmin\integris\ServiceRequest-ServiceNotFound.htm";
                // string filename = @"D:\Dev\Projects\ServiceAndSupply\Supplies.EA.Websolution\Views\ServiceRequest-ServiceNotFound.htm";
                if (!File.Exists(filename)) return false;
                // string filename = @"d:\Dev\repos\SupplyRequest.htm";
                //Get a StreamReader class that can be used to read the file 
                StreamReader objStreamReader = default(StreamReader);
                objStreamReader = File.OpenText(filename);
                //Now, read the entire file into a string 
                string contents = objStreamReader.ReadToEnd();
                body = ParseIntegrisServiceAssignment(contents, model);
                supplymail.Subject = "**Alert** Freedom Profit Recovery Service Call " + model.CallID + " Reported - " +
                                   DateTime.Now.ToShortDateString() + " @ " + DateTime.Now.ToShortTimeString();
            }
            if (type == 4)
            {
                model.CallID = "Unlisted Device";
                string filename = @"c:\inetpub\wwwroot\gvadmin\integris\SupplyRequest-SupplyNotFound.htm";
                //string filename = @"D:\Dev\Projects\ServiceAndSupply\Supplies.EA.Websolution\Views\SupplyRequest-SupplyNotFound.htm";
                if (!File.Exists(filename)) return false;
                // string filename = @"d:\Dev\repos\SupplyRequest.htm";
                //Get a StreamReader class that can be used to read the file 
                StreamReader objStreamReader = default(StreamReader);
                objStreamReader = File.OpenText(filename);

                //Now, read the entire file into a string 
                string contents = objStreamReader.ReadToEnd();
                body = ParseIntegrisSupplyAssignment(contents, model);
                supplymail.Subject = "**Alert** Freedom Profit Recovery Supply Call " + model.CallID + " Reported - " +
                                   DateTime.Now.ToShortDateString() + " @ " + DateTime.Now.ToShortTimeString();
            }

            
            supplymail.Body = body;
            supplymail.IsBodyHtml = true;
          
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
