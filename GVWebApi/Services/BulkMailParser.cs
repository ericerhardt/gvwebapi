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
    public class BulkMailParser
    {
       private static string ParseSupplyAssignment(String email,ServiceCallModel oSupplyInfo)
        {
            var freedomEntities = new CoFreedomEntities();
     
            try
            {



                if (oSupplyInfo.CallID != "Unlisted Device")
                {
                    
                    var oSupply = freedomEntities.vw_admin_EquipmentList_MeterGroup.Where(x => x.EquipmentID == oSupplyInfo.EquipmentID).FirstOrDefault();
                    var sActive = (oSupply.Active ? "Active" : "Inactive");
                    var EquipmentNumber = oSupply.VendorID != string.Empty ? $"{oSupply.EquipmentNumber} ({oSupply.VendorID})" : oSupply.EquipmentNumber;
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
                    oEmail = email.Replace("_#DEPARTMENT#_", oSupply.Department);
                    email = oEmail.Replace("_#USER#_", oSupply.AssetUser);
                    oEmail = email.Replace("_#ISSUE#_", oSupplyInfo.CallID);
                    email = oEmail.Replace("_#COMMENTS#_", oSupplyInfo.Description);
                    oEmail = email.Replace("_#FLOOR#_", oSupply.Floor);
                    return oEmail;

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
                    var sActive = (oSupply.Active ? "Active" : "Inactive");
                    var EquipmentNumber = oSupply.VendorID != string.Empty ? $"{oSupply.EquipmentNumber} ({oSupply.VendorID})" : oSupply.EquipmentNumber;
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
                    email = oEmail.Replace("_#DEVICEID#_", EquipmentNumber);
                    oEmail = email.Replace("_#ISSUEID#_", oSupplyInfo.CallID);
                    email = oEmail.Replace("_#MODEL#_", oSupply.Model);
                    oEmail = email.Replace("_#SERIAL#_", oSupply.SerialNumber);
                    email = oEmail.Replace("_#DEVICETYPE#_", oSupply.ModelCategory);
                    oEmail = email.Replace("_#DEVICESTATUS#_", sActive);
                    email = oEmail.Replace("_#LOCATION#_", oSupply.LocName);
                    oEmail = email.Replace("_#ADDRESS#_", oSupply.Address);
                    email = oEmail.Replace("_#CITYSTATEZIP#_", oSupply.City + ", " + oSupply.State + " " + oSupply.Zip);
                    oEmail = email.Replace("_#DEPARTMENT#_", oSupply.Department);
                    email = oEmail.Replace("_#FLOOR#_", oSupply.Floor);
                    oEmail = email.Replace("_#ISSUE#_", oSupplyInfo.CallID);
                    email = oEmail.Replace("_#COMMENTS#_", oSupplyInfo.Description);
                    oEmail = email.Replace("_#USER#_", oSupply.AssetUser);
                    return oEmail;

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
                    var sActive = (equipQuery.Active ? "Active" : "Inactive");
                    var EquipmentNumber = equipQuery.VendorID != string.Empty ? $"{equipQuery.EquipmentNumber} ({equipQuery.VendorID})" : equipQuery.EquipmentNumber;
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
                    oEmail = email.Replace("_#DEPARTMENT#_", equipQuery.Department);
                    email = oEmail.Replace("_#USER#_", equipQuery.AssetUser);
                    oEmail = email.Replace("_#ISSUEID#_", oSupplyInfo.CallID);
                    email = oEmail.Replace("_#IPADDRESS#_", equipQuery.IPAddress);
                    oEmail = email.Replace("_#DEVICEID#_", EquipmentNumber);
                    email = oEmail.Replace("_#FLOOR#_", equipQuery.Floor);
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
                    oEmail = email.Replace("_#OPERATIONAL#_", oSupplyInfo.isWorking ? "Yes" : "No");
                    email = oEmail.Replace("_#COMMENTS#_", oSupplyInfo.Description);
                    oEmail = email.Replace("_#DEVICEID#_", oSupplyInfo.EquipmentNumber);
                    email = oEmail.Replace("_#ISSUEID#_", oSupplyInfo.CallID);
                  
                   

                    return email;
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


                if (oSupplyInfo.CallID != "Unlisted Device")
                {

                    var equipQuery = freedomEntities.vw_admin_EquipmentList_MeterGroup.Where(x => x.EquipmentID == oSupplyInfo.EquipmentID).FirstOrDefault();
                    var EquipmentNumber = equipQuery.VendorID != string.Empty ? $"{equipQuery.EquipmentNumber} ({equipQuery.VendorID})" : equipQuery.EquipmentNumber;
                    var sActive = (equipQuery.Active ? "Active" : "Inactive");
                    string oEmail = email.Replace("_#SERVICESUPPLY#_", "Service ");
                    email = oEmail.Replace("_#ID#_", oSupplyInfo.CallID);
                    oEmail = email.Replace("_#CLIENTNAME#_", equipQuery.CustomerName);                   
                    email = oEmail.Replace("_#DATETIME#_", DateTime.Now.ToShortDateString());
                    oEmail = email.Replace("_#EMAIL#_", oSupplyInfo.Email);
                    email = oEmail.Replace("_#TELEPHONE#_", oSupplyInfo.Phone);
                    oEmail = email.Replace("_#OPERATIONAL#_", oSupplyInfo.IsWorking ? "Yes" : "No");
                    email = oEmail.Replace("_#COMMENTS#_", oSupplyInfo.Description);
                    oEmail = email.Replace("_#ISSUEID#_", oSupplyInfo.CallID);
                    email = oEmail.Replace("_#MODEL#_", equipQuery.Model);
                    oEmail = email.Replace("_#SERIAL#_", equipQuery.SerialNumber);
                    email = oEmail.Replace("_#DEVICETYPE#_", equipQuery.ModelCategory);
                    oEmail = email.Replace("_#DEVICESTATUS#_", sActive);
                    email = oEmail.Replace("_#LOCATION#_", equipQuery.LocName);
                    oEmail = email.Replace("_#ADDRESS#_", equipQuery.Address);
                    email = oEmail.Replace("_#CITYSTATEZIP#_", equipQuery.City + ", " + equipQuery.State + " " + equipQuery.Zip);
                    oEmail = email.Replace("_#DEPARTMENT#_", equipQuery.Department);
                    email = oEmail.Replace("_#FLOOR#_", equipQuery.Floor);
                    oEmail = email.Replace("_#USER#_", equipQuery.AssetUser);
                    email = oEmail.Replace("_#IPADDRESS#_", equipQuery.IPAddress);
                    oEmail = email.Replace("_#DEVICEID#_",  EquipmentNumber);
                    email = oEmail.Replace("_#REQUESTOR#_", oSupplyInfo.Name);
                    oEmail = email.Replace("_#ISSUEID#_", oSupplyInfo.CallID);
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
                    oEmail = email.Replace("_#OPERATIONAL#_", oSupplyInfo.IsWorking ? "Yes" : "No");
                    email = oEmail.Replace("_#COMMENTS#_", oSupplyInfo.Description);
                    oEmail = email.Replace("_#DEVICEID#_", oSupplyInfo.EquipmentNumber);
                    email = oEmail.Replace("_#ISSUEID#_", oSupplyInfo.CallID);
                    return email;
                }
            }
            finally
            {
                freedomEntities.Dispose();
            }
        }
        private static string ParseClientAssignment(String email, List<ServiceCallModel> oSupplyInfos)
        {
            var freedomEntities = new CoFreedomEntities();

            try
            {
                var ID = oSupplyInfos[0].EquipmentID;

                var equipQuery = freedomEntities.vw_admin_EquipmentList_MeterGroup.Where(x => x.EquipmentID == ID).FirstOrDefault();
                    var sActive = (equipQuery.Active ? "Active" : "Inactive");
                    var EquipmentNumber = equipQuery.VendorID != string.Empty ? $"{equipQuery.EquipmentNumber} ({equipQuery.VendorID})" : equipQuery.EquipmentNumber;
                    string oEmail = email.Replace("_#SERVICESUPPLY#_", "Service ");
                           email = oEmail.Replace("_#CLIENTNAME#_", equipQuery.CustomerName);
                           oEmail = email.Replace("_#REQUESTOR#_", oSupplyInfos[0].Name);
                           email = oEmail.Replace("_#DATETIME#_", DateTime.Now.ToShortDateString());
                           oEmail = email.Replace("_#EMAIL#_", oSupplyInfos[0].Email);
                           email = oEmail.Replace("_#TELEPHONE#_", oSupplyInfos[0].Phone);
                var CallList = string.Empty;
                foreach (var oSupplyInfo in oSupplyInfos)
                {
                    if (oSupplyInfo.CallType == 1)
                    {
                        CallList += "<tr><td>" + oSupplyInfo.CallID + "</td><td>" + oSupplyInfo.EquipmentNumber + "</td><td>" + oSupplyInfo.Description + "</td></tr>";
                    }
                    if (oSupplyInfo.CallType == 2)
                    {
                        var description = $"{oSupplyInfo.Description} Supplies Black:{oSupplyInfo.Black} Magenta: {oSupplyInfo.Magenta} Cyan: { oSupplyInfo.Cyan} Yellow: {oSupplyInfo.Yellow}";
                        CallList += "<tr><td>" + oSupplyInfo.CallID + "</td><td>" + oSupplyInfo.EquipmentNumber + "</td><td>" + description + "</td></tr>";
                    }
                     

                }
                oEmail = email.Replace("_#CALLLIST#_", CallList);

                return oEmail;
                 
            }
            finally
            {
                freedomEntities.Dispose();
            }
        }


        public static bool EmailSupportServiceCall(string id,ServiceCallModel model,int type)
        {
            var body = string.Empty;
            MailMessage supplymail = new MailMessage();
            supplymail.From = new MailAddress(ConfigurationManager.AppSettings["FromAddress"]);
            supplymail.To.Add(new MailAddress(ConfigurationManager.AppSettings["SupportAddress"]));
            supplymail.ReplyToList.Add(new MailAddress(ConfigurationManager.AppSettings["SupportAddress"]));
            if (type == 1)
            {
                var CallNumber = GetCallNumber(id);
                model.CallID = $"# {CallNumber}";
                 string filename = @"c:\inetpub\wwwroot\gvwebapi\templates\ServiceRequest.htm";
                //  string filename = @"D:\Dev\Repos\FPR\GVWebApi\GVWebapi\templates\ServiceRequest.htm";
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
                string filename = @"c:\inetpub\wwwroot\gvwebapi\templates\SupplyRequest.htm";
                // string filename = @"D:\Dev\Repos\FPR\GVWebApi\GVWebapi\templates\SupplyRequest.htm";
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
        public static bool EmailClientServiceCall(List<ServiceCallModel> models)
        {
            var body = string.Empty;
            MailMessage supplymail = new MailMessage();
            supplymail.From = new MailAddress(ConfigurationManager.AppSettings["FromAddress"]);
            supplymail.To.Add(new MailAddress(models[0].Email));
            supplymail.ReplyToList.Add(new MailAddress(ConfigurationManager.AppSettings["SupportAddress"]));
            
               
                string filename = @"c:\inetpub\wwwroot\gvwebapi\templates\ClientServiceRequest.htm";
                //  string filename = @"D:\Dev\Repos\FPR\GVWebApi\GVWebapi\templates\ClientServiceRequest.htm";
                if (!File.Exists(filename)) return false;
                StreamReader objStreamReader = default(StreamReader);

                objStreamReader = File.OpenText(filename);

                //Now, read the entire file into a string 
                string contents = objStreamReader.ReadToEnd();
                body = ParseClientAssignment(contents, models);
                supplymail.Subject = "Freedom Profit Recovery Service Call Summary" +
                                    DateTime.Now.ToShortDateString() + " @ " + DateTime.Now.ToShortTimeString();
            


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
            supplymail.ReplyToList.Add(new MailAddress(ConfigurationManager.AppSettings["SupportAddress"]));
            if (type == 1)
            {
                model.CallID = "# " + GetCallNumber(id);
                string filename = @"c:\inetpub\wwwroot\gvwebapi\templates\IntegrisServiceRequest.htm";
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
                string filename = @"c:\inetpub\wwwroot\gvwebapi\templates\IntegrisSupplyRequest.htm";
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
                string filename = @"c:\inetpub\wwwroot\gvwebapi\templates\ServiceRequest-ServiceNotFound.htm";
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
                string filename = @"c:\inetpub\wwwroot\gvwebapi\templates\SupplyRequest-SupplyNotFound.htm";
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
                   if(oCall != null)
                    {
                        return oCall.CallNumber;
                    } else
                    {
                        return "Call Not Added. Contact Support";
                    }
                   

                }
                finally
                {
                    freedomEntities.Dispose();

                }
            }

        }
    }
 }
