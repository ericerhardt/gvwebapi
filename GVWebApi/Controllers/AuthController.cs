using System;
using System.Net;
using System.Web.Http;
using System.Web.Security;
using System.Linq;
using System.DirectoryServices.Protocols;
using GVWebapi.Models;
using GVWebapi.RemoteData;
using System.Configuration;
using System.Net.Mail;
namespace GVWebapi.Controllers
{
    public class AuthController : ApiController
    {
        private readonly GlobalViewEntities _globalViewEntities = new GlobalViewEntities();

        [HttpGet, Route("api/login/{UserName}/{Password}")]
        public IHttpActionResult Login(string userName, string password)
        {
            try
            {
                var loginData = new LoginData(); 
                var nc = new NetworkCredential( userName,  password, "FreedomProfit");
                var directoryIdentifier = new LdapDirectoryIdentifier("192.168.10.21", 389, false, false);
                var ldc = new LdapConnection(directoryIdentifier, nc, AuthType.Basic);


                ldc.Credential = nc;
                ldc.AuthType = AuthType.Negotiate;
                ldc.Bind(nc); // user has authenticated at this point, as the credentials were used to login to the dc.
               
                var authTicket = new FormsAuthenticationTicket(1,  // version
                                    userName,
                                    DateTime.Now,
                                    DateTime.Now.AddMinutes(60),
                                    false, "Domain Users");
                // Now encrypt the ticket.
                var encryptedTicket =
                  FormsAuthentication.Encrypt(authTicket);
                // Create a cookie and add the encrypted ticket to the
                // cookie as data.
                loginData.UserName =  userName;
                loginData.Token = encryptedTicket;
                loginData.IsLoggedIn = true;
                return Json(loginData);
            }
            catch (LdapException ex)
            {
                var result = new NotFoundWithMessageResult(ex.Message);
                return result;
            }
        }

        [HttpGet,Route("api/customerlogin/{username}/{password}/{keycode}")]
        public  IHttpActionResult CustomerLogin(string username, string password, string keycode)
        {
            try
            {
                var servercode = ConfigurationManager.AppSettings["keycode"];
                if (servercode != keycode)
                {
                    return BadRequest(ModelState);
                }
                var user = _globalViewEntities.GlobalViewUsers.FirstOrDefault(g => g.Email == username && g.Password == password);
                if (user != null)
                {
                    user.isLoggedIn = true;
                    user.logindatetime = DateTime.Now;
                    user.logoutdatetime = DateTime.Now.AddMinutes(2.0);
                    _globalViewEntities.SaveChanges();
                    return Ok(user);
                }
                return BadRequest("Error: Unable to authenticate your accuount.");
            }
            catch(Exception e)
            {
                return BadRequest("error:" + e.Message);
            }
            
        }
        [HttpPost, Route("api/customerlogout/{id}")]
        public IHttpActionResult CustomerLogout(int id)
        {
            try
            {
                var user = _globalViewEntities.GlobalViewUsers.Find(id);
                if (user != null)
                {
                    user.isLoggedIn = false;
                    user.logoutdatetime = DateTime.Now;
                    _globalViewEntities.SaveChanges();
                    return Ok(user);
                }
                return BadRequest("Error: Unable to find user your accuount.");
            }
            catch (Exception e)
            {
                return BadRequest("error:" + e.Message);
            }
        }
        [HttpPost]
        [Route("api/recover")]
        public IHttpActionResult RecoverPassword(string username)
        {
            try
            {
              
                GlobalViewUser user = _globalViewEntities.GlobalViewUsers.Where(g => g.Email == username).FirstOrDefault();
                if (user != null)
                {
                    string  mailuser =  ConfigurationManager.AppSettings["mailuser"];
                    string mailpwd = ConfigurationManager.AppSettings["mailpwd"];
                    string userInfo = "Your user name: " + user.Email + "<br/>";
                    userInfo += "Your password: " + user.Password;
                    MailMessage mail = new MailMessage();
                    mail.To.Add(user.Email);
                    mail.From = new MailAddress("no-reply@fprus.com", "GlobalView Administrator", System.Text.Encoding.UTF8);
                    mail.Subject = "Global View Log in information";
                    mail.SubjectEncoding = System.Text.Encoding.UTF8;
                    mail.Body = userInfo;
                    mail.BodyEncoding = System.Text.Encoding.UTF8;
                    mail.IsBodyHtml = true;
                    mail.Priority = MailPriority.High;
                    SmtpClient client = new SmtpClient();
                    client.Credentials = new  NetworkCredential(mailuser, mailpwd);
                    client.Port = 587;
                    client.Host = "mail026-2.exch026.serverdata.net";
                    client.EnableSsl = true;
                    try
                    {
                        client.Send(mail);
                      
                    }
                    catch (Exception ex)
                    {
                        Exception ex2 = ex;
                        string errorMessage = string.Empty;
                        while (ex2 != null)
                        {
                            errorMessage += ex2.ToString();
                            ex2 = ex2.InnerException;
                        }
                       
                    }
                    return Ok(user);
                }
                return BadRequest("Error: Unable to recover your accuount.");
            }
            catch (Exception e)
            {
                return BadRequest("error:" + e.Message);
            }

        }

    }
}
