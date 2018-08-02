using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace GVWebapi.Models
{
    public class LoginData : ApiController
    {
        public string UserName { get; set; }
        public string Token { get; set; }
        public bool IsLoggedIn { get; set; }

    }
}
