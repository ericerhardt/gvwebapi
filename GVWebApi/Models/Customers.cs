using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models
{
    public class Customer 
    {
        public int CustomerID { get; set; } 
        public string CustomerNumber { get; set; }
        public string CustomerName { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string County { get; set; }
        public string Zip   { get; set; }
        public string Country { get; set; }
        public decimal? Longitude { get; set; }
        public decimal? Latitude { get; set; }
        public string Phone1 { get; set; }
        public string Phone2 { get; set; }
        public string Email { get; set; }
        public string WebSite { get; set; }
    } 

     
}