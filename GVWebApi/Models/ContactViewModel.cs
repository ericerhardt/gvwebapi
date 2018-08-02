using System.Collections.Generic;
using GVWebapi.RemoteData;

namespace GVWebapi.Models
{
    public class ContactViewModel
    {
       public  IEnumerable<FprContact> FprContacts {get;set;}
       public  IEnumerable<ContactClientsViewModel> ContactClients { get; set; }
       public IEnumerable<vw_ClientsOnContract> Clients { get; set; }
       public FprContact FprContact { get; set; }
    }
}