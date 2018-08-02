using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using GVWebapi.RemoteData;
using GVWebapi.Models;

namespace GVWebapi.Controllers
{
    public class FprContactController : ApiController
    {
        private readonly CoFreedomEntities _coFreedomEntities = new CoFreedomEntities();
        private readonly GlobalViewEntities _globalViewEntities = new GlobalViewEntities();

        [HttpGet, Route("api/contacts")]
        public IHttpActionResult Clients()
        {
            var viewModel = new ContactViewModel();
             viewModel.FprContacts = _globalViewEntities.FprContacts.ToList();
            return Ok(viewModel);
        }

        [HttpGet, Route("api/contact/{id}")]
        public IHttpActionResult GetContact(int id)
        {
            var contact = _globalViewEntities.FprContacts.Where(c => c.ContactId == id);
            return Ok(contact);
        }

        [HttpGet, Route("api/clientcontacts/{CustomerID}")]
        public IHttpActionResult GetClientContacts(int customerId)
        {
            var contacts = _globalViewEntities.FprContacts.Join(_globalViewEntities.ContactClients, client => client.ContactId, contact => contact.FprContactId, (client, contact) => new { FprContact = client, ClientContact = contact }).Where(contact => contact.ClientContact.CustomerId == customerId);
            return Ok(contacts.Select(x => x.FprContact));
        }

        [HttpGet, Route("api/contactclients/{Id}")]
        public IHttpActionResult ContactClients(int Id)
        {
            var viewModel = new ContactViewModel();
            viewModel.FprContact = _globalViewEntities.FprContacts.FirstOrDefault(c => c.ContactId == Id);
            var contactClients = new List<ContactClientsViewModel>();
            var clientList = _coFreedomEntities
                .vw_ClientsOnContract
                .Select(item => new { ContactClientId = 0, FprContactId = 0, item.CustomerID, item.CustomerName, Selected = false })
                .ToList();

            if (viewModel.FprContact != null)
            {
                foreach (var client in clientList)
                {
                    var contacts = _globalViewEntities.ContactClients.FirstOrDefault(c => c.CustomerId == client.CustomerID && c.FprContactId == Id);

                    if (contacts != null)
                    {
                        var contact = new ContactClientsViewModel(contacts.ContactClientId, contacts.FprContactId, contacts.CustomerId, client.CustomerName, true);
                        contactClients.Add(contact);
                    }
                    else
                    {
                        var contact = new ContactClientsViewModel(client.ContactClientId, client.FprContactId, client.CustomerID, client.CustomerName, client.Selected);
                        contactClients.Add(contact);
                    }
                }

            }
            else
            {
                foreach (var client in clientList)
                {
                    var contact = new ContactClientsViewModel(client.ContactClientId, 0, client.CustomerID, client.CustomerName, client.Selected);
                    contactClients.Add(contact);
                }
            }
            viewModel.ContactClients = contactClients;
            return Ok(viewModel);
        }

        [HttpPost, Route("api/addcontact/")]
        public IHttpActionResult ContactClients(ContactViewModel contact)
        {
            if (contact != null)
            {
                var viewModel = new ContactViewModel();
                var fprcontact = new FprContact();
                fprcontact.FirstName = contact.FprContact.FirstName;
                fprcontact.LastName = contact.FprContact.LastName;
                fprcontact.Title = contact.FprContact.Title;
                fprcontact.StartedDate = contact.FprContact.StartedDate;
                fprcontact.Email = contact.FprContact.Email;
                fprcontact.LinkedIn = contact.FprContact.LinkedIn;
                fprcontact.Story = contact.FprContact.Story;
                fprcontact.Photo = contact.FprContact.Photo;
                _globalViewEntities.FprContacts.Add(fprcontact);
                _globalViewEntities.SaveChanges();

                foreach (var client in contact.ContactClients)
                {
                    if (client.Selected)
                    {
                        var clientContact = new ContactClient();
                        clientContact.FprContactId = fprcontact.ContactId;
                        clientContact.CustomerId = client.CustomerId;
                        _globalViewEntities.ContactClients.Add(clientContact);
                        _globalViewEntities.SaveChanges();
                    }
                }
                viewModel.FprContacts = _globalViewEntities.FprContacts.ToList();
                viewModel.ContactClients = contact.ContactClients;
                return Ok(viewModel);
            }
            return NotFound();
        }

        [HttpPost, Route("api/editcontact/")]
        public IHttpActionResult EditContact(ContactViewModel contact)
        {
            if (contact != null)
            {
                var viewModel = new ContactViewModel();
                var fprcontact = _globalViewEntities.FprContacts.Find(contact.FprContact.ContactId);
                if (fprcontact == null) return BadRequest();
                fprcontact.FirstName = contact.FprContact.FirstName;
                fprcontact.LastName = contact.FprContact.LastName;
                fprcontact.Title = contact.FprContact.Title;
                fprcontact.StartedDate = contact.FprContact.StartedDate;

                fprcontact.Email = contact.FprContact.Email;
                fprcontact.LinkedIn = contact.FprContact.LinkedIn;
                fprcontact.Story = contact.FprContact.Story;
                if (contact.FprContact.Photo != null)
                    fprcontact.Photo = contact.FprContact.Photo;
                _globalViewEntities.SaveChanges();

                foreach (var client in contact.ContactClients)
                {

                    var editcontact = _globalViewEntities.ContactClients.Find(client.ContactClientId);
                    if (editcontact != null)
                    {
                        if (client.Selected)
                        {
                            editcontact.FprContactId = fprcontact.ContactId;
                            editcontact.CustomerId = client.CustomerId;
                            editcontact.ContactClientId = client.ContactClientId;
                            _globalViewEntities.SaveChanges();
                        }
                        else
                        {
                            _globalViewEntities.ContactClients.Remove(editcontact);
                            _globalViewEntities.SaveChanges();
                        }


                    }
                    else if (client.Selected)
                    {

                        var clientContact = new ContactClient();
                        clientContact.FprContactId = fprcontact.ContactId;
                        clientContact.CustomerId = client.CustomerId;
                        _globalViewEntities.ContactClients.Add(clientContact);
                        _globalViewEntities.SaveChanges();
                    }

                }
                viewModel.FprContact = fprcontact;
                viewModel.ContactClients = contact.ContactClients;
                return Ok(fprcontact);
            }
            return NotFound();
        }

        [HttpPost, Route("api/deletecontact/{id}")]
        public IHttpActionResult DeleteContact(long id)
        {
            var fprcontact = _globalViewEntities.FprContacts.Find(id);
            if (fprcontact == null) return BadRequest();
            _globalViewEntities.FprContacts.Remove(fprcontact);
            _globalViewEntities.SaveChanges();
            var contacts = _globalViewEntities.ContactClients.Where(item => item.FprContactId == id).ToList();

            foreach (var contact in contacts)
            {
                _globalViewEntities.ContactClients.Remove(contact);
                _globalViewEntities.SaveChanges();
            }

            return Ok();

        }
    }
}
