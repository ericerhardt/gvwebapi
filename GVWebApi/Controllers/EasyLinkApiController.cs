using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using GV.Domain;
using GV.Lookup;
using GV.Services;
using GVWebapi.Models;

namespace GVWebapi.Controllers
{
    public class EasyLinkApiController : ApiController
    {
        private readonly IEasyLinkFileSaveService _easyLinkFileSaveService;
        private readonly IEasyLinkService _easyLinkService;
        private readonly IEasyLinkFileDeleteService _easyLinkFileDeleteService;
        private readonly IUnitOfWork _unitOfWork;
        private readonly IEasyLinkChildManagerService _easyLinkChildManagerService;

        public EasyLinkApiController(IEasyLinkFileSaveService easyLinkFileSaveService, IEasyLinkService easyLinkService, IEasyLinkFileDeleteService easyLinkFileDeleteService, IUnitOfWork unitOfWork, IEasyLinkChildManagerService easyLinkChildManagerService)
        {
            _easyLinkFileSaveService = easyLinkFileSaveService;
            _easyLinkService = easyLinkService;
            _easyLinkFileDeleteService = easyLinkFileDeleteService;
            _unitOfWork = unitOfWork;
            _easyLinkChildManagerService = easyLinkChildManagerService;
        }

        [HttpPost, Route("api/easylink/uploadfile")]
        public async Task<IHttpActionResult> UploadFile()
        {
            if (Request.Content.IsMimeMultipartContent() == false)
                throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);

            var provider = new MultipartMemoryStreamProvider();
            await Request.Content.ReadAsMultipartAsync(provider);

            foreach (var file in provider.Contents)
            {
                var fileName = file.Headers.ContentDisposition.FileName;
                var filesBytes = await file.ReadAsByteArrayAsync();
                var easyLinkEntity = _easyLinkFileSaveService.SaveFile(new EasyLinkFileSaveModel(filesBytes, fileName));
                _easyLinkService.LoadFile(easyLinkEntity);
            }

            return Ok();
        }

        [HttpGet, Route("api/easylink/getall")]
        public IHttpActionResult GetAll()
        {
            return Ok(_easyLinkService.GetAll());
        }

        [HttpGet, Route("api/easylink/delete/{easyLinkId}")]
        public IHttpActionResult DeleteFile(long easyLinkId)
        {
            _easyLinkFileDeleteService.DeleteFile(easyLinkId);
            _unitOfWork.Commit();
            return Ok();
        }

        [HttpGet, Route("api/easylink/loadaddchildmanager")]
        public IHttpActionResult GetChildManagerAddData()
        {
            var customers = _easyLinkChildManagerService.GetAllCustomers();
            customers.Insert(0, new LookupInfo(0, "-- Select a Customer --"));

            var childIds = _easyLinkChildManagerService.GetAllEasyLinkChildIds();
            childIds.Insert(0, new LookupInfo(0, "-- Select an Easy Link Id --"));

            var viewModel = new
            {
                Customers = customers,
                ChildIds = childIds
            };
            return Ok(viewModel);
        }

        [HttpPost, Route("api/easylink/addchildmatch")]
        public IHttpActionResult AddChildMatch(EasyLinkChildMatchSaveModel model)
        {
            _easyLinkChildManagerService.AddEasyLinkChildMatch(model.CustomerId, model.ChildId, model.IsEasyLinkOnly);
            _unitOfWork.Commit();
            return Ok();
        }

        [HttpGet, Route("api/easylink/getallchild")]
        public IHttpActionResult GetChildActions(bool hideEasyLinkOnly)
        {
            return Ok(_easyLinkChildManagerService.GetChildLinks(hideEasyLinkOnly));
        }

        [HttpGet, Route("api/easylink/removelink")]
        public IHttpActionResult DeleteChildAction(long easyLinkChildId)
        {
            _easyLinkChildManagerService.RemoveLink(easyLinkChildId);
            _unitOfWork.Commit();
            return Ok();
        }

        [HttpPost,Route("api/easylink/loadaddchildmanagerforedit")]
        public IHttpActionResult GetChildManagerAddEditData(ChildManagerExistingData model)
        {
            var customers = _easyLinkChildManagerService.GetAllCustomers(model.CustomerId);
            customers.Insert(0, new LookupInfo(0, "-- Select a Customer --"));

            var childIds = _easyLinkChildManagerService.GetAllEasyLinkChildIds(model.ChildId);
            childIds.Insert(0, new LookupInfo(0, "-- Select an Easy Link Id --"));

            var viewModel = new
            {
                Customers = customers,
                ChildIds = childIds
            };
            return Ok(viewModel);
        }

        [HttpPost,Route("api/easylink/switchchildid")]
        public IHttpActionResult SwitchChild(EasyLinkChildSwitchModel model)
        {
            _easyLinkChildManagerService.SwitchChildLink(model);
            _unitOfWork.Commit();
            return Ok();
        }

        [HttpGet,Route("api/easylink/getunmappedchild")]
        public IHttpActionResult GetUnMappedChild()
        {
            return Ok(_easyLinkChildManagerService.GetUnMappedChildIds());
        }
    }
}