using System.Web.Http;
using GV.Domain;
using GVWebapi.Models.Reconciliation;
using GVWebapi.Services;

namespace GVWebapi.Controllers
{
    public class ReconciliationApiController : ApiController
    {
        private readonly IReconciliationService _reconciliationService;
        private readonly IUnitOfWork _unitOfWork;

        public ReconciliationApiController(IReconciliationService reconciliationService, IUnitOfWork unitOfWork)
        {
            _reconciliationService = reconciliationService;
            _unitOfWork = unitOfWork;
        }

        [HttpGet, Route("api/cyclehistory/reconciliation/{cycleId}/{custid}")]
        public IHttpActionResult GetReconciliation(long cycleId, int custid)
        {
            var reconSummary = _reconciliationService.GetReconciliation(cycleId, custid);
       
            return Ok(reconSummary);
        }

        [HttpPost,Route("api/cyclehistory/reconciliation/updatecredit")]
        public IHttpActionResult SaveReconCredit(ReconCreditSaveModel model)
        {
            _reconciliationService.UpdateReconCredit(model);
            _unitOfWork.Commit();
            return Ok();
        }
    }
}