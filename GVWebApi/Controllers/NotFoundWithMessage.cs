using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
 
using System.Web.Http;

namespace GVWebapi.Controllers
{
    public class NotFoundWithMessageResult : IHttpActionResult
    {
        private readonly string _message;

        public NotFoundWithMessageResult(string message)
        {
            _message = message;
        }

        public Task<HttpResponseMessage> ExecuteAsync(CancellationToken cancellationToken)
        {
            var response = new HttpResponseMessage(System.Net.HttpStatusCode.NotFound);
            response.Content = new StringContent(_message);
            return Task.FromResult(response);
        }
    }
}