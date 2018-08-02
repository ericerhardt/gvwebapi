using System.Collections.Generic;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using GVWebapi.RemoteData;
using GVWebapi.Models;
namespace GVWebapi.Controllers
{
    public class SurveysController : ApiController
    {
        private readonly CustomerPortalEntities _customerPortalEntities = new CustomerPortalEntities();

        public IQueryable<SurveyWithAvg> GetSurveys()
        {
            return _customerPortalEntities.SurveyWithAvgs;
        }

        [HttpGet,Route("api/surveysbyclient/{idclient}")]
        public IHttpActionResult GetSurveysByClient(int idclient)
        {
            var surveys = _customerPortalEntities.SurveyWithAvgs.Where(s => s.CustomerID == idclient).OrderByDescending(s => s.SurveyDate).AsEnumerable();
            var results = new List<SurveyViewModel>();
            foreach (var survey in surveys)
            {
                var result = new SurveyViewModel
                {
                    Survey = survey,
                    SurveyDetail = _customerPortalEntities.SurveyQuestionsWithAnswers.Where(s => s.SurveyID == survey.SurveyID).AsEnumerable(),

                };

                results.Add(result);
            }
            var totalAvg = _customerPortalEntities
                .SurveyWithAvgs
                .Where(c => c.CustomerID == idclient)
                .Select(c => new {c.SurveyID, c.Average }).Average(c => c.Average);

            var output = new
            {
                Results = results,
                Totals = totalAvg
            };
            return Json(output);
        }

        [Route("api/addsurvey/{Survey?}")]
        public IHttpActionResult Addsurvey(SurveyPostModel survey)
        {
            var newSurvey = new Survey()
            {
                CustomerID = survey.CustomerID,
                Name = survey.Name,
                Email = survey.Email,
                Title = survey.Title,
                SurveyTypeID = 2,
                SurveyDate = survey.SurveyDate
            };
            _customerPortalEntities.Surveys.Add(newSurvey);
            _customerPortalEntities.SaveChanges();

            foreach (var answer in survey.Answers)
            {
                var surveyAnswer = new SurveyAnswer()
                {
                    SurveyID = newSurvey.SurveyID,
                    QuestionID = answer.QuestionID,
                    Question = answer.Question,
                    AnswerNumeric = answer.AnswerNumeric,
                    Comments = answer.Comments,
                    NA = false
                };
                _customerPortalEntities.SurveyAnswers.Add(surveyAnswer);
                _customerPortalEntities.SaveChanges();
            }

            var surveys = _customerPortalEntities.SurveyWithAvgs.Where(s => s.CustomerID == survey.CustomerID).OrderByDescending(s => s.SurveyDate).AsEnumerable();
            var results = new List<SurveyViewModel>();

            foreach (var actualSurvey in surveys)
            {
                var surveyViewModel = new SurveyViewModel
                {
                    Survey = actualSurvey,
                    SurveyDetail = _customerPortalEntities.SurveyQuestionsWithAnswers.Where(s => s.SurveyID == actualSurvey.SurveyID).AsEnumerable()
                };
                results.Add(surveyViewModel);
            }

            return Json(results);
        }

        [ResponseType(typeof(Survey))]
        public async Task<IHttpActionResult> GetSurvey(int id)
        {
            var survey = await _customerPortalEntities.Surveys.FindAsync(id);
            if (survey == null)
            {
                return NotFound();
            }

            return Ok(survey);
        }

        [Route("api/surveydetail/{id}")]
        public IQueryable<SurveyQuestionsWithAnswer> GetSurveyDetail(int id)
        {
            return _customerPortalEntities.SurveyQuestionsWithAnswers.Where(q => q.SurveyID == id);

        }

        [ResponseType(typeof(void))]
        public async Task<IHttpActionResult> PutSurvey(int id, Survey survey)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != survey.SurveyID)
            {
                return BadRequest();
            }

            _customerPortalEntities.Entry(survey).State = System.Data.Entity.EntityState.Modified;

            try
            {
                await _customerPortalEntities.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!SurveyExists(id))
                {
                    return NotFound();
                }

                throw;
            }

            return StatusCode(HttpStatusCode.NoContent);
        }

        [ResponseType(typeof(Survey))]
        public async Task<IHttpActionResult> DeleteSurvey(int id)
        {
            var survey = await _customerPortalEntities.Surveys.FindAsync(id);
            if (survey == null)
            {
                return NotFound();
            }

            _customerPortalEntities.Surveys.Remove(survey);
            await _customerPortalEntities.SaveChangesAsync();

            return Ok(survey);
        }

        private bool SurveyExists(int id)
        {
            return _customerPortalEntities.Surveys.Count(e => e.SurveyID == id) > 0;
        }
    }
}