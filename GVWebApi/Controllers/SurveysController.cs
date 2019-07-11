using System.Collections.Generic;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using GVWebapi.RemoteData;
using GVWebapi.Models;
using System.Web;
using System.IO;
using System.Net.Http;
using System.Diagnostics;
using GVWebapi.Helpers;
using System.Collections.Specialized;
using Newtonsoft.Json;
using System;

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
       
        [HttpPost, Route("api/addsurvey/")]
        public async Task<IHttpActionResult> PostFormData()
        {
            var httpRequest = HttpContext.Current.Request;

            if (!Request.Content.IsMimeMultipartContent())
            {
                throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);
            }
             
            var provider = await Request.Content.ReadAsMultipartAsync(new InMemoryMultipartFormDataStreamProvider());
            NameValueCollection formData = provider.FormData;

            var Survey = new Survey();
            if (formData["isEdit"] == "true")
            {
                var id = int.Parse(formData["SurveyID"]);
                Survey = _customerPortalEntities.Surveys.Find(id);
                Survey.CustomerID = int.Parse(formData["CustomerID"]);
                Survey.Name = formData["Name"];
                Survey.Email = formData["Email"];
                Survey.Title = formData["Title"];
                Survey.SurveyTypeID = 2;
                Survey.SurveyDate = DateTime.Parse(formData["SurveyDate"]);
                if (formData["Answers"].Count() > 0 && Survey != null)
                {

                    IList<SurveyAnswer> surveyAnswers = JsonConvert.DeserializeObject<IList<SurveyAnswer>>(formData["Answers"]);
                    foreach (var answer in surveyAnswers)
                    {
                        var surveyAnswer = _customerPortalEntities.SurveyAnswers.Where(sa => sa.SurveyID == id && sa.QuestionID == answer.QuestionID).FirstOrDefault();
                        if (surveyAnswer != null)
                        {
                            surveyAnswer.SurveyID = Survey.SurveyID;
                            surveyAnswer.QuestionID = answer.QuestionID;
                            surveyAnswer.Question = answer.Question;
                            surveyAnswer.AnswerNumeric = answer.AnswerNumeric;
                            surveyAnswer.Comments = answer.Comments;
                            surveyAnswer.NA = false;
                        }
                       
                    }
                    _customerPortalEntities.SaveChanges();
                }
            }
            else
            {
                var newSurvey = new Survey()
                {
                    CustomerID = int.Parse(formData["CustomerID"]),
                    Name = formData["Name"],
                    Email = formData["Email"],
                    Title = formData["Title"],
                    SurveyTypeID = 2,
                    SurveyDate = DateTime.Parse(formData["SurveyDate"]),
                };
                _customerPortalEntities.Surveys.Add(newSurvey);
                _customerPortalEntities.SaveChanges();
                Survey = newSurvey;
                if (formData["Answers"].Count() > 0)
                {

                    IList<SurveyAnswer> surveyAnswers = JsonConvert.DeserializeObject<IList<SurveyAnswer>>(formData["Answers"]);
                    foreach (var answer in surveyAnswers)
                    {
                        var surveyAnswer = new SurveyAnswer()
                        {
                            SurveyID = Survey.SurveyID,
                            QuestionID = answer.QuestionID,
                            Question = answer.Question,
                            AnswerNumeric = answer.AnswerNumeric,
                            Comments = answer.Comments,
                            NA = false
                        };
                        _customerPortalEntities.SurveyAnswers.Add(surveyAnswer);
                        _customerPortalEntities.SaveChanges();
                    }
                }
            }
           

              
          
           
            

            //access files  
            IList<HttpContent> files = provider.Files;
            if (files.Count() > 0)
            {
                HttpContent file1 = files[0];
                var fileName = file1.Headers.ContentDisposition.FileName.Trim('\"');
                Stream input = await file1.ReadAsStreamAsync();


                var path = Path.Combine(HttpContext.Current.Server.MapPath("~/uploads/surveys"), fileName);
                //Deletion exists file  
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                using (Stream file = File.Create(path))
                {
                    input.CopyTo(file);
                    //close file  
                    file.Close();
                    var survey = _customerPortalEntities.Surveys.Find(Survey.SurveyID);
                    if (survey != null)
                    {
                        survey.Attachment =  fileName;
                          _customerPortalEntities.SaveChanges();
                    }
                }

            }

            var surveys = _customerPortalEntities.SurveyWithAvgs.Where(s => s.CustomerID == Survey.CustomerID).OrderByDescending(s => s.SurveyDate).AsEnumerable();
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
        public async Task<IHttpActionResult> GetSurveyDetail(int id)
        {
            var surveyinfo = await _customerPortalEntities.Surveys.FindAsync(id);
            if (surveyinfo == null)
            {
                return NotFound();
            }

            var answers =  _customerPortalEntities.SurveyQuestionsWithAnswers.Where(q => q.SurveyID == id).ToList();

            return Json(new { survey = surveyinfo, answers });

        }

        
        [HttpPost,Route("api/deletesurvey/{id}")]       
        public async Task<IHttpActionResult> DeleteSurvey(int id)
        {
            var surveyAnswers = _customerPortalEntities.SurveyAnswers.Where(a => a.SurveyID == id).ToList();
            if(surveyAnswers!= null)
            {
                foreach(var answer in surveyAnswers)
                {
                    _customerPortalEntities.SurveyAnswers.Remove(answer);
                    await _customerPortalEntities.SaveChangesAsync();
                }
              
            }
            var survey = await _customerPortalEntities.Surveys.FindAsync(id);
            if (survey == null)
            {
                return NotFound();
            }
            var path = string.Empty;
            if(!String.IsNullOrEmpty(survey.Attachment))
              path = Path.Combine(HttpContext.Current.Server.MapPath("~/uploads/surveys"), survey.Attachment);
            //Deletion exists file  
            if (File.Exists(path))
            {
                File.Delete(path);
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