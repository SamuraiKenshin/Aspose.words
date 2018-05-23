using SendGrid;
using SendGrid.Helpers.Mail;
using Swashbuckle.Swagger.Annotations;
using System.IO;
using System.Net;
using System.Web.Http;
using System.Web.Hosting;
using System;
using System.Net.Http;

namespace EmailPDF.Controllers
{
    public class SendController : ApiController
    {
        [HttpGet]
        [SwaggerResponse(HttpStatusCode.OK)]
        [SwaggerResponse(HttpStatusCode.InternalServerError, Type = typeof(HttpError))]
        public IHttpActionResult GetString()
        {
            return Ok("Test");
        }

        [HttpPost]
        [SwaggerResponse(HttpStatusCode.OK)]
        [SwaggerResponse(HttpStatusCode.InternalServerError, Type = typeof(HttpError))]
        [Route("DownloadFile")]
        public HttpResponseMessage DownloadFile()
        {
            Aspose.Words.Document doc;
#if (DEBUG)
            doc = new Aspose.Words.Document(@"C:\Users\sarav\source\repos\EmailPDF\Send.docx");
#else
            doc = new Aspose.Words.Document(HostingEnvironment.MapPath("Send.docx");
#endif
            MemoryStream docOutStream = new MemoryStream();
            doc.Save(docOutStream, Aspose.Words.SaveFormat.Pdf);
            byte[] docBytes = docOutStream.ToArray();
            //adding bytes to memory stream   
            var dataStream = new MemoryStream(docBytes);
            HttpResponseMessage httpResponseMessage;
            httpResponseMessage = Request.CreateResponse(HttpStatusCode.OK);
            httpResponseMessage.Content = new StreamContent(dataStream);
            httpResponseMessage.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment");
            httpResponseMessage.Content.Headers.ContentDisposition.FileName = "Send.pdf";
            httpResponseMessage.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/pdf");
            httpResponseMessage.Content.Headers.ContentLength = dataStream.Length;

            return httpResponseMessage;
        }

        [HttpPost]
        [SwaggerResponse(HttpStatusCode.OK)]
        [SwaggerResponse(HttpStatusCode.InternalServerError, Type = typeof(HttpError))]
        [Route("SendEmail")]
        public IHttpActionResult SendEmail()
        {
            Aspose.Words.Document doc;
#if (DEBUG)
            doc = new Aspose.Words.Document(@"C:\Users\sarav\source\repos\EmailPDF\Send.docx");
#else
            doc = new Aspose.Words.Document(HostingEnvironment.MapPath("Send.docx");
#endif
            MemoryStream docOutStream = new MemoryStream();       
            doc.Save(docOutStream, Aspose.Words.SaveFormat.Pdf);
            byte[] docBytes = docOutStream.ToArray();
            var file = Convert.ToBase64String(docBytes);
            var client = new SendGridClient("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");
            var myMessage = new SendGridMessage
            {
                From = new EmailAddress("sarav1988k@gmail.com"),
                Subject = "Pdf converted from Aspose.words and sent from SendGrid",
                HtmlContent = "<p>Hello!</p>",
                PlainTextContent = "Plain Text!"
            };
            myMessage.AddTo("ranirajan58@gmail.com");
            myMessage.AddAttachment("Send.pdf", file);
            var response = client.SendEmailAsync(myMessage).GetAwaiter().GetResult();
            return Ok("Mail Sent");
        }
    }
}
