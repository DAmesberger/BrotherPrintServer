using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Threading.Tasks;
using EmbedIO;
using EmbedIO.Routing;
using EmbedIO.Utilities;
using EmbedIO.WebApi;
using Unosquare.Tubular;

namespace BrotherPrintServer
{
    public class BrotherController : WebApiController
    {
        // Gets all records.
        // This will respond to
        //     GET http://localhost:9696/api/people
        [Route(HttpVerbs.Get, "/printers")]
        public Task<IEnumerable<PrinterInfo>> GetAllPrinters() => Brother.GetPrintersAsync();

        [Route(HttpVerbs.Get, "/templates")]
        public Task<IEnumerable<TemplateInfo>> GetAllTemplates() => Brother.GetTemplatesAsync();

        [Route(HttpVerbs.Get, "/update")]
        public Task Update() => Brother.UpdateAsync();

        [Route(HttpVerbs.Post, "/print")]
        public async Task Print() => await Brother.PrintAsync(await HttpContext.GetRequestDataAsync<PrintData>());

        [Route(HttpVerbs.Post, "/preview")]
        public async Task<string> Preview() => await Brother.PreviewAsync(await HttpContext.GetRequestDataAsync<PreviewData>());
    } 
}
