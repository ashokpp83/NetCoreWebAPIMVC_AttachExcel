using Microsoft.AspNetCore.Mvc;
using WebAPIMVC_AttachExcel.Interfaces;

namespace WebAPIMVC_AttachExcel
{
    /// <summary>
    ///AttachExcelController
    /// </summary>
    [Route("api/AttachExcel")]
    [ApiController]
    public class AttachExcelController : ControllerBase
    {
        private readonly IAttachExcelService AttachExcelService;

        /// <summary>
        /// AttachExcelController constructor
        /// </summary>
        /// <param name="injectedManipulateExcelService"></param>
        public AttachExcelController(IAttachExcelService injectedAttachExcelServiceService)
        {
            AttachExcelService = injectedAttachExcelServiceService;
        }


        [HttpGet]
        public string Get()
        {
            return ".NET Core 3.1 - AttachExcelAPI is running";
        }

        //api/AttachExcel/
        /// <summary>
        /// This API will return calculated Data for GRR Report
        /// </summary>       
        /// <returns></returns>
        [HttpPost]
        //[Route("UploadToMaximo")]
        public string UploadToMaximo([FromBody] string json)
        {
            if (AttachExcelService.AttachExcelToMaximo(json))
                return "Successfully Uploaded PDF Report to Maximo";
            else
                return "Error Occurred while Download/Upload the PDF Report :" + json;
        }

        [HttpPost]
        [Route("Test")]
        public string Test()
        {
            
            return "test - post return";
            
        }

    }       
}
