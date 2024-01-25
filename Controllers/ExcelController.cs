using ExcelWorkbookPivotTable.Models.RequestModel;
using ExcelWorkbookPivotTable.Services.DBService;
using ExcelWorkbookPivotTable.Services.ExcelService;
using Microsoft.AspNetCore.Mvc;
using Swashbuckle.AspNetCore.Annotations;

namespace ExcelWorkbookPivotTable.Controllers
{
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly IPivotTableLogicService _pivotTableService;
        public readonly IWorkbookLogicService _workbookLogicService;
        private readonly IDatabaseService _databaseService;

        public ExcelController(IWorkbookLogicService wprkbookLogicService, IDatabaseService databaseService)
        {
            _workbookLogicService = wprkbookLogicService;
            _databaseService = databaseService;
        }
        [SwaggerOperation(Summary = "To check the API working.")]
        [HttpGet]
        [Route("InitialAPICheck")]
        public string InitialAPICheck()
        {
            return "MRM Reporting API started successfully.";
        }

        [SwaggerOperation(Summary = "To Get the Excel Workbook and Worksheets.")]
        [HttpPost]
        [Route("GetPivotTable")]
        public IActionResult GetExcelData(UserRequest request)
        {
            var response = _workbookLogicService.CreateNewWorksheet(request);
            return Ok(response);
        }

        [SwaggerOperation(Summary = "To Encrypt the Connection strings.")]
        [ApiExplorerSettings(IgnoreApi = true)]
        [HttpPost]
        [Route("AddEncryptionValues")]
        public void EncryptionValues(string encryption)
        {
            var sample = _databaseService.EncryptData(encryption);
        }

    }
}
