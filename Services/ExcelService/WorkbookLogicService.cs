using ExcelWorkbookPivotTable.Models.ResponseModel;
using ExcelWorkbookPivotTable.Services.DBService;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Serilog.Core;
using Spire.Xls;
using Spire.Xls.Core;
using Spire.Xls.Core.Spreadsheet.PivotTables;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ExcelWorkbookPivotTable.Services.ExcelService
{
    public interface IWorkbookLogicService
    {
        public UserResponse CreateNewWorksheet(Models.RequestModel.UserRequest request);
    }

    public class WorkbookLogicService : IWorkbookLogicService
    {

        private readonly IDatabaseService _databaseService;
        public readonly IPivotTableLogicService _pivotTableLogicService;
        private readonly IConfiguration _configuration;
        private readonly ILogger<WorkbookLogicService> _logger;


        public WorkbookLogicService(IDatabaseService databaseService, ILogger<WorkbookLogicService> logger, IPivotTableLogicService pivotTableLogicService, IConfiguration configuration)
        {
            _databaseService = databaseService;
            _logger = logger;
            _pivotTableLogicService = pivotTableLogicService;
            _configuration = configuration;
        }


        public UserResponse CreateNewWorksheet(Models.RequestModel.UserRequest request)
        {
            try
            {
                DataTable dataTable = _databaseService.GetDataForPivot(request);
                //DataTable dataTable = _databaseService.GenerateRandomDataTable(request.startDate, request.endDate, request.rowCount);

                Workbook workbook = _pivotTableLogicService.CreateWorkbook(dataTable);


                string folderPath = _configuration["DownloadStrings:DownloadURL"];
                Directory.CreateDirectory(folderPath); // Create the folder if it doesn't exist

                string fileName = $"Pivot Table.{DateTime.Now:yyyy-MM_dd_HH_mm_ss}.Workbook.xlsx";
                string filePath = Path.Combine(folderPath, fileName);

                workbook.SaveToFile(filePath, ExcelVersion.Version2010);

                File.SetAttributes(filePath, File.GetAttributes(filePath) | FileAttributes.Normal);

                //// Use the ProcessStartInfo class to specify that the file should be opened with the default program
                //System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo(filePath);
                //psi.UseShellExecute = true;
                //System.Diagnostics.Process.Start(psi);

                byte[] filebyte = System.IO.File.ReadAllBytes(filePath);

                if (System.IO.File.Exists(filePath) && _pivotTableLogicService.isPivotTableCreated && _pivotTableLogicService.isBranchPivotCreated)
                {
                    //System.IO.File.Delete(filePath);
                    //return File(System.IO.File.OpenRead(filePath), "application/octet-stream", Path.GetFileName(filePath));
                    return new UserResponse
                    {
                        IsCreated = true,
                        Message = "Created Complete Workbook Successfully..",
                        FileName = fileName,
                        FilePath = filePath,
                        FileByte = filebyte,
                    };
                }
                else
                {
                    return new UserResponse
                    {
                        IsCreated = false,
                        Message = "Complete Workbook Creation is not Successfull..",
                        FileName = string.Empty,
                        FilePath = string.Empty,
                        FileByte = filebyte,
                    };
                }


            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Complete Workbook Creationhas failed....");
                return new UserResponse
                {
                    FileName = string.Empty,
                    FilePath = string.Empty,
                    FileByte = new byte[4],
                    IsCreated = false,
                    Message = "Pivot Table Creation is not Successfull.."
                };
            }

        }
    }
}
