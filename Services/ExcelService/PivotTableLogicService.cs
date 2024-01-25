using ExcelWorkbookPivotTable.Constants;
using Serilog;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.PivotTables;
using Spire.Xls.Core;
using System.Data;
using System.Drawing;

namespace ExcelWorkbookPivotTable.Services.ExcelService
{   
    public interface IPivotTableLogicService
    {
        bool isPivotTableCreated { get; }
        bool isBranchPivotCreated { get; }

        Workbook CreateWorkbook(DataTable dataTable);
    }
    public class PivotTableLogicService:IPivotTableLogicService
    {
        #region Constants

        List<string> monthSheetsToDelete = new List<string>();
        public bool isPivotTableCreated { get; private set; }

        public bool isBranchPivotCreated { get; private set; }
        #endregion


        // to create a new workbook, where all the sheets are excel created
        #region Main method to create multiple sheets, pivot tables etc.
        public Workbook CreateWorkbook(DataTable dataTable)
        {
            Workbook workbook = new Workbook();
            try
            {

                workbook.Worksheets[0].Remove();
                //workbook.Worksheets['sheet1'].Remove();
                // to filter the data based on months

                Dictionary<string, DataTable> filteredDataByMonth = FilterDataByMonth(dataTable);

                Worksheet pivotTableSheet = workbook.Worksheets[1];
                pivotTableSheet.Name = $"{DateTime.Now:yyyy} Teammate";

                Worksheet dashboardSheet = workbook.Worksheets.Add("Dashboard Data");
                CreateDashboardTable(dataTable, dashboardSheet, workbook);




                int currentColumnIndex = 1; // Initialize column index to 1
                int currentMonthColumnIndex = 1; // Initialize month column index to 1
                int monthSheetCount = 0;
                foreach (var x in filteredDataByMonth)
                {
                    int lastRow = x.Value.Rows.Count + 1;
                    int lastCol = x.Value.Columns.Count;

                    string month = x.Key;
                    DataTable monthTable = x.Value;

                    Worksheet monthSheet = workbook.Worksheets.Add(month); // Create a new sheet with the month's name
                    monthSheet.InsertDataTable(monthTable, true, 1, 1);
                    monthSheet.Visibility = WorksheetVisibility.StrongHidden; // to hide the month sheet
                    monthSheetCount++;
                    // Calculate the target cell address for the pivot table
                    string position = GetExcelCellAddress(currentColumnIndex, 2);
                    string monthPosition = GetExcelCellAddress(currentMonthColumnIndex, 1);

                    // Create the pivot table in the Pivot Table sheet
                    workbook = CreatePivotTable(workbook, monthSheet, pivotTableSheet, lastRow, lastCol, position, month, monthPosition);

                    // Update the current column index for the next iteration
                    currentColumnIndex += 3; // Move to the second next column 
                    currentMonthColumnIndex += 3;
                    // Add the month sheet to the list for deletion
                    //monthSheetsToDelete.Add(month);
                }
                // Create the branchsheet
                Worksheet branchSheet = workbook.Worksheets.Add($"{DateTime.Now:yyyy} Branch");
                // Create the pivot table on the branchSheet
                workbook = CreateBranchTable(dataTable, branchSheet, dashboardSheet, workbook, monthSheetCount);

                //workbook = DeleteMonthSheets(monthSheetsToDelete, workbook);
                workbook.Worksheets[0].Remove();
                branchSheet.MoveWorksheet(0);
            }
            catch (Exception ex)
            {

                Log.Error(ex, "Error occurred while creating the workbook.");
            }
            return workbook;
        }
        #endregion

        // to delete the month based sheets after the pivot table creation
        public Workbook DeleteMonthSheets(List<string> monthSheetsToDelete, Workbook workbook)
        {
            foreach (var months in monthSheetsToDelete)
            {
                workbook.Worksheets.Remove(months);
            }
            return workbook;
        }

        //to create a Dashboard Data sheet
        public void CreateDashboardTable(DataTable dataTable, Worksheet dashboardSheet, Workbook workbook)
        {
            try
            {
                dashboardSheet.InsertDataTable(dataTable, true, 1, 1);
                dashboardSheet.AllocatedRange.AutoFitColumns();
                dashboardSheet.AllocatedRange.AutoFitRows();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error occurred while creating the Dashboard Sheet Table.");
            }
        }


        //to filter the Data table by months
        public Dictionary<string, DataTable> FilterDataByMonth(DataTable dataTable)
        {
            Dictionary<string, DataTable> filteredDataByMonth = new Dictionary<string, DataTable>();

            foreach (DataRow row in dataTable.Rows)
            {
                string month = row["Date"].ToString();
                if (!filteredDataByMonth.ContainsKey(month))
                {
                    DataTable filteredDataTable = dataTable.Clone();
                    filteredDataTable.TableName = month;
                    filteredDataByMonth.Add(month, filteredDataTable);
                }

                filteredDataByMonth[month].Rows.Add(row.ItemArray);
            }

            return filteredDataByMonth;
        }

        //to get the Column name that need to position the Pivot Table
        public string GetExcelCellAddress(int columnIndex, int rowIndex)
        {
            try
            {
                int dividend = columnIndex;
                string columnName = string.Empty;

                while (dividend > 0)
                {
                    int modulo = (dividend - 1) % 26;
                    columnName = Convert.ToChar(65 + modulo) + columnName;
                    dividend = (dividend - modulo) / 26;
                }

                return columnName + rowIndex.ToString();
            }
            catch (Exception ex)
            {
                // Log the exception using Serilog
                Log.Error(ex, "Error occurred while getting the Excel cell address.");

                // Return an error message
                return "Error: ";
            }
        }


        #region To Create Pivot Tables based on months
        //to create pivot tables based on months
        public Workbook CreatePivotTable(Workbook workbook, Worksheet monthSheet, Worksheet pivotTableSheet, int lastRow, int lastCol, string position, string month, string monthPosition)
        {
            try
            {
                pivotTableSheet.Range[monthPosition].Text = month.GetMonth();
                pivotTableSheet.Range[monthPosition].Style.Font.IsBold = true;

                CellRange dataRange = monthSheet.Range[1, 1, lastRow, lastCol]; // Define the range using two coordinates
                PivotCache cache = workbook.PivotCaches.Add(dataRange);
                IPivotTable pt = pivotTableSheet.PivotTables.Add("Pivot Table", pivotTableSheet.Range[position], cache);

                // Set Row Labels
                var r1 = pt.PivotFields["Source"];
                r1.Axis = AxisTypes.Row;
                pt.Options.RowHeaderCaption = "Location";

                // Add Pivot Fields
                var referralsField = pt.PivotFields["Originator"];
                referralsField.Axis = AxisTypes.Row;

                // Set Data Fields (referrals)
                pt.DataFields.Add(pt.PivotFields["Originator"], "Referrals", SubtotalTypes.Count);

                // Sort the "Source" field in asceding order
                PivotField sourceField = pt.PivotFields["Source"] as PivotField;
                sourceField.SortType = PivotFieldSortType.Ascending;

                // Set Style
                pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium6;

                for (int i = 0; i < pivotTableSheet.PivotTables.Count; i++)
                {
                    XlsPivotTable xlsPivotTable = (XlsPivotTable)pivotTableSheet.PivotTables[i];
                    xlsPivotTable.Cache.IsRefreshOnLoad = true;
                    xlsPivotTable.CalculateData();

                }
                pivotTableSheet.AllocatedRange.AutoFitColumns();
                pivotTableSheet.AllocatedRange.AutoFitRows();
                isPivotTableCreated = workbook.Worksheets.Contains(pivotTableSheet);

            }
            catch (Exception ex)
            {

                Log.Error(ex, "Error occurred while creating the Team Mate Sheet Pivot table.");
            }
            return workbook;
        }
        #endregion

        // to create the branch sheet table
        #region To Create Branch Sheet Table
        public Workbook CreateBranchTable(DataTable dataTable, Worksheet branchSheet, Worksheet dashboardSheet, Workbook workbook, int monthSheetCount)
        {
            try
            {

                CellRange dataRange = dashboardSheet.Range[1, 1, dataTable.Rows.Count + 1, dataTable.Columns.Count];
                PivotCache cache = workbook.PivotCaches.Add(dataRange);
                PivotTable pt = branchSheet.PivotTables.Add("Pivot Table", branchSheet.Range[1, 1], cache);

                // Set Row Labels
                var r1 = pt.PivotFields["Source"];
                r1.Axis = AxisTypes.Row;
                pt.Options.RowHeaderCaption = "Center";

                // Set Column Labels
                var c1 = pt.PivotFields["Date"];
                c1.Axis = AxisTypes.Column;

                // Set the value field
                pt.DataFields.Add(pt.PivotFields["Comments"], "", SubtotalTypes.Count);
                // Sort the "Source" field in asceding order
                PivotField sourceField = pt.PivotFields["Source"] as PivotField;
                sourceField.SortType = PivotFieldSortType.Ascending;

                // Refresh the branch sheet
                for (int i = 0; i < branchSheet.PivotTables.Count; i++)
                {
                    XlsPivotTable xlsPivotTable = (XlsPivotTable)branchSheet.PivotTables[i];
                    xlsPivotTable.Cache.IsRefreshOnLoad = false;
                    xlsPivotTable.CalculateData();
                }

                // Apply column width
                branchSheet.AllocatedRange.AutoFitColumns();
                branchSheet.AllocatedRange.AutoFitRows();

                // Set the header cell to blank with a column width of 5
                branchSheet.Range["B1"].Text = " ";
                branchSheet.Range["B1"].ColumnWidth = 8;

                int numColumns;
                int numRows;
                //to solve the styling issue of borders for the branch sheet pivot table
                if (monthSheetCount >= 5)
                {
                    numColumns = branchSheet.Columns.Length;
                    numRows = branchSheet.Rows.Length;
                }
                if (monthSheetCount <= 4)
                {
                    numColumns = branchSheet.Columns.Length - 1;
                    numRows = branchSheet.Rows.Length;
                    if (monthSheetCount == 3)
                    {
                        numColumns = branchSheet.Columns.Length - 2;
                        numRows = branchSheet.Rows.Length;
                    }
                    if (monthSheetCount == 2)
                    {
                        numColumns = branchSheet.Columns.Length - 3;
                        numRows = branchSheet.Rows.Length;
                    }
                    if (monthSheetCount == 1)
                    {
                        if (branchSheet.Columns.Length <= 14)
                        {
                            numColumns = branchSheet.Columns.Length - 4;
                            numRows = branchSheet.Rows.Length - 4;
                        }

                        //int sample = branchSheet.LastRow;
                        numColumns = branchSheet.Columns.Length - 4;
                        numRows = branchSheet.Rows.Length;
                    }
                }
                else
                {
                    numColumns = branchSheet.Columns.Length;
                    numRows = branchSheet.Rows.Length;
                }


                for (int colIndex = 1; colIndex <= numColumns; colIndex++)
                {
                    CellRange colRange = branchSheet.Range[1, colIndex, numRows, colIndex];
                    colRange.Style.HorizontalAlignment = HorizontalAlignType.Center;
                    colRange.Style.Font.Size = 8;
                }

                // Apply borders between columns
                for (int colIndex = 1; colIndex <= numColumns; colIndex++)
                {
                    CellRange colRange = branchSheet.Range[1, colIndex, numRows, colIndex];
                    colRange.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
                    colRange.Borders[BordersLineType.EdgeRight].Color = Color.LightGray;
                }

                branchSheet.AllocatedRange.Style.Borders.Color = Color.Gray;
                pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleLight1;

                // Calculate the cell addresses for "YTD" and "MTD"
                string cellAddressYTD = GetExcelCellAddress(numColumns, 2); // Second row and last column
                string cellAddressMTD = GetExcelCellAddress(1, numRows); // Last row and first column

                // Set the text for "YTD" and "MTD" in the calculated cell addresses
                branchSheet.Range[cellAddressYTD].Text = "YTD";
                branchSheet.Range[cellAddressMTD].Text = "Mo Total";
                isBranchPivotCreated = workbook.Worksheets.Contains(branchSheet);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error occurred while creating the Branch Sheet Pivot Table....");
            }

            return workbook;
        }

        #endregion
    }
}
