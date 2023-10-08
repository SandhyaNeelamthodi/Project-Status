using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToBarChart
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = true;

                workbook = excelApp.Workbooks.Open(@"C:\Users\248773\Downloads\SAMPLE.xlsx");
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];

                // Assuming data has headers (ID, Name, ProjectName, Location, ProjectStatus)
                int idColumn = 1; // Column A for ID
                int nameColumn = 2;
                int statusColumn = 5; // Column E for ProjectStatus

                // Get the unique IDs from the data
                Excel.Range idRange = worksheet.Range["A2:A" + worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row];
                object[,] idValues = (object[,])idRange.Value2;

                // Define the root directory where images will be saved
                string rootOutputDirectory = @"C:\Users\248773\Downloads\individual_charts\";

                // Iterate through each unique ID
                for (int i = 1; i <= idValues.GetLength(0); i++)
                {
                    string id = idValues[i, 1]?.ToString();

                    if (!string.IsNullOrEmpty(id))
                    {
                        // Turn off any existing filters
                        worksheet.AutoFilterMode = false;

                        // Filter the data for the current ID
                        idRange.AutoFilter(1, id);

                        // Get the visible (filtered) rows
                        Excel.Range visibleRows = worksheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                        if (visibleRows.Rows.Count > 1)
                        {
                            string name = visibleRows.Cells[i + 1, nameColumn].Value?.ToString();
                            // Get the second value of the "ProjectStatus" column
                            string projectStatus = visibleRows.Cells[i + 1, statusColumn].Value?.ToString();

                            // Define the subdirectory based on the "ProjectStatus" value
                            string subdirectory = Path.Combine(rootOutputDirectory, projectStatus);

                            // Ensure the subdirectory exists
                            Directory.CreateDirectory(subdirectory);

                            // Create a new chart object
                            Excel.ChartObjects charts = (Excel.ChartObjects)worksheet.ChartObjects(Type.Missing);
                            Excel.ChartObject chartObj = charts.Add(100, 100, 300, 300); // Adjust the position and size

                            // Create a bar chart
                            Excel.Chart chart = chartObj.Chart;
                            chart.ChartType = Excel.XlChartType.xlBarClustered; // Use xlBarClustered for a bar chart

                            // Create a data range for the project status
                            Excel.Range statusRange = worksheet.Cells[i + 1, statusColumn];
                            Excel.Range categoryRange = worksheet.Cells[i + 1, idColumn];

                            // Set the data source for the chart
                            chart.SetSourceData(statusRange, Excel.XlRowCol.xlColumns);
                            chart.SeriesCollection(1).XValues = categoryRange;

                            // Save the chart as an image in the subdirectory
                            string imagePath = Path.Combine(subdirectory, $"BarChart_ID_{id}.png");
                            chartObj.Chart.Export(imagePath, "PNG");

                            // Release Excel objects for the current chart
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(chartObj);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(charts);

                            Console.WriteLine($"Bar chart created for ID: {id} in folder: {name}");
                        }
                    }
                }

                // Turn off the filter for all columns
                worksheet.AutoFilterMode = false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // Release Excel objects in reverse order of creation
                if (worksheet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                }
                if (workbook != null)
                {
                    workbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }

                // Ensure Excel process is killed
                if (excelApp != null)
                {
                    System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("excel");
                    foreach (System.Diagnostics.Process process in processes)
                    {
                        process.Kill();
                    }
                }
            }
        }
    }
}
