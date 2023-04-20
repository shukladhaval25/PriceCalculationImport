using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;

namespace PriceCalculationImport
{
    internal class ExcelProcess
    {
        DataTable dtWorkbook = new DataTable();
        DataTable dtPriceMaster = new DataTable();
        DataTable dtPrice = new DataTable();
        DataTable dtFinal = new DataTable();
        DataTable dtError = new DataTable();
        string selectedFileName = string.Empty;
        string selectedPath = string.Empty;

        public ExcelProcess(string selectedFileName, string selectedPath)
        {
            this.selectedFileName = selectedFileName;
            this.selectedPath = selectedPath;

            dtError.Columns.Add("WorksheetName");
            dtError.Columns.Add("Range");
            dtError.Columns.Add("Clarity");
            dtError.Columns.Add("Colour");
            dtError.Columns.Add("Description");

            dtPrice.Columns.Add("WorksheetName");
            dtPrice.Columns.Add("Range");
            dtPrice.Columns.Add("Clarity");
            dtPrice.Columns.Add("Cut");
            dtPrice.Columns.Add("Colour");
            dtPrice.Columns.Add("Florescence");
            dtPrice.Columns.Add("OriginalPrice").DataType = typeof(double);
            dtPrice.Columns.Add("NewPrice").DataType = typeof(double);
            dtPrice.Columns.Add("Percentage").DataType = typeof(double);
            dtPrice.Columns.Add("Date").DataType = typeof(DateTime);

            dtPriceMaster.Columns.Add("Range");
            dtPriceMaster.Columns.Add("Clarity");
            dtPriceMaster.Columns.Add("Colour");
            dtPriceMaster.Columns.Add("Price").DataType = typeof(double);
        }
        public List<string> GetAllWorksheets(string fileName)
        {
            dtWorkbook.Clear();
            dtWorkbook.Columns.Add("WorksheetName");
            dtWorkbook.Columns.Add("Status");

            List<string> worksheets = new List<string>();
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            //Initialize a new Workboook object
            Microsoft.Office.Interop.Excel.Workbook workbook;
            //Load the document
            workbook = application.Workbooks.Open(fileName);
            //Get all worksheets
            foreach (Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Visible == XlSheetVisibility.xlSheetVisible)
                {
                    DataRow dr = dtWorkbook.NewRow();
                    dr["WorksheetName"] = sheet.Name;
                    dr["Status"] = "";
                    dtWorkbook.Rows.Add(dr);
                    worksheets.Add(sheet.Name);
                }
            }
            
            workbook.Close();
            application.Quit();
            workbook = null;
            GC.Collect();
            return worksheets;
        }
        public DataTable GetAllWorksheetAsDataTbale(string fileName)
        {
            GetAllWorksheets(fileName);
            return dtWorkbook;
        }

        public string GetFileName
        {
            get { return selectedFileName; }
        }

        public DataTable GetFinalPriceTable()
        {
            return dtFinal;
        }

        public async void ProcessEachWorkSheet(string workSheetName)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
                //Initialize a new Workboook object
                Microsoft.Office.Interop.Excel.Workbook workbook;
                //Load the document
                workbook = application.Workbooks.Open(selectedPath);
                //TODO: Add logic to process each worksheet with business logic

                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets[workSheetName];

                if (sheet.Name.Equals(workSheetName))
                {
                    sheet.Activate();
                    Excel.Range xlRange = sheet.UsedRange;
                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;
                    int lineBreackRowIndex = 0;
                    bool isLineBreack = false;

                    //dtPriceMaster.Columns.Clear();
                    dtPriceMaster.Rows.Clear();
                    //dtPriceMaster.Clear();
                   

                    processPriceMaster(xlRange, rowCount, colCount, ref lineBreackRowIndex, ref isLineBreack);

                    isLineBreack = false;
                    // From here onward calculate new price based on percentage define.
                    designFinalProcessTableStructure();

                    isLineBreack = processForCalculateNewPriceBasedOnPercentage(workSheetName, xlRange, rowCount, colCount, lineBreackRowIndex, isLineBreack);
                }

                workbook.Close();
                application.Quit();
                workbook = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace);

            }
        }

        private bool processForCalculateNewPriceBasedOnPercentage(string workSheetName, Range xlRange, int rowCount, int colCount, int lineBreackRowIndex, bool isLineBreack)
        {
            for (int i = lineBreackRowIndex + 4; i <= rowCount; i++)
            {
                string rangeValue = string.Empty;
                string colourValue = string.Empty;
                string clarityValue = string.Empty;
                string lastRangeValue = string.Empty;
                string lastColourValue = string.Empty;
                string lastClarityValue = string.Empty;
                double originalPrice = 0;

                for (int j = 1; j <= colCount; j++)
                {
                    //new line
                    if (j == 1)
                        rangeValue = xlRange.Cells[1, j].Value2.ToString();

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {

                        if (j == 1)
                        {
                            if (xlRange.Cells[i, j].Value2.ToString() == "")
                            {
                                for (int rowIndex = i; rowIndex >= lineBreackRowIndex + 4; rowIndex--)
                                {
                                    if (xlRange.Cells[rowIndex, j].Value2.ToString() != "")
                                    {
                                        colourValue = xlRange.Cells[rowIndex, j].Value2.ToString();
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                colourValue = xlRange.Cells[i, j].Value2.ToString();
                            }

                            if (colourValue.Equals("Range =>"))
                            {
                                isLineBreack = true;
                                dtFinal.Merge(dtPrice);
                                break;
                            }
                        }
                        else
                        {
                            if (isLineBreack == false)
                            {
                                double percentageValue = 0;
                                double newPriceValue = 0;
                                double originalPercentage = 0;
                                double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out percentageValue);
                                originalPercentage = percentageValue;
                                if (percentageValue != 0)
                                {
                                    if (xlRange.Cells[lineBreackRowIndex + 1, j].Value2 == null)
                                    {
                                        for (int columnIndex = j; columnIndex >= 1; columnIndex--)
                                        {
                                            if (xlRange.Cells[lineBreackRowIndex + 1, columnIndex].Value2 != null)
                                            {
                                                clarityValue = xlRange.Cells[lineBreackRowIndex + 1, columnIndex].Value2.ToString();
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        clarityValue = xlRange.Cells[lineBreackRowIndex + 1, j].Value2.ToString();
                                    }
                                    string cutValue = xlRange.Cells[lineBreackRowIndex + 2, j].Value2.ToString();
                                    string florescenceValue = xlRange.Cells[i, 2].Value2.ToString();

                                    if (!rangeValue.Equals(lastRangeValue) && !clarityValue.Equals(lastClarityValue) && !colourValue.Equals(lastColourValue))
                                    {
                                        DataRow[] result = dtPriceMaster.Select("Range ='" + rangeValue + "' and Clarity ='" + clarityValue + "' and Colour = '" + colourValue + "'");
                                        if (result.Count() == 0)
                                        {
                                            originalPrice = 0;
                                            DataRow drError = dtError.NewRow();
                                            drError["WorksheetName"] = workSheetName;
                                            drError["Range"] = rangeValue;
                                            drError["Clarity"] = clarityValue;
                                            drError["Colour"] = colourValue;
                                            drError["Description"] = "Unable to fetch price value from price master";
                                            dtError.Rows.Add(drError);
                                        }
                                        else
                                        {
                                            foreach (DataRow row in result)
                                            {
                                                originalPrice = (double)row["Price"];
                                            }
                                        }
                                        lastRangeValue = rangeValue;
                                        lastClarityValue = clarityValue;
                                        lastColourValue = colourValue;
                                    }
                                    // If negative then add those percentage into original value else deduct that percentage amount from value.
                                    if (percentageValue < 0)
                                    {
                                        percentageValue = percentageValue * -1;
                                        double val = (percentageValue / 100);
                                        newPriceValue = originalPrice * (1 + val);
                                    }
                                    else
                                    {
                                        newPriceValue = originalPrice - ((originalPrice * percentageValue) / 100);
                                    }
                                    DataRow dataRow = dtPrice.NewRow();
                                    dataRow["WorksheetName"] = workSheetName;
                                    dataRow["Range"] = rangeValue;
                                    dataRow["Clarity"] = clarityValue;
                                    dataRow["Cut"] = cutValue;
                                    dataRow["Colour"] = colourValue;
                                    dataRow["Florescence"] = florescenceValue;
                                    dataRow["Percentage"] = originalPercentage;
                                    dataRow["OriginalPrice"] = System.Math.Round(originalPrice);
                                    dataRow["NewPrice"] = System.Math.Round(newPriceValue);
                                    dataRow["Date"] = DateTime.Now.Date;
                                    dtPrice.Rows.Add(dataRow);
                                }
                            }
                        }
                    }
                    else
                    {
                        if (j == 1)
                        {
                            for (int rowIndex = i; rowIndex >= lineBreackRowIndex + 4; rowIndex--)
                            {
                                if (xlRange.Cells[rowIndex, j].Value2 != null)
                                {
                                    colourValue = xlRange.Cells[rowIndex, j].Value2.ToString();
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            return isLineBreack;
        }

        private void designFinalProcessTableStructure()
        {
            //dtPrice.Clear();
            //dtPrice.Columns.Clear();
            dtPrice.Rows.Clear();
           
        }

        private void processPriceMaster(Range xlRange, int rowCount, int colCount, ref int lineBreackRowIndex, ref bool isLineBreack)
        {
            for (int i = 2; i <= rowCount; i++)
            {
                string rangeValeue = string.Empty;
                string colourValue = string.Empty;
                string clarityValue = string.Empty;
                double priceValue = 0;
                for (int j = 1; j <= colCount; j++)
                {
                    if (j == 1)
                        rangeValeue = xlRange.Cells[1, j].Value2.ToString();

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        if (j == 1)
                        {
                            colourValue = xlRange.Cells[i, j].Value2.ToString();
                            if (colourValue.Equals("Range =>"))
                            {
                                isLineBreack = true;
                                lineBreackRowIndex = i;
                                break;
                            }
                        }
                        else
                        {
                            if (isLineBreack == false)
                            {
                                clarityValue = xlRange.Cells[1, j].Value2.ToString();
                                double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out priceValue);

                                DataRow dataRow = dtPriceMaster.NewRow();
                                dataRow["Range"] = rangeValeue;
                                dataRow["Clarity"] = clarityValue;
                                dataRow["Colour"] = colourValue;
                                dataRow["Price"] = priceValue;
                                dtPriceMaster.Rows.Add(dataRow);
                            }
                        }
                    }
                }
                if (isLineBreack)
                {
                    break;
                }
            }
        }
    }
}
