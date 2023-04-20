
using Microsoft.Office.Interop.Excel;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;

namespace PriceCalculationImport
{
    public partial class Form1 : Form
    {
        DataTable dtWorkbook = new DataTable();

        ExcelProcess process;
        public Form1()
        {
            InitializeComponent();
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel Files|*.xlsx;*.xls";
            openFileDialog1.Multiselect = false;
           
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                progressBar1.Show();
                lblSucessMsg.Text = "";
                lblStartTime.Text = DateTime.Now.ToShortTimeString();
                string selectedFileName = string.Empty;
                string selectedPath = string.Empty;
                selectedFileName = System.IO.Path.GetFileName(openFileDialog1.FileName);
                selectedPath = openFileDialog1.FileName;
                process = new ExcelProcess(selectedFileName,selectedPath);
                txtFile.Text = openFileDialog1.FileName;
                dataGridSettings();
                fillupWorkSheetInformation(selectedPath);
                processWorkSheet();
            }
            else
            {
                txtFile.Text = "";
                return;
            }

        }

        private async void processWorkSheet()
        {
            try
            {
                if (dtWorkbook.Rows.Count > 0)
                {
                    foreach (DataRow row in dtWorkbook.Rows)
                    {
                        if ((!row["WorksheetName"].ToString().Contains(" To ") &&
                            !row["WorksheetName"].ToString().Contains("-")))
                        {
                            row["Status"] = "NA";
                        }
                        else
                        {
                            row["Status"] = "Prcessing";
                            txtProcessFileName.Text = row["WorksheetName"].ToString();
                            await Task.Run(() => process.ProcessEachWorkSheet(row["WorksheetName"].ToString()));
                            row["Status"] = "Completed";
                            dataGridFinalProcess.DataSource = process.GetFinalPriceTable();
                        }
                    }

                    exportToCSV();
                    lblStopTime.Text = DateTime.Now.ToShortTimeString();
                    lblSucessMsg.Text = "File Processing completed sucessfully";
                    progressBar1.Hide();
                }
            }
            catch(Exception ex)
            {
                lblSucessMsg.ForeColor = System.Drawing.Color.Red;
                lblSucessMsg.Text = "Error while processing data." +  ex.Message;
            }
        }

        private void dataGridSettings()
        {
            dataGridFinalProcess.AlternatingRowsDefaultCellStyle.BackColor = Color.White;
            dataGridFinalProcess.RowsDefaultCellStyle.BackColor = Color.LightBlue;
            dataGridFinalProcess.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Tahoma", 9F, ((System.Drawing.FontStyle)(System.Drawing.FontStyle.Bold)));
            dataGridFinalProcess.ForeColor = Color.Black;
            dataGridFinalProcess.BorderStyle = BorderStyle.None;
        }

        private void exportToCSV()
        {
            DataTable dataTable = process.GetFinalPriceTable();
            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                // create object for the StringBuilder class
                StringBuilder sb = new StringBuilder();

                // Get name of columns from datatable and assigned to the string array
                string[] columnNames = dataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();

                // Create comma sprated column name based on the items contains string array columnNames
                sb.AppendLine(string.Join(",", columnNames));

                // Fatch rows from datatable and append values as comma saprated to the object of StringBuilder class 
                foreach (DataRow row in dataTable.Rows)
                {
                    IEnumerable<string> fields = row.ItemArray.Select(field => string.Concat("\"", field.ToString().Replace("\"", "\"\""), "\""));
                    sb.AppendLine(string.Join(",", fields));
                }

                // save the file
                File.WriteAllText(@"D:\FinalProcessWithPriceCalcuation.csv", sb.ToString());
            }
        }

        private void exportToExcel()
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Workbooks.Add();
            // single worksheet
            Microsoft.Office.Interop.Excel._Worksheet workSheet = excelApp.ActiveSheet;

            DataTable dataTable = process.GetFinalPriceTable();

            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                // create object for the StringBuilder class
                StringBuilder sb = new StringBuilder();

                // Get name of columns from datatable and assigned to the string array
                string[] columnNames = dataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();

                // Create comma sprated column name based on the items contains string array columnNames
                sb.AppendLine(string.Join(",", columnNames));

                // Fatch rows from datatable and append values as comma saprated to the object of StringBuilder class 
                foreach (DataRow row in dataTable.Rows)
                {
                    IEnumerable<string> fields = row.ItemArray.Select(field => string.Concat("\"", field.ToString().Replace("\"", "\"\""), "\""));
                    sb.AppendLine(string.Join(",", fields));
                }

                // save the file
                File.WriteAllText(@"D:\FinalProcessWithPriceCalcuation.xlsx", sb.ToString());
            }
        }


        private void fillupWorkSheetInformation(string selectedPath)
        {
            dtWorkbook = process.GetAllWorksheetAsDataTbale(selectedPath);
            dataGridWorkbookStatus.DataSource = dtWorkbook;
            dataGridWorkbookStatus.Columns[0].ReadOnly = true;
            dataGridWorkbookStatus.Columns[1].ReadOnly = true;
        }

        public DataTable SelectData(string query)
        {
            NpgsqlConnection connection = new NpgsqlConnection("Server=127.0.0.1;Port=5432;CommandTimeout=5000;User Id=postgres;" +
                                                         "Password=Admin@123;Database=MyDB;");

            connection.Open();
            using (var cmd = new NpgsqlCommand(query, connection))
            {
                cmd.Prepare();

                NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);

                DataSet _ds = new DataSet();
                DataTable _dt = new DataTable();

                da.Fill(_ds);

                try
                {
                    _dt = _ds.Tables[0];
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Erro: ---> " + ex.Message);
                }

                connection.Close();
                return _dt;
            }
        }

        private void btnReadData_Click(object sender, EventArgs e)
        {
            DataTable dataTable = SelectData("Select * from finalprocessdata");
            dataGridFinalProcess.DataSource = dataTable;


            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                // create object for the StringBuilder class
                StringBuilder sb = new StringBuilder();

                // Get name of columns from datatable and assigned to the string array
                string[] columnNames = dataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();

                // Create comma sprated column name based on the items contains string array columnNames
                sb.AppendLine(string.Join(",", columnNames));

                // Fatch rows from datatable and append values as comma saprated to the object of StringBuilder class 
                foreach (DataRow row in dataTable.Rows)
                {
                    IEnumerable<string> fields = row.ItemArray.Select(field => string.Concat("\"", field.ToString().Replace("\"", "\"\""), "\""));
                    sb.AppendLine(string.Join(",", fields));
                }

                // save the file
                File.WriteAllText(@"D:\Codingvila.csv", sb.ToString());
            }
        }
    }
}