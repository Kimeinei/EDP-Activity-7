using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing;
using System.Drawing.Imaging;

namespace Activity_4
{
    public partial class Reports : Form
    {
        private string connectionString = "Server=localhost;Database=rentaldb;Uid=root;";

        public Reports()
        {
            InitializeComponent();
        }

        private void LoadMonthlyData()
        {
            string query = "SELECT monthID, monthDate, monthSale FROM monthly;";
            LoadDataIntoDataGridView(query, dataGridView2);
        }

        private void LoadDailyData()
        {
            string query = "SELECT dailyID, dailyDate, dailySale FROM daily;";
            LoadDataIntoDataGridView(query, dataGridView1);
        }

        private void LoadDataIntoDataGridView(string query, DataGridView dataGridView)
        {
            DataTable dataTable = new DataTable();

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                MySqlCommand command = new MySqlCommand(query, connection);

                try
                {
                    connection.Open();
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(dataTable);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error loading data: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            dataGridView.DataSource = dataTable;

            // Set auto-size mode to fill
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // Adjust column headers height
            dataGridView.AutoResizeColumnHeadersHeight();

            // Ensure the DataGridView stretches in all directions
            dataGridView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Dashboard dash = new Dashboard();
            dash.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            About about = new About();
            about.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Hide();
            Login lg = new Login();
            lg.ShowDialog();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            LoadMonthlyData();
            LoadDailyData();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            this.Hide();
            Management manage = new Management();
            manage.ShowDialog();
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            Search sea = new Search();
            sea.ShowDialog();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            ExportDataAndChartToExcel();
        }

        private int GetRowCount(string query)
        {
            int rowCount = 0;

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                DataTable dataTable = new DataTable();

                MySqlCommand command = new MySqlCommand(query, connection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);

                try
                {
                    connection.Open();
                    adapter.Fill(dataTable);
                    rowCount = dataTable.Rows.Count;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error loading data: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return rowCount;
        }

        private void AddSignatureLine(ExcelWorksheet worksheet, int startRow)
        {
            // Row for underscores
            int underscoreRow = startRow;

            // Row for the signature
            int signatureRow = startRow + 1;

            // Add underscores at cell D11
            worksheet.Cells[underscoreRow, 4].Value = "________________________";

            string name = "KIMI CZAR L. SALTING";

            // Add the name at cell D12
            worksheet.Cells[signatureRow, 4].Value = name;
        }

        private void LoadDataIntoWorksheet(string query, ExcelWorksheet worksheet, int startRow)
        {
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                DataTable dataTable = new DataTable();

                MySqlCommand command = new MySqlCommand(query, connection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);

                try
                {
                    connection.Open();
                    adapter.Fill(dataTable);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error loading data: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                // Load data into worksheet starting from the specified row
                int rowCount = dataTable.Rows.Count;
                int colCount = dataTable.Columns.Count;

                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < colCount; j++)
                    {
                        if (dataTable.Columns[j].ColumnName == "monthDate" || dataTable.Columns[j].ColumnName == "dailyDate")
                        {
                            // Convert DateTime value to string with date format "MM/dd/yyyy"
                            worksheet.Cells[startRow + i, j + 1].Value = Convert.ToDateTime(dataTable.Rows[i][j]).ToString("MM/dd/yyyy");
                        }
                        else if (dataTable.Columns[j].ColumnName == "monthSale" || dataTable.Columns[j].ColumnName == "dailySale")
                        {
                            // Convert sales value to a number if possible
                            if (double.TryParse(dataTable.Rows[i][j].ToString(), out double sales))
                            {
                                // Set the value as a number
                                worksheet.Cells[startRow + i, j + 1].Value = sales;
                            }
                            else
                            {
                                // Set the value as string if conversion fails
                                worksheet.Cells[startRow + i, j + 1].Value = dataTable.Rows[i][j].ToString();
                            }
                        }
                        else
                        {
                            // Add other values to the cell as usual
                            worksheet.Cells[startRow + i, j + 1].Value = dataTable.Rows[i][j].ToString();
                        }
                    }
                }
            }
        }

        private void AddLogoToWorksheet(string logoPath, ExcelWorksheet worksheet)
        {
            if (File.Exists(logoPath))
            {
                ExcelPicture logo = worksheet.Drawings.AddPicture("Logo", new FileInfo(logoPath));
                logo.SetPosition(0, 0);
                logo.SetSize(200, 100);
            }
            else
            {
                MessageBox.Show("Logo not found at the specified path.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void ExportDataAndChartToExcel()
        {
            // Create a new Excel package
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                // Add worksheets for monthly and daily data
                ExcelWorksheet monthlyWorksheet = excelPackage.Workbook.Worksheets.Add("Monthly Data");
                ExcelWorksheet dailyWorksheet = excelPackage.Workbook.Worksheets.Add("Daily Data");
                ExcelWorksheet monthlyChartWorksheet = excelPackage.Workbook.Worksheets.Add("Monthly Chart");
                ExcelWorksheet dailyChartWorksheet = excelPackage.Workbook.Worksheets.Add("Daily Chart");

                // Load data into worksheets starting from the 6th row
                int startRowMonthly = 6;
                LoadDataIntoWorksheet("SELECT monthID, monthDate, monthSale FROM monthly;", monthlyWorksheet, startRowMonthly);
                AddSignatureLine(monthlyWorksheet, startRowMonthly + 1 + GetRowCount("SELECT monthID, monthDate, monthSale FROM monthly;"));

                int startRowDaily = 6;
                LoadDataIntoWorksheet("SELECT dailyID, dailyDate, dailySale FROM daily;", dailyWorksheet, startRowDaily);
                AddSignatureLine(dailyWorksheet, startRowDaily + 1 + GetRowCount("SELECT dailyID, dailyDate, dailySale FROM daily;"));

                // Add logo to data worksheets
                string logoPath = @"C:\\Users\\user\\Desktop\\BSIT stuff\\3rd year\\2nd sem\\IT 120 - EDP\\Activity 6\\Activity_4\\logo.png";
                AddLogoToWorksheet(logoPath, monthlyWorksheet);
                AddLogoToWorksheet(logoPath, dailyWorksheet);

                // Insert monthly and daily charts into their respective worksheets
                InsertChartIntoWorksheet(monthlyWorksheet, startRowMonthly, monthlyChartWorksheet, "Monthly Sales Report");
                InsertChartIntoWorksheet(dailyWorksheet, startRowDaily, dailyChartWorksheet, "Daily Sales Report");

                // Save the Excel package
                string filePath = "C:\\Users\\user\\Downloads\\Sales.xlsx";
                FileInfo excelFile = new FileInfo(filePath);
                excelPackage.SaveAs(excelFile);

                // Display a message to the user
                MessageBox.Show("Data and charts exported successfully!", "Export Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // Method to insert charts into a worksheet
        private void InsertChartIntoWorksheet(ExcelWorksheet sourceWorksheet, int startRow, ExcelWorksheet targetWorksheet, string chartTitle)
        {
            // Find the last row with data in column B (dates)
            int lastRowDate = sourceWorksheet.Dimension.End.Row;

            // Find the last row with data in column C (sales)
            int lastRowSales = sourceWorksheet.Dimension.End.Row;

            // Determine the range for the chart using the last row with data in both columns
            ExcelRangeBase rangeDates = sourceWorksheet.Cells["B" + startRow + ":B" + lastRowDate];
            ExcelRangeBase rangeSales = sourceWorksheet.Cells["C" + startRow + ":C" + lastRowSales];

            // Add the chart to the target worksheet
            var chart = targetWorksheet.Drawings.AddChart(chartTitle, eChartType.ColumnClustered);
            chart.SetPosition(1, 0, 0, 0);
            chart.SetSize(600, 400);

            // Load chart data from the determined ranges
            chart.Series.Add(rangeSales, rangeDates);

            // Set the title and axis labels for the chart
            chart.Title.Text = chartTitle;
            chart.XAxis.Title.Text = "Date";
            chart.YAxis.Title.Text = "Sales";
        }

    }
}