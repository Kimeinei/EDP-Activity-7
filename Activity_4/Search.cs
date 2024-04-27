using MySql.Data.MySqlClient;
using OfficeOpenXml.Drawing;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml.Drawing.Chart;

namespace Activity_4
{
    public partial class Search : Form
    {
        private string connectionString = "Server=localhost;Database=rentaldb;Uid=root;";
        public Search()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Dashboard dash = new Dashboard();
            dash.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            Reports rep = new Reports();
            rep.ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {

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

        private void button6_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void Search_Load(object sender, EventArgs e)
        {
            
        }

        private void button8_Click(object sender, EventArgs e)
        {
            this.Hide();
            Management manage = new Management();
            manage.ShowDialog();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {
            ExportRoomsDataToExcel();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string query = "SELECT * FROM rooms;";
            LoadDataIntoDataGridView(query, dataGridView1);
        }

        private void button10_Refresh_Click(object sender, EventArgs e)
        {
            UpdateDatabase();
            button10_Click(sender, e);
        }

        private void UpdateDatabase()
        {
            string updateQuery = "UPDATE rooms SET roomActivity = @roomActivity WHERE roomId = @roomId;";

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                MySqlCommand command = new MySqlCommand(updateQuery, connection);

                command.Parameters.Add("@roomActivity", MySqlDbType.VarChar);
                command.Parameters.Add("@roomId", MySqlDbType.Int32);

                try
                {
                    connection.Open();
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            string roomId = row.Cells["roomId"].Value.ToString();
                            string roomActivity = row.Cells["roomActivity"].Value.ToString();

                            command.Parameters["@roomActivity"].Value = roomActivity;
                            command.Parameters["@roomId"].Value = roomId;
                            command.ExecuteNonQuery();
                        }
                    }
                    MessageBox.Show("Data updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error updating data: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            // Define the SQL query to select rows with 'free' in the roomActivity column
            string query = "SELECT * FROM rooms WHERE roomActivity LIKE '%free%';";

            // Call the method to load data into dataGridView1
            LoadDataIntoDataGridView(query, dataGridView1);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            // Define the SQL query to select rows with 'free' in the roomActivity column
            string query = "SELECT * FROM rooms WHERE roomActivity LIKE '%occupied%';";

            // Call the method to load data into dataGridView1
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

            // Calculate the total width of all columns
            int totalColumnWidth = 0;
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                totalColumnWidth += column.Width;
            }

            // Get the width of the DataGridView minus the width of the vertical scrollbar (if visible)
            int availableWidth = dataGridView.ClientSize.Width - (dataGridView.RowHeadersVisible ? dataGridView.RowHeadersWidth : 0);
            if (dataGridView.ScrollBars == ScrollBars.Vertical)
            {
                availableWidth -= SystemInformation.VerticalScrollBarWidth;
            }

            // Calculate the zoom factor to fit the columns in the available width
            float zoomFactor = (float)availableWidth / totalColumnWidth;

            // Set the width of each column based on the calculated zoom factor
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                column.Width = (int)(column.Width * zoomFactor);
            }
        }
        private void ExportRoomsDataToExcel()
        {
            // Create a new Excel package
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                // Add a worksheet for room data
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Room Data");

                // Define the SQL query to select all columns from the rooms table
                string query = "SELECT * FROM rooms;";

                // Load data into the worksheet
                LoadDataIntoWorksheet(query, worksheet);

                // Add logo to the worksheet
                string logoPath = @"C:\\Users\\user\\Desktop\\BSIT stuff\\3rd year\\2nd sem\\IT 120 - EDP\\Activity 6\\Activity_4\\logo.png";
                AddLogoToWorksheet(logoPath, worksheet);

                // Add signature to the worksheet
                int signatureRow = worksheet.Dimension.End.Row + 1;
                AddSignatureLine(worksheet, signatureRow);

                // Auto-fit columns
                worksheet.Cells.AutoFitColumns();

                // Add a new worksheet for chart data
                ExcelWorksheet chartDataSheet = excelPackage.Workbook.Worksheets.Add("Chart Data");

                // Calculate the count of occupied and free rooms
                int occupiedCount = CountRoomsByActivity("occupied");
                int freeCount = CountRoomsByActivity("free");

                // Add headers and counts to the chart data sheet
                chartDataSheet.Cells["D6"].Value = "Room Status";
                chartDataSheet.Cells["E6"].Value = "Count";
                chartDataSheet.Cells["D7"].Value = "Occupied";
                chartDataSheet.Cells["E7"].Value = occupiedCount;
                chartDataSheet.Cells["D8"].Value = "Free";
                chartDataSheet.Cells["E8"].Value = freeCount;

                // Add pie chart to represent the count of occupied and free rooms
                AddPieChart(chartDataSheet, occupiedCount, freeCount);

                // Save the Excel package
                string filePath = "C:\\Users\\user\\Downloads\\RoomsData.xlsx"; // Adjust the file path as needed
                FileInfo excelFile = new FileInfo(filePath);
                excelPackage.SaveAs(excelFile);

                // Display a message to the user
                MessageBox.Show("Rooms data exported successfully!", "Export Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private int CountRoomsByActivity(string activity)
        {
            string query = $"SELECT COUNT(*) FROM rooms WHERE roomActivity LIKE '%{activity}%';";

            int count = 0;

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                MySqlCommand command = new MySqlCommand(query, connection);

                try
                {
                    connection.Open();
                    count = Convert.ToInt32(command.ExecuteScalar());
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error counting rooms: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return count;
        }

        private void AddPieChart(ExcelWorksheet worksheet, int occupiedCount, int freeCount)
        {
            // Add pie chart
            var chart = worksheet.Drawings.AddChart("Room Status", eChartType.PieExploded3D);
            chart.SetPosition(2, 0, 7, 0);
            chart.SetSize(400, 400);

            // Add data series
            var series = chart.Series.Add(worksheet.Cells["E7:E8"], worksheet.Cells["D7:D8"]);
            series.Header = "Room Status";

            // Set chart title and legend
            chart.Title.Text = "Room Status";
            chart.Legend.Position = eLegendPosition.Right;
        }


        private void LoadDataIntoWorksheet(string query, ExcelWorksheet worksheet)
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

            // Load data into the worksheet starting from row 6
            int startRow = 6;

            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    worksheet.Cells[i + startRow, j + 1].Value = dataTable.Rows[i][j].ToString();
                }
            }
        }

        private void AddSignatureLine(ExcelWorksheet worksheet, int startRow)
        {
            // Add underscores at cell D11
            worksheet.Cells[startRow, 4].Value = "________________________";

            string name = "KIMI CZAR L. SALTING";

            // Add the name at cell D12
            worksheet.Cells[startRow + 1, 4].Value = name;
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

    }
}
