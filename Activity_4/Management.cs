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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Activity_4
{
    public partial class Management : Form
    {
        private string connectionString = "Server=localhost;Database=rentaldb;Uid=root;";
        private DataTable _userData;
        private DataGridView dataGridView1;

        public Management()
        {
            InitializeComponent();
            InitializeDataGridView();
            _userData = new DataTable();
        }

        private void InitializeDataGridView()
        {
            // Create a new instance of DataGridView
            dataGridView1 = new DataGridView();

            // Set properties of the DataGridView
            dataGridView1.Name = "dataGridView1";
            dataGridView1.Dock = DockStyle.Fill;

            // Add the DataGridView to the form's Controls collection
            panel3.Controls.Add(dataGridView1);
            dataGridView1.Location = new Point(10, 10);

            dataGridView1.CellClick += DataGridView1_CellClick;
        }

        private void button8_Click(object sender, EventArgs e)
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

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Dashboard dash = new Dashboard();
            dash.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            Reports report = new Reports();
            report.ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.Hide();
            Search search = new Search();
            search.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            About ab = new About();
            ab.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Hide();
            Login log = new Login();
            log.ShowDialog();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Check if the clicked cell belongs to a valid row
            if (e.RowIndex >= 0 && e.RowIndex < dataGridView1.Rows.Count)
            {
                // Get the selected row
                DataGridViewRow selectedRow = dataGridView1.Rows[e.RowIndex];

                // Extract data from the selected row and populate TextBoxes and DateTimePicker
                textBox2.Text = selectedRow.Cells["account_id"].Value.ToString(); // Assuming the column name is "account_id"
                textBox3.Text = selectedRow.Cells["username"].Value.ToString();
                textBox4.Text = selectedRow.Cells["password"].Value.ToString();
                dateTimePicker1.Value = Convert.ToDateTime(selectedRow.Cells["date_applied"].Value); // Assuming the column name is "date_applied"
            }
        }

        private void UpdateChanges()
        {
            string accountIdEdit = textBox2.Text;
            string usernameEdit = textBox3.Text;
            string passwordEdit = textBox4.Text;
            DateTime dateAppliedEdit = dateTimePicker1.Value;

            // Update the DataRow in _userData with the new values
            foreach (DataRow row in _userData.Rows)
            {
                if (row["account_id"].ToString() == accountIdEdit)
                {
                    row["username"] = usernameEdit;
                    row["password"] = passwordEdit;
                    row["date_applied"] = dateAppliedEdit;

                    break;
                }
            }

            // Refresh the DataGridView to reflect the changes
            dataGridView1.DataSource = _userData;

            // Update the database with the changes
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                string query = "UPDATE accounts SET username = @username, password = @password, date_applied = @dateApplied WHERE account_id = @accountId";
                MySqlCommand command = new MySqlCommand(query, connection);
                command.Parameters.AddWithValue("@username", usernameEdit);
                command.Parameters.AddWithValue("@password", passwordEdit);
                command.Parameters.AddWithValue("@dateApplied", dateAppliedEdit);
                command.Parameters.AddWithValue("@accountId", accountIdEdit);

                try
                {
                    connection.Open();
                    int rowsAffected = command.ExecuteNonQuery();
                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Account information updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("No rows were updated.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error updating account information: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string searchText = textBox1.Text.Trim();

            if (string.IsNullOrEmpty(searchText))
            {
                // If the search text is empty, reload all data
                LoadUserData();
            }
            else
            {
                // Filter the DataTable based on search text
                DataTable filteredTable = _userData.Clone(); // Create a clone of the original DataTable structure
                foreach (DataRow row in _userData.Rows)
                {
                    if (row["username"].ToString().Contains(searchText))
                    {
                        filteredTable.ImportRow(row);
                    }
                }

                // Update the DataGridView with filtered data
                dataGridView1.DataSource = filteredTable;
            }
        }


        private void LoadUserData()
        {
            _userData = new DataTable(); // Initialize _userData

            // Retrieve user data from the database and populate _userData
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                string query = "SELECT account_id, username, password, date_applied, activity FROM accounts";
                MySqlCommand command = new MySqlCommand(query, connection);

                try
                {
                    connection.Open();
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(_userData);
                    dataGridView1.DataSource = _userData; // Bind _userData to DataGridView
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error loading user data: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            UpdateChanges();
        }
    }
}
