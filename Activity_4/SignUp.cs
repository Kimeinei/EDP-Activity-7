using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace Activity_4
{
    public partial class SignUp : Form
    {
        string connectionString = "Server=localhost;Port=3306;Database=rentaldb;Uid=root;";
        public SignUp()
        {
            InitializeComponent();
        }

        private void SignUp_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            Login lg = new Login();
            lg.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string accountId = textBox1.Text;
            string username = textBox2.Text;
            string password = textBox3.Text;
            DateTime dateApplied = dateTimePicker1.Value;

            if (string.IsNullOrEmpty(accountId) || string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Please fill in all required fields.", "Required", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (CreateUserAccount(accountId, username, password, dateApplied))
            {
                MessageBox.Show("Sign up successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Hide();
                Login formLogin = new Login();
                formLogin.Show();
            }
            else
            {
                MessageBox.Show("Failed to sign up. Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool CreateUserAccount(string accountId, string username, string password, DateTime dateApplied)
        {
            try
            {
                // Use MySqlConnection instead of SqlConnection
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                {
                    connection.Open();
                    string query = "INSERT INTO accounts (account_id, username, password, date_applied, activity) VALUES (@accountId, @username, @password, @dateApplied, 'online')";

                    // Use MySqlCommand instead of SqlCommand
                    MySqlCommand command = new MySqlCommand(query, connection);
                    command.Parameters.AddWithValue("@accountId", accountId);
                    command.Parameters.AddWithValue("@username", username);
                    command.Parameters.AddWithValue("@password", password);
                    command.Parameters.AddWithValue("@dateApplied", dateApplied);
                    int rowsAffected = command.ExecuteNonQuery();
                    return rowsAffected > 0;
                }
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
    }
}
