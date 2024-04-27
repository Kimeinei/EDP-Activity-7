using MySql.Data.MySqlClient;
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

namespace Activity_4
{
    public partial class NewPassword : Form
    {
        private const string connectionString = "Server=127.0.0.1;Uid=root;Database=rentaldb;";

        public NewPassword()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Login fpw = new Login();
            fpw.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string username = userBox.Text.Trim();
            string newPassword = textBox1.Text.Trim();
            string confirmPassword = textBox2.Text.Trim();

            if (newPassword != confirmPassword)
            {
                MessageBox.Show("Passwords do not match. Please try again.");
                return;
            }

            if (UpdatePassword(username, newPassword))
            {
                MessageBox.Show("Password updated successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Hide(); // Hide the current ForgotPass form
                Login log = new Login();
                log.Show();
            }
            else
            {
                MessageBox.Show("Passwords do not match", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool UpdatePassword(string username, string newPassword)
        {
            string query = "UPDATE accounts SET password = @newPassword WHERE username = @username";
            try
            {
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                MySqlCommand command = new MySqlCommand(query, connection);
                command.Parameters.AddWithValue("@username", username);
                command.Parameters.AddWithValue("@newPassword", newPassword);

                connection.Open();
                int rowsAffected = command.ExecuteNonQuery();

                return rowsAffected > 0;
            }

            }catch(MySqlException ex) 
            {
                MessageBox.Show("Error: " + ex.Message);
                return false;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void NewPassword_Load(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
