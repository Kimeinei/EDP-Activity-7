using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using MySql.Data.MySqlClient;

namespace Activity_4
{
    public partial class Login : Form
    {
        private const string connectionString = "Server=127.0.0.1;Uid=root;Database=rentaldb;";
        public static string Username;

        public Login()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            Login ffp = new Login();
            ffp.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string enteredUsername = textBox1.Text.Trim();
            string enteredPassword = textBox2.Text.Trim();

            if (AuthenticateUser(enteredUsername, enteredPassword))
            {
                this.Hide();
                Dashboard dash = new Dashboard();
                dash.Show();
            }
            else
            {
                MessageBox.Show("Invalid Username or Password", "Failed to Login", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool AuthenticateUser(string username, string password)
        {
            bool isAuthenticated = false;

            try
            {
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                {
                    string query = "SELECT COUNT(*) FROM accounts WHERE username = @username AND password = @password";
                    MySqlCommand command = new MySqlCommand(query, connection);
                    command.Parameters.AddWithValue("@username", username);
                    command.Parameters.AddWithValue("@password", password);

                    connection.Open();
                    int count = Convert.ToInt32(command.ExecuteScalar());
                    if (count > 0)
                    {
                        // Retrieve the user's data including the "activity" column
                        string userDataQuery = "SELECT activity FROM accounts WHERE username = @username";
                        MySqlCommand userDataCommand = new MySqlCommand(userDataQuery, connection);
                        userDataCommand.Parameters.AddWithValue("@username", username);

                        MySqlDataAdapter adapter = new MySqlDataAdapter(userDataCommand);
                        DataTable userData = new DataTable();
                        adapter.Fill(userData);

                        // Check if the user's activity is "online"
                        if (userData.Rows.Count > 0 && userData.Rows[0]["activity"].ToString() == "online")
                        {
                            // Update the activity status to "online"
                            string updateQuery = "UPDATE accounts SET activity = 'online' WHERE username = @username";
                            MySqlCommand updateCommand = new MySqlCommand(updateQuery, connection);
                            updateCommand.Parameters.AddWithValue("@username", username);
                            updateCommand.ExecuteNonQuery();

                            isAuthenticated = true;
                        }
                        else
                        {
                            MessageBox.Show("This account is currently inactive. Please contact support for assistance.", "Account Inactive", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid Username or Password", "Failed to Login", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                return isAuthenticated;
            }
            catch (Exception ex)
            {
                // Handle the exception (e.g., log the error)
                MessageBox.Show("An error occurred while authenticating user: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return isAuthenticated;
        }


        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {
            this.Hide();
            NewPassword forgot = new NewPassword();
            forgot.Show();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
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

        private void Login_Load(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {
            this.Hide();
            SignUp sign = new SignUp();
            sign.Show();
        }
    }
}