using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using MySql.Data.MySqlClient;

namespace Activity_4
{
    public partial class Dashboard : Form
    {
        public Dashboard()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Normal)
                this.WindowState = FormWindowState.Maximized;
            else
                this.WindowState = FormWindowState.Normal;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            Reports rep = new Reports();
            rep.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            About about = new About();
            about.ShowDialog();
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            // Display a confirmation message
            DialogResult result = MessageBox.Show("Are you sure you want to logout?", "Logout Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            // Check the user's response
            if (result == DialogResult.Yes)
            {
                string connectionString = "Server=localhost;Port=3306;Database=rentaldb;Uid=root;Pwd=rentuaya2022;";
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                {
                    try
                    {
                        connection.Open();
                        string selectQuery = "SELECT username FROM accounts WHERE activity = 'online'";
                        MySqlCommand selectCommand = new MySqlCommand(selectQuery, connection);
                        string loggedInUsername = selectCommand.ExecuteScalar()?.ToString();

                        if (!string.IsNullOrEmpty(loggedInUsername))
                        {
                            // Update the is_active field to false upon logout
                            string updateQuery = "UPDATE accounts SET activity = 'offline' WHERE username = @username";
                            MySqlCommand updateCommand = new MySqlCommand(updateQuery, connection);
                            updateCommand.Parameters.AddWithValue("@username", loggedInUsername);
                            updateCommand.ExecuteNonQuery();
                        }
                        else
                        {
                            MessageBox.Show("Error: Unable to retrieve logged-in user information.");
                        }
                    }
                    catch (MySqlException ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                }

                // Create an instance of the login form
                Login backloginForm = new Login();

                // Show the login form
                backloginForm.Show();

                // Close the current form (About)
                this.Close();
            }
            // If the user chooses No, do nothing
            /*
            this.Hide();
            Login lg = new Login();
            lg.ShowDialog();
            */
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.Hide();
            Search search = new Search();
            search.ShowDialog();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            this.Hide();
            Management manage = new Management();
            manage.ShowDialog();
        }
    }
}
