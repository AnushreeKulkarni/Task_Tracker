using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;

namespace TaskTrackerWPF
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        public Window1()
        {
            InitializeComponent();
            if(string.IsNullOrWhiteSpace(txtUsername.Text))
            {
                txtUsername.Text = "Please enter your ID";
            }
            if(string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                txtPassword.Text = "Please enter your password";
            }
         
        }

        private void LoginToTaskTracker(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(txtUsername.Text) && !string.IsNullOrWhiteSpace(txtPassword.Text) && txtUsername.Text.ToString() != "Please enter your ID" && txtPassword.Text.ToString() != "Please enter your password")
                {
                    HelperModel helper = new HelperModel();
                    List<UserInfo> userList;
                    userList = helper.BindEmployeeData();
                    string userName = txtUsername.Text;
                    string password = txtPassword.Text;
                    bool radioInput = false;
                    string access = "";
                    if (rbnYes.IsChecked==true)
                    {
                        access = "Yes";

                    }
                    else if(rbnNo.IsChecked==true)
                    {
                        access = "No";
                    }
                    var list = (from u in userList
                                where u.EmpId.Equals(userName) && u.Password.Equals(password) && u.AdminAccess.Equals(access)
                                select new { u.EmpId, u.Password,u.AdminAccess }).ToList();
                    if (list.Count != 0)
                    {
                        if (rbnYes.IsChecked == true)
                        {
                            
                            radioInput = true;
                            MainWindow window = new MainWindow(userName, radioInput);
                            window.Show();
                            this.Close();
                        }
                        else if (rbnNo.IsChecked == true)
                        {
                            radioInput = false;
                            MainWindow window = new MainWindow(userName, radioInput);
                            window.Show();
                            this.Close();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Username or password is incorrect or please verify your access type");
                        rbnNo.IsChecked = false;
                        rbnYes.IsChecked = false;
                        txtUsername.Text = "Please enter your ID";
                        txtPassword.Text = "Please enter your password";
                    }
                }
                else
                {
                    MessageBox.Show("Please fill in all details");

                }
            }
            catch
            {
                MessageBox.Show("Something went wrong");
            }


        }

        private void txtUsername_MouseEnter(object sender, MouseEventArgs e)
        {
            if (txtUsername.Text.ToString() == "Please enter your ID")
            {
                txtUsername.Text = "";
            }

        }

        private void txtUsername_MouseLeave(object sender, MouseEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtUsername.Text))
            {
                txtUsername.Text = "Please enter your ID";
            }
        }

        private void txtPassword_MouseEnter(object sender, MouseEventArgs e)
        {
            if (txtPassword.Text.ToString() == "Please enter your password")
            {
                txtPassword.Text = "";
            }
        }

        private void txtPassword_MouseLeave(object sender, MouseEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                txtPassword.Text = "Please enter your password";
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            txtUsername.Text = "Please enter your ID";
            txtPassword.Text = "Please enter your password";
            rbnNo.IsChecked = false;
            rbnYes.IsChecked = false;
        }
    }
}
