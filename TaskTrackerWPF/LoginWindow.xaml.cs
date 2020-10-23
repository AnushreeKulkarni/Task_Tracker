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
        private EventHandler handler;
        public Window1()
        {
            InitializeComponent();
            if(string.IsNullOrWhiteSpace(txtUsername.Text))
            {
                txtUsername.Text = "Please enter your ID";
            }
            if(string.IsNullOrWhiteSpace(txtPassword.Password))
            {
                txtPassword.Password = "Please enter your password";
            }
            Int32 timeout = 300;
            handler = delegate
            {
                System.Windows.Threading.DispatcherTimer timer = new System.Windows.Threading.DispatcherTimer();
                timer.Interval = TimeSpan.FromSeconds(timeout);
                timer.Tick += delegate
                {
                    if (timer != null)
                    {
                        timer.Stop();
                        timer = null;
                        System.Windows.Interop.ComponentDispatcher.ThreadIdle -= handler;
                        System.Windows.Interop.ComponentDispatcher.ThreadIdle += handler;
                        System.Windows.Application.Current.Shutdown();

                    }
                };
                timer.Start();
                System.Windows.Threading.Dispatcher.CurrentDispatcher.Hooks.OperationPosted += delegate
                {
                    if (timer != null)
                    {
                        timer.Stop();
                        timer = null;
                    }
                };


            };
            System.Windows.Interop.ComponentDispatcher.ThreadIdle += handler;

        }

        private void LoginToTaskTracker(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(txtUsername.Text) && !string.IsNullOrWhiteSpace(txtPassword.Password) && txtUsername.Text.ToString() != "Please enter your ID" && txtPassword.Password.ToString() != "Please enter your password")
                {
                    HelperClass helper = new HelperClass();
                    List<UserInfo> userList;
                    userList = helper.BindEmployeeData();
                    string userName = txtUsername.Text;
                    string password = txtPassword.Password;
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
                        txtPassword.Password = "Please enter your password";
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
            if (txtPassword.Password.ToString() == "Please enter your password")
            {
                txtPassword.Password = "";
            }
        }

        private void txtPassword_MouseLeave(object sender, MouseEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPassword.Password))
            {
                txtPassword.Password = "Please enter your password";
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }
    }
}
