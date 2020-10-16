using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;

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
            HelperModel helper = new HelperModel();
            string filePath = ConfigurationManager.AppSettings["xlsxPath"];
            try
            {
                if (!string.IsNullOrWhiteSpace(txtUsername.Text) && !string.IsNullOrWhiteSpace(txtPassword.Text))
                {
                   
                    string userName = txtUsername.Text;
                    string password = txtPassword.Text;
                    bool radioInput = false;
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
                    Excel.Worksheet workSheet = workBook.Worksheets[3];
                    Excel.Range range = workSheet.UsedRange.Columns["A", Type.Missing];
                    int rowCount = range.Rows.Count;
                    for (int i = 2; i <= rowCount; i++)
                    {
                        if (workSheet.Cells[i, "B"] != null && workSheet.Cells[i, "B"].Value != null && workSheet.Cells[i, "D"] != null && workSheet.Cells[i, "D"].Value != null)
                        {
                            if (userName == workSheet.Cells[i, "B"].Value.ToString() && password == workSheet.Cells[i, "D"].Value.ToString())
                            {
                                if (rbnYes.IsChecked == true)
                                {
                                    radioInput = true;
                                    MainWindow window = new MainWindow(userName, radioInput);
                                    window.Show();
                                    this.Close();
                                    workBook.Close();
                                    xlApp.Quit();
                                    GC.Collect();
                                }
                                else if (rbnNo.IsChecked == true)
                                {
                                    radioInput = false;
                                    MainWindow window = new MainWindow(userName, radioInput);
                                    window.Show();
                                    this.Close();
                                    workBook.Close();
                                    xlApp.Quit();
                                    GC.Collect();
                                }
                                else
                                {
                                    MessageBox.Show("Please Select an Option");
                                    workBook.Close();
                                    xlApp.Quit();
                                    GC.Collect();
                                }
                            }



                        }

                    }
                }
                else
                {
                    MessageBox.Show("Please fill in all details");
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

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
