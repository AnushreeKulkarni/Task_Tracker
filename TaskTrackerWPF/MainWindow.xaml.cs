using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.Data;
using System.Windows.Threading;

namespace TaskTrackerWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private EventHandler handler;
       
        string filePath = ConfigurationManager.AppSettings["xlsxPath"];
        HelperClass helpers = new HelperClass();
        string id;
        public static bool isAdmin;
        public MainWindow(string _id,bool _isAdmin)
        {
            InitializeComponent();
            id = _id;
            isAdmin = _isAdmin;
            Excel.Application xlApp = new Excel.Application();
            List<UserInfo> employeeList;
            employeeList = helpers.BindEmployeeData();
            List<ModelTaskTracker> dailyList;
            dailyList = helpers.GetDailyTaskList();
            List<ModelTask> taskList;
            taskList = helpers.GetTaskList(id, isAdmin); 
            Object result;
            result = helpers.DailyTaskList(dailyList, employeeList, taskList);
            if (isAdmin == false)
            {
                empTab.Visibility = Visibility.Hidden;
                reportTab.Visibility = Visibility.Hidden;
                dailytrackGrid.Columns[0].Visibility = Visibility.Visible;
                //Populating Daily Task list in the Datagrid
                dailytrackGrid.ItemsSource = (System.Collections.IEnumerable)result;
                taskGrid.Columns[7].Visibility = Visibility.Visible;
                //Populating Task data in Datagrid
                taskGrid.ItemsSource = helpers.GetTaskList(id, isAdmin);

            }
            else
            {
                //Populating Daily Task list in the Datagrid
                dailytrackGrid.ItemsSource = (System.Collections.IEnumerable)result;
                lblemp.Visibility = Visibility.Visible;
                dropdownEmp.Visibility = Visibility.Visible;
                dropdownEmp.ItemsSource = employeeList;
                //Populating Task data in Datagrid
                taskGrid.ItemsSource = helpers.GetTaskList(id, isAdmin);
                //Binding of Employee list to the Datagrid
                empGrid.ItemsSource = helpers.BindEmployeeData();
            }
            dropdownTask.ItemsSource = taskList.Distinct();
            dropdownState.Items.Add("NEW");
            dropdownState.Items.Add("IN PROGRESS");
            dropdownState.Items.Add("COMPLETED");
            dropdownState.Items.Add("BLOCKED");
            dropdownPriority.Items.Add("LOW");
            dropdownPriority.Items.Add("MEDIUM");
            dropdownPriority.Items.Add("HIGH");
            dropdownTaskType.Items.Add("Bug");
            dropdownTaskType.Items.Add("Feature");
            dropdownTaskType.Items.Add("Daily Task");
            dropdownTaskType.Items.Add("Weekly Task");
            dropdownTaskType.Items.Add("Monthly Task");
            dropdownTaskType.Items.Add("Other");
            combo.Items.Add("ALL");
            if(string.IsNullOrWhiteSpace(txtDate.Text))
            {
                txtDate.Text = "dd/mm/yyyy";
            }
            if(string.IsNullOrWhiteSpace(txtPSD.Text))
            {
                txtPSD.Text = "dd/mm/yyyy";
            }
            if(string.IsNullOrWhiteSpace(txtPED.Text))
            {
                txtPED.Text = "dd/mm/yyyy";
            }
            if(string.IsNullOrWhiteSpace(txtASD.Text))
            {
                txtASD.Text = "dd/mm/yyyy";
            }
            if(string.IsNullOrWhiteSpace(txtAED.Text))
            {
                txtAED.Text = "dd/mm/yyyy";
            }
            if(string.IsNullOrWhiteSpace(txtSD.Text))
            {
                txtSD.Text = "dd/mm/yyyy";
            }
            if(string.IsNullOrWhiteSpace(txtED.Text))
            {
                txtED.Text = "dd/mm/yyyy";
            }
            xlApp.Quit();
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
        public void Reload()
        {
            try
            {
                List<UserInfo> employeeList;
                employeeList = helpers.BindEmployeeData();
                List<ModelTaskTracker> dailyList;
                dailyList = helpers.GetDailyTaskList();
                List<ModelTask> taskList;
                taskList = helpers.GetTaskList(id, isAdmin);
                helpers.DailyTaskList(dailyList, employeeList, taskList);
                Object result;
                result = helpers.DailyTaskList(dailyList, employeeList, taskList);
                if (isAdmin == false)
                {
                    dailytrackGrid.ItemsSource = (System.Collections.IEnumerable)result;
                    taskGrid.ItemsSource = helpers.GetTaskList(id, isAdmin);
                    List<ModelTask> tasklist;
                    tasklist = helpers.GetTaskList(id, isAdmin);
                    var taskid = (from ta in tasklist
                                  select ta.TaskId).Last();
                    int Tid = Convert.ToInt32(taskid);
                    txtTaskId.Text = (Tid + 1).ToString();
                    txtTaskId.IsReadOnly = true;
                }
                else
                {
                    dailytrackGrid.ItemsSource = (System.Collections.IEnumerable)result;
                    taskGrid.ItemsSource = helpers.GetTaskList(id, isAdmin);
                    empGrid.ItemsSource = helpers.BindEmployeeData();
                    List<ModelTask> tasklist;
                    tasklist = helpers.GetTaskList(id, isAdmin);
                    var taskid = (from ta in tasklist
                                  select ta.TaskId).Last();

                    int Tid = Convert.ToInt32(taskid);
                    txtTaskId.Text = (Tid + 1).ToString();
                    txtTaskId.IsReadOnly = true;

                }
                dropdownTask.ItemsSource = taskList;
            }
            catch
            {
                MessageBox.Show("Something went wrong");
            }
        }
     


        //Method to add details of new employee
        private void AddEmployee(object sender, RoutedEventArgs e)
        {      
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[3];
            Excel.Range range = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = range.Rows.Count;
            int i = rowCount + 1;
            try
            {
                if (!string.IsNullOrWhiteSpace(txtName.Text) && !string.IsNullOrWhiteSpace(txtId.Text) && !string.IsNullOrWhiteSpace(txtUa.Text) && !string.IsNullOrWhiteSpace(txtPwd.Password) &&(rbnYes.IsChecked==true || rbnNo.IsChecked==true ))
                {
                    if (i > rowCount)
                    {

                        workSheet.Cells[i, "A"].Value = txtName.Text.ToString();
                        workSheet.Cells[i, "B"].Value = txtId.Text.ToString();
                        workSheet.Cells[i, "C"].Value = txtUa.Text.ToString();
                        workSheet.Cells[i, "D"].Value = txtPwd.Password.ToString();
                        if(rbnYes.IsChecked==true)
                        {
                            workSheet.Cells[i, "E"].Value = "Yes";
                        }
                        else if(rbnNo.IsChecked==true)
                        {
                            workSheet.Cells[i, "E"].Value = "No";
                        }
                    }

                    MessageBox.Show("Details added successfully");
                    txtId.Text = "";
                    txtName.Text = "";
                    txtUa.Text = "";
                    txtPwd.Password = "";
                    rbnYes.IsChecked = false;
                    rbnNo.IsChecked = false;
                    workBook.Save();
                    workBook.Close();
                    xlApp.Quit();
                    GC.Collect();
                    empGrid.ItemsSource = helpers.BindEmployeeData();
                }
                else
                {
                    MessageBox.Show("Please enter all details");
                    workBook.Save();
                    workBook.Close();
                    xlApp.Quit();
                    GC.Collect();
                }
            }
            catch (Exception ex)
            {
               MessageBox.Show(ex.Message);
            }

        }
        //Method to add task data of a new task
        public void AddTaskData(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dropdownState.SelectedItem.ToString() == "COMPLETED")
                {
                    if (string.IsNullOrWhiteSpace(txtASD.Text) && string.IsNullOrWhiteSpace(txtAED.Text) || (txtASD.Text.ToString() == "dd/mm/yyyy" && txtAED.Text.ToString() == "dd/mm/yyyy"))
                    {
                        
                         MessageBox.Show("Please enter Actual Start and End dates");
                        
                    }
                    else
                    {
                        try
                        {
                            if (!string.IsNullOrWhiteSpace(txtTicket.Text) && !string.IsNullOrWhiteSpace(txtTaskTitle.Text) && !string.IsNullOrWhiteSpace(txtTaskDesc.Text) && !string.IsNullOrEmpty(dropdownState.Text) && !string.IsNullOrEmpty(dropdownPriority.Text) && !string.IsNullOrWhiteSpace(txtEfforts.Text) && !string.IsNullOrWhiteSpace(txtPSD.Text) && !string.IsNullOrWhiteSpace(txtPED.Text))
                            {
                                Excel.Application xlApp = new Excel.Application();
                                Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
                                Excel.Worksheet workSheet = workBook.Worksheets[2];
                                Excel.Range range = workSheet.UsedRange.Columns["A", Type.Missing];
                                int rowCount = range.Rows.Count;
                                int i = rowCount + 1;
                                if (i > rowCount)
                                {
                                    if (txtPSD.Text.ToString() != "dd/mm/yyyy" && txtPED.Text.ToString() != "dd/mm/yyyy")
                                    {
                                     
                                        workSheet.Cells[i, "A"].Value = txtTicket.Text.ToString();
                                        workSheet.Cells[i, "B"].Value = txtTaskId.Text.ToString();
                                        workSheet.Cells[i, "C"].Value = txtTaskTitle.Text.ToString();
                                        workSheet.Cells[i, "D"].Value = txtTaskDesc.Text.ToString();
                                        workSheet.Cells[i, "E"].Value = dropdownTaskType.SelectedItem.ToString();
                                        workSheet.Cells[i, "F"].Value = dropdownState.SelectedItem.ToString();
                                        workSheet.Cells[i, "G"].Value = dropdownPriority.SelectedItem.ToString();
                                        workSheet.Cells[i, "H"].Value = id;
                                        workSheet.Cells[i, "I"].Value = txtEfforts.Text.ToString();
                                        workSheet.Cells[i, "j"].Value = txtPSD.Text.ToString();
                                        workSheet.Cells[i, "K"].Value = txtPED.Text.ToString();
                                        workSheet.Cells[i, "L"].Value = txtASD.Text.ToString();
                                        workSheet.Cells[i, "M"].Value = txtAED.Text.ToString();
                                        MessageBox.Show("Details Added Successfully");
                                        txtTicket.Text = "";
                                        txtTaskId.Text = "";
                                        txtTaskTitle.Text = "";
                                        txtTaskDesc.Text = "";
                                        txtEfforts.Text = "";
                                        txtPSD.Text = "";
                                        txtPED.Text = "";
                                        txtASD.Text = "";
                                        txtAED.Text = "";
                                        txtTicket.IsReadOnly = false;
                                        txtTaskTitle.IsReadOnly = false;
                                        txtTaskDesc.IsReadOnly = false;
                                        dropdownPriority.IsEnabled = true;
                                        txtPSD.IsReadOnly = false;
                                        txtPED.IsReadOnly = false;
                                        txtEfforts.IsReadOnly = false;
                                        workBook.Save();
                                        workBook.Close();
                                        xlApp.Quit();
                                        GC.Collect();
                                        Reload();
                                    }
                                    else
                                    {
                                        MessageBox.Show("Please enter Planned Start and End Dates");
                                        workBook.Save();
                                        workBook.Close();
                                        xlApp.Quit();
                                        GC.Collect();

                                    }        
                                }
                       
                            }
                            else
                            {
                                MessageBox.Show("Fill in all the details");
                            
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                    }

                }
                else if (dropdownState.SelectedItem.ToString() != "COMPLETED")
                {
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(txtTicket.Text) && !string.IsNullOrWhiteSpace(txtTaskId.Text) && !string.IsNullOrWhiteSpace(txtTaskTitle.Text) && !string.IsNullOrWhiteSpace(txtTaskDesc.Text) && !string.IsNullOrEmpty(dropdownState.Text) && !string.IsNullOrEmpty(dropdownPriority.Text) && !string.IsNullOrWhiteSpace(txtEfforts.Text) && !string.IsNullOrWhiteSpace(txtPSD.Text) && !string.IsNullOrWhiteSpace(txtPED.Text))
                        {
                            Excel.Application xlApp = new Excel.Application();
                            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
                            Excel.Worksheet workSheet = workBook.Worksheets[2];
                            Excel.Range range = workSheet.UsedRange.Columns["A", Type.Missing];
                            int rowCount = range.Rows.Count;
                            int i = rowCount + 1;
                            if (i > rowCount)
                            {
                                if (txtPSD.Text.ToString() != "dd/mm/yyyy" && txtPED.Text.ToString() != "dd/mm/yyyy")
                                {

                                    workSheet.Cells[i, "A"].Value = txtTicket.Text.ToString();
                                    workSheet.Cells[i, "B"].Value = txtTaskId.Text.ToString();
                                    workSheet.Cells[i, "C"].Value = txtTaskTitle.Text.ToString();
                                    workSheet.Cells[i, "D"].Value = txtTaskDesc.Text.ToString();
                                    workSheet.Cells[i, "E"].Value = dropdownTaskType.SelectedItem.ToString();
                                    workSheet.Cells[i, "F"].Value = dropdownState.SelectedItem.ToString();
                                    workSheet.Cells[i, "G"].Value = dropdownPriority.SelectedItem.ToString();
                                    workSheet.Cells[i, "H"].Value = id;
                                    workSheet.Cells[i, "I"].Value = txtEfforts.Text.ToString();
                                    workSheet.Cells[i, "j"].Value = txtPSD.Text.ToString();
                                    workSheet.Cells[i, "K"].Value = txtPED.Text.ToString();
                                    if (txtASD.Text.ToString() != "dd/mm/yyyy")
                                    {

                                        workSheet.Cells[i, "L"].Value = txtASD.Text.ToString();
                                    }
                                    else
                                    {
                                        workSheet.Cells[i, "L"].Value = "";
                                    }
                                    workSheet.Cells[i, "M"].Value = "";
                                    MessageBox.Show("Details Added Successfully");
                                    txtTicket.Text = "";
                                    txtTaskId.Text = "";
                                    txtTaskTitle.Text = "";
                                    txtTaskDesc.Text = "";
                                    txtEfforts.Text = "";
                                    txtPSD.Text = "";
                                    txtPED.Text = "";
                                    txtASD.Text = "";
                                    txtAED.Text = "";
                                    dropdownTaskType.Text = "";
                                    dropdownPriority.Text = "";
                                    txtTicket.IsReadOnly = false;
                                    txtTaskTitle.IsReadOnly = false;
                                    txtTaskDesc.IsReadOnly = false;
                                    dropdownPriority.IsEnabled = true;
                                    txtPSD.IsReadOnly = false;
                                    txtPED.IsReadOnly = false;
                                    txtEfforts.IsReadOnly = false;
                                    workBook.Save();
                                    workBook.Close();
                                    xlApp.Quit();
                                    GC.Collect();
                                    Reload();
                                }
                                else
                                {
                                    MessageBox.Show("Please enter Planned Start and End Dates");
                                    workBook.Save();
                                    workBook.Close();
                                    xlApp.Quit();
                                    GC.Collect();
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Fill in all the details");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

            }
            catch
            {
                MessageBox.Show("Fill all details");

            }

        }
        //Method to update an existing task
        public void UpdateTaskData(object sender, RoutedEventArgs e)
        {
            try
            {
                if (dropdownState.SelectedItem.ToString() == "COMPLETED")
                {
                    if (string.IsNullOrWhiteSpace(txtASD.Text) && string.IsNullOrWhiteSpace(txtAED.Text))
                    {
                        MessageBox.Show("Please enter Actual Start and End dates");
                    }
                    else
                    {
                        ModelTask model = new ModelTask();
                        model.TicketNumber = txtTicket.Text.ToString();
                        model.TaskId = txtTaskId.Text.ToString();
                        model.TaskTitle = txtTaskTitle.Text.ToString();
                        model.TaskDescription = txtTaskDesc.Text.ToString();
                        model.TaskType = dropdownTaskType.SelectedItem.ToString();
                        model.State = dropdownState.SelectedItem.ToString();
                        model.Priority = dropdownPriority.SelectedItem.ToString();
                        model.AssignedTo = id;
                        model.Efforts = txtEfforts.Text.ToString();
                        if (txtPED.Text.ToString() != "dd/mm/yyyy" && txtPSD.Text.ToString() != "dd/mm/yyyy")
                        {
                            model.PlannedStartDate = txtPSD.Text.ToString();
                            model.PlannedEndDate = txtPED.Text.ToString();
                        }
                        if (txtASD.Text.ToString() != "dd/mm/yyyy" && txtAED.Text.ToString() != "dd/mm/yyyy")
                        {
                            model.ActualStartDate = txtASD.Text.ToString();
                            model.ActualEndDate = txtAED.Text.ToString();
                            helpers.UpdateTask(model);
                            MessageBox.Show("Details Updated Successfully");
                            txtTicket.Text = "";
                            txtTaskId.Text = "";
                            txtTaskTitle.Text = "";
                            txtTaskDesc.Text = "";
                            txtEfforts.Text = "";
                            txtPSD.Text = "";
                            txtPED.Text = "";
                            txtASD.Text = "";
                            txtAED.Text = "";
                            Reload();
                        }
                       else
                        {
                            MessageBox.Show("Please enter Actual Start and End Dates");
                        }

                    }

                }
                else
                {
                    ModelTask model = new ModelTask();
                    model.TicketNumber = txtTicket.Text.ToString();
                    model.TaskId = txtTaskId.Text.ToString();
                    model.TaskTitle = txtTaskTitle.Text.ToString();
                    model.TaskDescription = txtTaskDesc.Text.ToString();
                    model.TaskType = dropdownTaskType.SelectedItem.ToString();
                    model.State = dropdownState.SelectedItem.ToString();
                    model.Priority = dropdownPriority.SelectedItem.ToString();
                    model.AssignedTo = id;
                    model.Efforts = txtEfforts.Text.ToString();
                    model.PlannedStartDate = txtPSD.Text.ToString();
                    model.PlannedEndDate = txtPED.Text.ToString();
                    if (txtASD.Text.ToString() != "dd/mm/yyyy" && txtAED.Text.ToString() != "dd/mm/yyyy")
                    {
                        model.ActualStartDate = txtASD.Text.ToString();
                        model.ActualEndDate = txtAED.Text.ToString();
                  
                    }
                    else
                    {
                        model.ActualStartDate = "";
                        model.ActualEndDate = "";
                    }
                    helpers.UpdateTask(model);
                    MessageBox.Show("Details Updated Successfully");
                    txtTicket.Text = "";
                    txtTaskId.Text = "";
                    txtTaskTitle.Text = "";
                    txtTaskDesc.Text = "";

                    txtEfforts.Text = "";
                    txtPSD.Text = "";
                    txtPED.Text = "";
                    txtASD.Text = "";
                    txtAED.Text = "";
                    Reload();
                }
            }
            catch
            {
                MessageBox.Show("Something went wrong");
            }
        }
        //Method for generating a report
        public void Generate_Report(object sender, RoutedEventArgs e)
        {
            string reportPath = ConfigurationManager.AppSettings["reportPath"]+"Analysis_Report"+DateTime.Now.ToShortDateString()+".xlsx";
            try
            {
                if (!string.IsNullOrEmpty(txtSD.Text) && !string.IsNullOrWhiteSpace(txtED.Text) && txtSD.Text.ToString() != "dd/mm/yyyy" && txtED.Text.ToString() != "dd/mm/yyyy")
                {
                    if (combo.SelectedItem.ToString() == "ALL")
                    {
                        List<UserInfo> list2;
                        list2 = helpers.BindEmployeeData();
                        List<ModelTaskTracker> list1;
                        list1 = helpers.GetDailyTaskList();
                        List<ModelTask> list3;
                        list3 = helpers.GetTaskList(id, isAdmin);
                        List<ReportGenerationModel> report;
                        string startDate = "";
                        string endDate = "";
                            startDate = txtSD.Text.ToString();
                            endDate = txtED.Text.ToString();

                            report = helpers.GenerateReport(list1, list2, list3, startDate, endDate);
                            try
                            {
                                Excel.Application app = new Excel.Application();
                                Excel.Workbook book = app.Workbooks.Add();
                                Excel.Worksheet work = book.Worksheets[1];
                                Excel.Range ra = work.UsedRange.Columns["A", Type.Missing];
                                int count = ra.Rows.Count;
                                work.Name = "Report";
                                work.Cells[1, "A"] = "Employee Name";
                                work.Cells[1, "B"] = "Bug";
                                work.Cells[1, "C"] = "Feature";
                                work.Cells[1, "D"] = "Daily Task";
                                work.Cells[1, "E"] = "Weekly Task";
                                work.Cells[1, "F"] = "Monthly Task";
                                work.Cells[1, "G"] = "Other";
                                work.Cells[1, "H"] = "Total";
                                work.Cells[1, "I"] = "Total Completed";
                                foreach (var re in report)
                                {
                                    count = count + 1;

                                    int i = count;

                                    work.Cells[i, "A"].Value = re.EmpName;
                                    work.Cells[i, "B"].Value = re.BugCount;
                                    work.Cells[i, "C"].Value = re.FeatureCount;
                                    work.Cells[i, "D"].Value = re.DailyTaskCount;
                                    work.Cells[i, "E"].Value = re.WeeklyTaskCount;
                                    work.Cells[i, "F"].Value = re.MonthlyTaskCount;
                                    work.Cells[i, "G"].Value = re.OthersCount;
                                    work.Cells[i, "H"].Value = re.TotalCount;
                                    work.Cells[i, "I"].Value = re.TotalCompletedCount;
                                }
                                book.SaveAs(reportPath);
                                book.Close();
                                app.Quit();
                                MessageBox.Show("Report Generated");
                                txtSD.Text = "";
                                txtED.Text = "";

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                    }
                    else if (combo.SelectedItem.ToString() != "ALL")
                    {
                        MessageBox.Show("Select the option");
                    }
               
                }
                else
                {
                    MessageBox.Show("Please fill all the fields");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //Method to add a new daily task 
        public void AddDailyTask(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[1];
            Excel.Range range = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = range.Rows.Count;
            int i = rowCount + 1;
            try
            {
                if (i > rowCount && isAdmin==false)
                {
                    workSheet.Cells[i, "B"].Value = id;
                }
                else if(i > rowCount && isAdmin == true)
                {
                    workSheet.Cells[i, "B"].Value = dropdownEmp.SelectedValue.ToString();
                }
                if (!string.IsNullOrWhiteSpace(dropdownTask.Text) && !string.IsNullOrWhiteSpace(txtHours.Text) && !string.IsNullOrWhiteSpace(txtDate.Text) && !string.IsNullOrWhiteSpace(txtRemarks.Text))
                {
                    if (txtDate.Text.ToString() != "dd/mm/yyyy")
                    {
                        workSheet.Cells[i, "C"].Value = txtDate.Text.ToString();
                        workSheet.Cells[i, "D"].Value = dropdownTask.SelectedValue.ToString();
                        workSheet.Cells[i, "E"].Value = txtHours.Text.ToString();
                        workSheet.Cells[i, "F"].Value = txtRemarks.Text.ToString();
                        MessageBox.Show("Details added successfully");
                        txtDate.Text = "";
                        txtHours.Text = "";
                        txtRemarks.Text = "";
                        workBook.Save();
                        workBook.Close();
                        xlApp.Quit();
                        GC.Collect();
                        Reload();
                    }
                    else
                    {
                        MessageBox.Show("Please enter date");
                        workBook.Save();
                        workBook.Close();
                        xlApp.Quit();
                        GC.Collect();
                    }
                }
                else
                {
                    MessageBox.Show("Please fill in all details");
                    workBook.Save();
                    workBook.Close();
                    xlApp.Quit();
                    GC.Collect();
                }
 
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                workBook.Save();
                workBook.Close();
                xlApp.Quit();
                GC.Collect();
            }

        }

        private void dailyTrackGrid_OnLoadingRow(object sender, DataGridRowEventArgs e)
        { 
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }
        //Method for updating a task
        
        //Edit button event handler
        public void EdittTask(object sender, RoutedEventArgs e)
        {
            try
            {
                if (taskGrid.SelectedCells != null)
                {
                    btnEditTask.Visibility = Visibility.Hidden;
                    btnAddTask.Visibility = Visibility.Hidden;
                    addtask.Visibility = Visibility.Hidden;
                    editUpdate.Visibility = Visibility.Visible;
                    upd.Visibility = Visibility.Visible;
                    DataGridCellInfo cell0 = taskGrid.SelectedCells[0];
                    txtTaskId.Text = ((TextBlock)cell0.Column.GetCellContent(cell0.Item)).Text;
                    DataGridCellInfo cell1 = taskGrid.SelectedCells[1];
                    txtTicket.Text = ((TextBlock)cell1.Column.GetCellContent(cell1.Item)).Text;
                    DataGridCellInfo cell2 = taskGrid.SelectedCells[2];
                    txtTaskTitle.Text = ((TextBlock)cell2.Column.GetCellContent(cell2.Item)).Text;
                    DataGridCellInfo cell3 = taskGrid.SelectedCells[3];
                    txtTaskDesc.Text = ((TextBlock)cell3.Column.GetCellContent(cell3.Item)).Text;
                    DataGridCellInfo cell4 = taskGrid.SelectedCells[4];
                    dropdownTaskType.Text = ((TextBlock)cell4.Column.GetCellContent(cell4.Item)).Text;
                    DataGridCellInfo cell5 = taskGrid.SelectedCells[5];
                    dropdownState.Text = ((TextBlock)cell5.Column.GetCellContent(cell5.Item)).Text;
                    DataGridCellInfo cell6 = taskGrid.SelectedCells[6];
                    dropdownPriority.Text = ((TextBlock)cell6.Column.GetCellContent(cell6.Item)).Text;
                    DataGridCellInfo cell8 = taskGrid.SelectedCells[8];
                    txtEfforts.Text = ((TextBlock)cell8.Column.GetCellContent(cell8.Item)).Text;
                    DataGridCellInfo cell9 = taskGrid.SelectedCells[9];
                    txtPSD.Text = ((TextBlock)cell9.Column.GetCellContent(cell9.Item)).Text;
                    DataGridCellInfo cell10 = taskGrid.SelectedCells[10];
                    txtPED.Text = ((TextBlock)cell10.Column.GetCellContent(cell10.Item)).Text;
                    DataGridCellInfo cell11 = taskGrid.SelectedCells[11];
                    txtASD.Text = ((TextBlock)cell11.Column.GetCellContent(cell11.Item)).Text;
                    DataGridCellInfo cell12 = taskGrid.SelectedCells[12];
                    txtAED.Text = ((TextBlock)cell12.Column.GetCellContent(cell12.Item)).Text;
                    txtTaskId.IsReadOnly = true;
                    txtTicket.IsReadOnly = true;
                    txtTaskTitle.IsReadOnly = true;
                    dropdownTaskType.IsEnabled = false;
                    txtTaskDesc.IsReadOnly = true;
                    dropdownPriority.IsEnabled = false;
                    txtPSD.IsReadOnly = true;
                    txtPED.IsReadOnly = true;
                    txtEfforts.IsReadOnly = true;
                }
                else 
                {
                    editUpdate.Visibility = Visibility.Hidden;
                    upd.Visibility = Visibility.Hidden;
                    btnEditTask.Visibility = Visibility.Visible;
                    btnAddTask.Visibility = Visibility.Visible;
                    MessageBox.Show("Please select a row");
                }
            }
            catch
            {
                editUpdate.Visibility = Visibility.Hidden;
                upd.Visibility = Visibility.Hidden;
                btnEditTask.Visibility = Visibility.Visible;
                btnAddTask.Visibility = Visibility.Visible;
                MessageBox.Show("Please select a row");
            }

        }
        //AddTask button event handler
        public void AddTask(object sender, RoutedEventArgs e)
        {
            List<ModelTask> tasklist;
            tasklist = helpers.GetTaskList(id, isAdmin);
            var taskid = "";
            if (tasklist.Count() != 0)
            {
                 taskid = (from ta in tasklist
                              select ta.TaskId).Last();         
            }
            else
            {
                taskid = "0";
            }
            
            int Tid = Convert.ToInt32(taskid);
            txtTaskId.Text = (Tid+1).ToString();
            txtTaskId.IsReadOnly = true;
            txtTicket.IsReadOnly = false;
            txtTaskTitle.IsReadOnly = false;
            dropdownTaskType.IsEnabled = true;
            txtTaskDesc.IsReadOnly = false;
            dropdownPriority.IsEnabled = true;
            txtPSD.IsReadOnly = false;
            txtPED.IsReadOnly = false;
            txtEfforts.IsReadOnly = false;
            txtASD.IsEnabled = false;
            txtAED.IsEnabled = false;
            addtask.Visibility = Visibility.Visible;
            editUpdate.Visibility = Visibility.Visible;   
            btnEditTask.Visibility = Visibility.Hidden;
            btnAddTask.Visibility = Visibility.Hidden;
            upd.Visibility = Visibility.Hidden;
        }
        //Cancel button event handler
        public void CancelAction(object sender,RoutedEventArgs e)
        {
                txtTicket.Text = "";
                txtTaskId.Text = "";
                txtTaskTitle.Text = "";
                dropdownTaskType.Text = "";
                txtTaskDesc.Text = "";
                dropdownPriority.Text = "";
                txtEfforts.Text = "";
                txtPSD.Text = "";
                txtPED.Text = "";
                txtASD.Text = "";
                txtAED.Text = "";
            
            editUpdate.Visibility = Visibility.Hidden;
            btnAddTask.Visibility = Visibility.Visible;
            btnEditTask.Visibility = Visibility.Visible;
        }
        public void AddDailyyTask(object sender, RoutedEventArgs e)
        {
            daily.Visibility = Visibility.Visible;
            btnAdd.Visibility = Visibility.Hidden;
        }
        public void CanAction(object sender, RoutedEventArgs e)
        {
            daily.Visibility = Visibility.Hidden;
            btnAdd.Visibility = Visibility.Visible;
            txtDate.Text = "";
            dropdownTask.Text = "";
            txtHours.Text = "";
            txtRemarks.Text = "";

        }
        private void dropdownState_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dropdownState.SelectedItem.ToString() != "COMPLETED")
            {
                txtASD.IsEnabled = false;
                txtAED.IsEnabled = false;
            }
            else
            {
                txtASD.IsEnabled = false;
                txtAED.IsEnabled = true;

            }
            if(dropdownState.SelectedItem.ToString()== "IN PROGRESS")
            {
                DateTime today = DateTime.Today;
                txtASD.Text= today.ToString("dd/MM/yyyy");
                txtAED.IsEnabled = false;
            }
            else if(dropdownState.SelectedItem.ToString()!="COMPLETED" && dropdownState.SelectedItem.ToString()=="NEW" || dropdownState.SelectedItem.ToString()=="BLOCKED")
            {
                txtASD.Text = "dd/mm/yyyy";
                txtASD.IsEnabled = false;
                txtAED.IsEnabled = false;
            }
        }

        private void txtDate_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (txtDate.Text.ToString() == "dd/mm/yyyy")
            {
                txtDate.Text = "";
            }
        }

        private void txtDate_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtDate.Text))
            {
                txtDate.Text = "dd/mm/yyyy";
            }

        }

        private void txtPSD_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (txtPSD.Text.ToString() == "dd/mm/yyyy")
            {
                txtPSD.Text = "";
            }
        }

        private void txtPSD_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPSD.Text))
            {
                txtPSD.Text = "dd/mm/yyyy";
            }
        }

        private void txtPED_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (txtPED.Text.ToString() == "dd/mm/yyyy")
            {
                txtPED.Text = "";
            }
        }

        private void txtPED_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPED.Text))
            {
                txtPED.Text = "dd/mm/yyyy";
            }
        }

        private void txtASD_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (txtASD.Text.ToString() == "dd/mm/yyyy")
            {
                txtASD.Text = "";
            }
        }

        private void txtASD_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtASD.Text))
            {
                txtASD.Text = "dd/mm/yyyy";
            }
        }

        private void txtAED_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (txtAED.Text.ToString() == "dd/mm/yyyy")
            {
                txtAED.Text = "";
            }

        }

        private void txtAED_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtAED.Text))
            {
                txtAED.Text = "dd/mm/yyyy";
            }
        }

        private void txtSD_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if(txtSD.Text.ToString()=="dd/mm/yyyy")
            {
                txtSD.Text = "";
            }

        }

        private void txtSD_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if(string.IsNullOrWhiteSpace(txtSD.Text))
            {
                txtSD.Text = "dd/mm/yyyy";
            }
        }

        private void txtED_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if(txtED.Text.ToString()=="dd/mm/yyyy")
            {
                txtED.Text = "";
            }
        }

        private void txtED_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if(string.IsNullOrWhiteSpace(txtED.Text))
            {
                txtED.Text = "dd/mm/yyyy";
            }
        }
    }

}
