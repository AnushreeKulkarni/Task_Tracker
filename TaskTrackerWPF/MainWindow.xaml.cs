using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using TaskTracker;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.IO;
using System.Runtime.Remoting;
using System.Reflection;
using System.Data;
using System.Globalization;

namespace TaskTrackerWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string filePath = ConfigurationManager.AppSettings["xlsxPath"];
        
        public static bool isAdmin = false;
        string id = "2";
        public MainWindow()
        {
            InitializeComponent();
            empGrid.ItemsSource = BindEmployeeData();
            List<UserInfo> list2;
            list2 = BindEmployeeData();
            List<ModelTaskTracker> list1;
            list1 = GetDailyTaskList();
            List<ModelTask> list3;
            list3 = GetTaskList(id, isAdmin);
            DailyTaskList(list1, list2, list3);
            
            Object result;
            result = DailyTaskList(list1, list2, list3);
            if (isAdmin == false)
            {
                empTab.Visibility = Visibility.Hidden;
                reportTab.Visibility = Visibility.Hidden;
                //dailytrackGrid.Width = 650;
                dailytrackGrid.Columns[0].Visibility = Visibility.Visible;
                dailytrackGrid.ItemsSource = (System.Collections.IEnumerable)result;
                taskGrid.Columns[7].Visibility = Visibility.Visible;
                taskGrid.ItemsSource = GetTaskList(id, isAdmin);

            }
            else
            {
                dailytrackGrid.ItemsSource = (System.Collections.IEnumerable)result;
                lblemp.Visibility = Visibility.Visible;
                dropdownEmp.Visibility = Visibility.Visible;
                dropdownEmp.ItemsSource = list2;
                taskGrid.ItemsSource = GetTaskList(id, isAdmin);
                empGrid.ItemsSource = BindEmployeeData();

            }
            dropdownTask.ItemsSource = list3.Distinct();
            dropdownState.Items.Add("NEW");
            dropdownState.Items.Add("IN PROGRESS");
            dropdownState.Items.Add("COMPLETED");
            dropdownState.Items.Add("BLOCKED");
            dropdownPriority.Items.Add("LOW");
            dropdownPriority.Items.Add("MEDIUM");
            dropdownPriority.Items.Add("HIGH");
            combo.Items.Add("ALL");
            //GenerateReport(list1, list2,list3);
        }

        public List<UserInfo> BindEmployeeData()
        {
            Excel.Application xlApp = new Excel.Application();
            List<UserInfo> userList = new List<UserInfo>();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[7];
            Excel.Range range = workSheet.UsedRange;
            Excel.Range ra = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = ra.Rows.Count;
            try
            {
                for (int i = 2; i <=rowCount; i++)
                {
                    UserInfo ua = new UserInfo();
                    if (workSheet.Cells[i, "A"] != null && workSheet.Cells[i, "A"].Value != null)
                    {

                        ua.EmpName = workSheet.Cells[i, "A"].Value.ToString();
                    }
                    else
                    {
                        ua.EmpName = "";
                    }
                    if (workSheet.Cells[i, "B"] != null && workSheet.Cells[i, "B"].Value != null)
                    {
                        ua.EmpId = workSheet.Cells[i, "B"].Value.ToString();
                    }
                    else
                    {
                        ua.EmpId = "";
                    }
                    if (workSheet.Cells[i, "C"] != null && workSheet.Cells[i, "C"].Value != null)
                    {
                        ua.UaId = workSheet.Cells[i, "C"].Value.ToString();
                    }
                    else
                    {
                        ua.UaId = "";
                    }
                    if (workSheet.Cells[i, "D"] != null && workSheet.Cells[i, "D"].Value != null)
                    {
                        ua.Password = workSheet.Cells[i, "D"].Value.ToString();
                    }
                    else
                    {
                        ua.Password = "";
                    }
                    userList.Add(ua);
                }
                workBook.Close();

                xlApp.Quit();
                GC.Collect();
                return userList;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return null;
        }

        private void AddEmployee(object sender, RoutedEventArgs e)
        {      
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[7];
            Excel.Range range = workSheet.UsedRange;
            Excel.Range ra = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = ra.Rows.Count;
           
            int i = rowCount + 1;
            try
            {
                if (i > rowCount)
                {
                    Console.WriteLine("Data being inserted at:" + i);
                    workSheet.Cells[i, "A"].Value = txtName.Text.ToString();
                    workSheet.Cells[i, "B"].Value = txtId.Text.ToString();
                    workSheet.Cells[i, "C"].Value = txtUa.Text.ToString();
                    workSheet.Cells[i, "D"].Value = txtPwd.Text.ToString();
                }
                MessageBox.Show("Details added successfully");
                txtId.Text = "";
                txtName.Text = "";
                txtUa.Text = "";
                txtPwd.Text = "";
                workBook.Save();
                workBook.Close();
                xlApp.Quit();
                GC.Collect();
                empGrid.ItemsSource = BindEmployeeData();


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }
        public List<ModelTask> GetTaskList(string userId, bool isAdmin)
        {
            Excel.Application xlApp = new Excel.Application();
            List<ModelTask> taskList = new List<ModelTask>();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[6];
            Excel.Worksheet workSheetemp = workBook.Worksheets[7];
            Excel.Range range = workSheet.UsedRange;
            Excel.Range ra = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = ra.Rows.Count;
            Excel.Range rangeemp = workSheetemp.UsedRange;
            Excel.Range raemp = workSheetemp.UsedRange.Columns["A", Type.Missing];
            int rowCountemp = raemp.Rows.Count;

            try
            {
                for (int i = 3; i <= rowCount; i++)
                {

                    //if ((workSheet.Cells[i, "H"] != null && workSheet.Cells[i, "H"].Value != null ) || (isAdmin == true))
                    //{
                        ModelTask ta = new ModelTask();


                        if (workSheet.Cells[i, "A"] != null && workSheet.Cells[i, "A"].Value != null)
                        {

                            ta.TicketNumber = workSheet.Cells[i, "A"].Value.ToString();
                        }
                        else
                        {
                           
                        }
                        if (workSheet.Cells[i, "B"] != null && workSheet.Cells[i, "B"].Value != null)
                        {
                            ta.TaskId = workSheet.Cells[i, "B"].Value.ToString();
                        }
                        else
                        {
                            
                        }
                        if (workSheet.Cells[i, "C"] != null && workSheet.Cells[i, "C"].Value != null)
                        {
                            ta.TaskTitle = workSheet.Cells[i, "C"].Value.ToString();
                        }
                        else
                        {
                            
                        }
                        if (workSheet.Cells[i, "D"] != null && workSheet.Cells[i, "D"].Value != null)
                        {
                            ta.TaskDescription = workSheet.Cells[i, "D"].Value.ToString();
                        }
                        else
                        {
                            
                        }
                        if (workSheet.Cells[i, "E"] != null && workSheet.Cells[i, "E"].Value != null)
                        {
                            ta.TaskType = workSheet.Cells[i, "E"].Value.ToString();
                        }
                        else
                        {
                            
                        }
                        if (workSheet.Cells[i, "F"] != null && workSheet.Cells[i, "F"].Value != null)
                        {
                            ta.State = workSheet.Cells[i, "F"].Value.ToString();
                        }
                        else
                        {
                            
                        }
                        if (workSheet.Cells[i, "G"] != null && workSheet.Cells[i, "G"].Value != null)
                        {
                            ta.Priority = workSheet.Cells[i, "G"].Value.ToString();
                        }
                        else
                        {
                            
                        }
                        if (workSheet.Cells[i, "H"] != null && workSheet.Cells[i, "H"].Value != null)
                        {
                            for (int j = 3; j <= rowCountemp; j++)
                            {
                                if (workSheetemp.Cells[j, "B"] != null && workSheetemp.Cells[j, "B"].Value != null)
                                {
                                    if (workSheetemp.Cells[j, "B"].Value.ToString() == workSheet.Cells[i, "H"].Value.ToString())
                                    {

                                        ta.AssignedTo = workSheetemp.Cells[j, "A"].Value.ToString();

                                    }
                                    else
                                    {
                                        
                                    }
                                }
                                else
                                {
                                  
                                }
                            }
                            
                            
                        }
                        else
                        {
                           
                        }
                        if (workSheet.Cells[i, "I"] != null && workSheet.Cells[i, "I"].Value != null)
                        {
                            ta.Efforts = workSheet.Cells[i, "I"].Value.ToString();
                        }
                        else
                        {
                            
                        }
                        if (workSheet.Cells[i, "J"] != null && workSheet.Cells[i, "J"].Value != null)
                        {
                        //var date = workSheet.Cells[i, "J"].Value.ToString();
                        ta.PlannedStartDate = workSheet.Cells[i, "J"].Value.ToString();
                        //ta.PlannedStartDate = DateTime.ParseExact(date, "dd-MM-yyyy HH:mm:ss", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy");
                        
                        }
                        else
                        {
                           
                        }
                        if (workSheet.Cells[i, "K"] != null && workSheet.Cells[i, "K"].Value != null)
                        {
                            ta.PlannedEndDate = workSheet.Cells[i, "K"].Value.ToString();
                        }
                        else
                        {
                            
                        }
                        if (workSheet.Cells[i, "L"] != null && workSheet.Cells[i, "L"].Value != null)
                        {
                            ta.ActualStartDate= workSheet.Cells[i, "L"].Value.ToString();
                        }
                        else
                        {
                            
                        }
                        if (workSheet.Cells[i, "M"] != null && workSheet.Cells[i, "M"].Value != null)
                        {
                            ta.ActualEndDate = workSheet.Cells[i, "M"].Value.ToString();
                        }
                        else
                        {
                           
                        }
                       
                        taskList.Add(ta);
                    }
               // }

                workBook.Close();
                xlApp.Quit();
                GC.Collect();
                return taskList;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return null;

        }
        public void AddTaskData(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[6];
            Excel.Range range = workSheet.UsedRange;
            Excel.Range ra = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = ra.Rows.Count;
            int i = rowCount + 1;
            try
            {
                if (txtTicket.Text != null && txtTaskId.Text != null && txtTaskTitle.Text != null && txtTaskDesc.Text != null && txtTaskType.Text != null && dropdownState.SelectedItem != null && dropdownPriority.SelectedItem != null && txtEfforts.Text != null && txtPSD.Text != null && txtPED.Text != null)
                {
                    if (i > rowCount)
                    {

                        workSheet.Cells[i, "A"].Value = txtTicket.Text.ToString();
                        workSheet.Cells[i, "B"].Value = txtTaskId.Text.ToString();
                        workSheet.Cells[i, "C"].Value = txtTaskTitle.Text.ToString();
                        workSheet.Cells[i, "D"].Value = txtTaskDesc.Text.ToString();
                        workSheet.Cells[i, "E"].Value = txtTaskType.Text.ToString();
                        workSheet.Cells[i, "F"].Value = dropdownState.SelectedItem.ToString();
                        workSheet.Cells[i, "G"].Value = dropdownPriority.SelectedItem.ToString();
                        workSheet.Cells[i, "H"].Value = id;
                        workSheet.Cells[i, "I"].Value = txtEfforts.Text.ToString();
                        workSheet.Cells[i, "j"].Value = txtPSD.Text.ToString();      // DateTime.ParseExact(txtPSD.Text.ToString(),"dd/MM/yyyy",CultureInfo.InvariantCulture).ToString("dd/MM/yyyy");
                        workSheet.Cells[i, "K"].Value = txtPED.Text.ToString();
                        workSheet.Cells[i, "L"].Value = txtASD.Text.ToString();
                        workSheet.Cells[i, "M"].Value = txtAED.Text.ToString();


                    }


                    MessageBox.Show("Details Added Successfully");
                    txtTicket.Text = "";
                    txtTaskId.Text = "";
                    txtTaskTitle.Text = "";
                    txtTaskDesc.Text = "";
                    txtTaskType.Text = "";
                    txtEfforts.Text = "";
                    txtPSD.Text = "";
                    txtPED.Text = "";
                    txtASD.Text = "";
                    txtAED.Text = "";
                    //txtHoursSpent.Text = "";
                    //txtHoursRemaining.Text = "";
                    //txtExtraHours.Text = "";

                    workBook.Save();


                    workBook.Close();

                    xlApp.Quit();

                    GC.Collect();

                    List<UserInfo> list2;
                    list2 = BindEmployeeData();
                    List<ModelTaskTracker> list1;
                    list1 = GetDailyTaskList();
                    List<ModelTask> list3;
                    list3 = GetTaskList(id, isAdmin);
                    DailyTaskList(list1, list2, list3);

                    Object result;
                    result = DailyTaskList(list1, list2, list3);
                    if (isAdmin == false)
                    {
                        //empTab.Visibility = Visibility.Hidden;
                        //dailytrackGrid.Columns[0].Visibility = Visibility.Hidden;
                        dailytrackGrid.ItemsSource = (System.Collections.IEnumerable)result;
                        //taskGrid.Columns[6].Visibility = Visibility.Hidden;
                        taskGrid.ItemsSource = GetTaskList(id, isAdmin);

                    }
                    else
                    {
                        dailytrackGrid.ItemsSource = (System.Collections.IEnumerable)result;
                        taskGrid.ItemsSource = GetTaskList(id, isAdmin);
                        empGrid.ItemsSource = BindEmployeeData();

                    }
                    dropdownTask.ItemsSource = list3;
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
        public void UpdateTaskData(object sender, RoutedEventArgs e)
        {
            if (txtTicket.Text != null && txtTaskId.Text != null && txtTaskTitle.Text != null && txtTaskDesc.Text != null && txtTaskType.Text != null && dropdownState.SelectedItem != null && dropdownPriority.SelectedItem != null && txtEfforts.Text != null && txtPSD.Text != null && txtPED.Text != null)
            {
                ModelTask model = new ModelTask();
                model.TicketNumber = txtTicket.Text.ToString();
                model.TaskId = txtTaskId.Text.ToString();
                model.TaskTitle = txtTaskTitle.Text.ToString();
                model.TaskDescription = txtTaskDesc.Text.ToString();
                model.TaskType = txtTaskType.Text.ToString();
                model.State = dropdownState.SelectedItem.ToString();
                model.Priority = dropdownPriority.SelectedItem.ToString();
                model.AssignedTo = id;
                model.Efforts = txtEfforts.Text.ToString();
                model.PlannedStartDate = txtPSD.Text.ToString();
                model.PlannedEndDate = txtPED.Text.ToString();
                model.ActualStartDate = txtASD.Text.ToString();
                model.ActualEndDate = txtAED.Text.ToString();

                UpdateTask(model);

                MessageBox.Show("Details Updated Successfully");
                txtTicket.Text = "";
                txtTaskId.Text = "";
                txtTaskTitle.Text = "";
                txtTaskDesc.Text = "";
                txtTaskType.Text = "";
                txtEfforts.Text = "";
                txtPSD.Text = "";
                txtPED.Text = "";
                txtASD.Text = "";
                txtAED.Text = "";
                List<UserInfo> list2;
                list2 = BindEmployeeData();
                List<ModelTaskTracker> list1;
                list1 = GetDailyTaskList();
                List<ModelTask> list3;
                list3 = GetTaskList(id, isAdmin);
                DailyTaskList(list1, list2, list3);

                Object result;
                result = DailyTaskList(list1, list2, list3);
                if (isAdmin == false)
                {
                    //empTab.Visibility = Visibility.Hidden;
                    //dailytrackGrid.Columns[0].Visibility = Visibility.Hidden;
                    dailytrackGrid.ItemsSource = (System.Collections.IEnumerable)result;
                    //taskGrid.Columns[6].Visibility = Visibility.Hidden;
                    taskGrid.ItemsSource = GetTaskList(id, isAdmin);

                }
                else
                {
                    dailytrackGrid.ItemsSource = (System.Collections.IEnumerable)result;
                    taskGrid.ItemsSource = GetTaskList(id, isAdmin);
                    empGrid.ItemsSource = BindEmployeeData();

                }
                dropdownTask.ItemsSource = list3;
            }
            else
            {
                MessageBox.Show("Fill in all the details");
            }



        }

        public List<ModelTaskTracker> GetDailyTaskList()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[1];
            Excel.Range range = workSheet.UsedRange;
            Excel.Range ra = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = ra.Rows.Count;
            List<ModelTaskTracker> trackerList = new List<ModelTaskTracker>();
            try
            {
                for (int i = 2; i <= rowCount; i++)
                {
                        ModelTaskTracker track = new ModelTaskTracker();
                        if (workSheet.Cells[i, "B"] != null && workSheet.Cells[i, "B"].Value != null)
                        {
                            track.EmployeeId = workSheet.Cells[i, "B"].Value.ToString();

                        }
                        else
                        {
                            track.EmployeeId = "No Value";
                        }
                        if (workSheet.Cells[i, "C"] != null && workSheet.Cells[i, "C"].Value != null)
                        {
                            track.Date = workSheet.Cells[i, "C"].Value.ToString();
                        }
                        else
                        {
                            track.Date = "No Value";
                        }
                        if (workSheet.Cells[i, "D"] != null && workSheet.Cells[i, "D"].Value != null)
                        {
                            track.TaskId = workSheet.Cells[i, "D"].Value.ToString();
                        }
                        else
                        {
                            track.TaskId = "No Value";
                        }
                        if (workSheet.Cells[i, "E"] != null && workSheet.Cells[i, "E"].Value != null)
                        {
                            track.HoursSpent = Convert.ToInt32(workSheet.Cells[i, "E"].Value);
                        }
                        else
                        {
                            track.HoursSpent = 0;
                        }
                        if (workSheet.Cells[i, "F"] != null && workSheet.Cells[i, "F"].Value != null)
                        {
                            track.Remarks = workSheet.Cells[i, "F"].Value.ToString();
                        }
                        else
                        {
                            track.Remarks = "No Value";
                        }
                        trackerList.Add(track);


                    }
                workBook.Close();

                xlApp.Quit();
                GC.Collect();



                return trackerList;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return null;
        }
        public Object DailyTaskList(List<ModelTaskTracker> trackerList, List<UserInfo> us, List<ModelTask> tasklist)
        {
            
                var List1 = (from m in trackerList
                             join n in tasklist
                             on m.TaskId equals n.TaskId
                             select new { m.EmployeeId, m.Date, m.TaskId, n.TaskTitle, n.TaskType, n.State, n.Priority, m.HoursSpent, m.Remarks }).ToList();
                var List2 = (from m in trackerList
                             join n in us
                             on m.EmployeeId equals n.EmpId
                             select new { m.TaskId, n.EmpId, n.EmpName }).ToList();
                var DailyTaskList2 = (from p in List1
                                      join q in List2
                                      on p.TaskId equals q.TaskId
                                      select new { p.EmployeeId, q.EmpName, p.Date, p.TaskId, p.TaskTitle, p.TaskType, p.State, p.Priority, p.HoursSpent, p.Remarks }).Distinct().ToList();

                return DailyTaskList2;
        }


        public List<ReportGenerationModel> GenerateReport(List<ModelTaskTracker> track,List<UserInfo> user, List<ModelTask> task,string startDate, string endDate)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[1];
            Excel.Worksheet workSheetemp = workBook.Worksheets[7];
            Excel.Range ra = workSheet.UsedRange.Columns["A", Type.Missing];
            Excel.Range raa = workSheetemp.UsedRange.Columns["B", Type.Missing];
            int rowCountemp = raa.Rows.Count;
            int rowCount = ra.Rows.Count;
            List<ReportGenerationModel> report = new List<ReportGenerationModel>();
 
       
            var masterList = (from t in track
                              join ta in task
                              on t.TaskId equals ta.TaskId
                              join e in user
                              on t.EmployeeId equals e.EmpId
                              where DateTime.Parse(t.Date)>=DateTime.Parse(startDate) && DateTime.Parse(t.Date)<=DateTime.Parse(endDate)
                              select new { t.EmployeeId, e.EmpName, t.TaskId, ta.TaskType,ta.State,t.Date }).ToList();
            var name = (from m in masterList
                        select m.EmpName).Distinct().ToList();
           
           
                foreach (var na in name)
                {
                ReportGenerationModel mo = new ReportGenerationModel();
                mo.EmpName = na;
                    var bugCount = (from p in track
                                    join q in task
                                    on p.TaskId equals q.TaskId
                                    where q.TaskType.Equals("Bug") && q.AssignedTo.Equals(na)
                                    select q.TaskId).Distinct().Count(); //use distinct() and check on monday
                    mo.BugCount = bugCount;
                    var featureCount = (from p in track
                                        join q in task
                                        on p.TaskId equals q.TaskId
                                        where q.TaskType.Equals("Feature") && q.AssignedTo.Equals(na)
                                        select q.TaskId).Distinct().Count();
                    mo.FeatureCount = featureCount;

                    var dailyTaskCount = (from p in track
                                          join q in task
                                          on p.TaskId equals q.TaskId
                                          where q.TaskType.Equals("Daily Task") && q.AssignedTo.Equals(na)
                                          select q.TaskId).Distinct().Count();
                    mo.DailyTaskCount = dailyTaskCount;

                    var weeklyTaskCount = (from p in track
                                           join q in task
                                           on p.TaskId equals q.TaskId
                                           where q.TaskType.Equals("Weekly Task") && q.AssignedTo.Equals(na)
                                           select q.TaskId).Distinct().Count();
                    mo.WeeklyTaskCount = weeklyTaskCount;

                    var monthlyTaskCount = (from p in track
                                            join q in task
                                            on p.TaskId equals q.TaskId
                                            where q.TaskType.Equals("Monthly Task") && q.AssignedTo.Equals(na)
                                            select q.TaskId).Distinct().Count();
                    mo.MonthlyTaskCount = monthlyTaskCount;
                    var otherCount = (from p in track
                                      join q in task
                                      on p.TaskId equals q.TaskId
                                      where q.TaskType.Equals("Other") && q.AssignedTo.Equals(na)
                                      select q.TaskId).Distinct().Count();
                    mo.OthersCount = otherCount;
                    var totalCount = bugCount + featureCount + dailyTaskCount + weeklyTaskCount + monthlyTaskCount + otherCount;
                    mo.TotalCount = totalCount;

                    var completedCount = (from p in track
                                          join q in task
                                          on p.TaskId equals q.TaskId
                                          where q.State.Equals("COMPLETED") && q.AssignedTo.Equals(na)
                                          select q.State).Count();
                    mo.TotalCompletedCount = completedCount;
                
                report.Add(mo);
                  
                
                  
                }
                      
            workBook.Close();
            xlApp.Quit();
            GC.Collect();
            return report;

        }
        public void Generate_Report(object sender, RoutedEventArgs e)
        {
            string reportPath = ConfigurationManager.AppSettings["reportPath"]+"MyReport"+DateTime.Now.ToShortDateString()+".xlsx";
            if (combo.SelectedItem.ToString() == "ALL")
            {
                List<UserInfo> list2;
                list2 = BindEmployeeData();
                List<ModelTaskTracker> list1;
                list1 = GetDailyTaskList();
                List<ModelTask> list3;
                list3 = GetTaskList(id, isAdmin);
                List<ReportGenerationModel> report;
                string startDate = CalendarSD.SelectedDate.ToString();
                string endDate = CalendarED.SelectedDate.ToString();
                report = GenerateReport(list1, list2, list3,startDate,endDate);
            


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

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
           else if( combo.SelectedItem.ToString()==null)
            {
                MessageBox.Show("Select Option");
            }
      

        }

        public void AddDailyTask(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[1];
            Excel.Range range = workSheet.UsedRange;
            Excel.Range ra = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = ra.Rows.Count;
            int i = rowCount + 1;
          

            try
            {
                if (i > rowCount && isAdmin==false)
                {
        
                    workSheet.Cells[i, "B"].Value = id;
                    workSheet.Cells[i, "C"].Value = txtDate.Text.ToString();
                    workSheet.Cells[i, "D"].Value = dropdownTask.SelectedValue.ToString();
                    workSheet.Cells[i, "E"].Value = txtHours.Text.ToString();
                    workSheet.Cells[i, "F"].Value = txtRemarks.Text.ToString();

                }
                else if(i > rowCount && isAdmin == true)
                {
                    workSheet.Cells[i, "B"].Value = dropdownEmp.SelectedValue.ToString();
                    workSheet.Cells[i, "C"].Value = txtDate.Text.ToString();
                    workSheet.Cells[i, "D"].Value = dropdownTask.SelectedValue.ToString();
                    workSheet.Cells[i, "E"].Value = txtHours.Text.ToString();
                    workSheet.Cells[i, "F"].Value = txtRemarks.Text.ToString();
                }
             
                MessageBox.Show("Details added successfully");
                txtDate.Text = "";
                txtHours.Text = "";
                txtRemarks.Text = "";
                workBook.Save();
                workBook.Close();
                xlApp.Quit();
               
                GC.Collect();
     
                List<UserInfo> list2;
                list2 = BindEmployeeData();
                List<ModelTaskTracker> list1;
                list1 = GetDailyTaskList();
                List<ModelTask> list3;
                list3 = GetTaskList(id, isAdmin);
                DailyTaskList(list1, list2, list3);
                Object result;
                result = DailyTaskList(list1, list2, list3);
                if (isAdmin == false)
                {
                    
                    dailytrackGrid.ItemsSource = (System.Collections.IEnumerable)result;
                }
                else
                {
                    dailytrackGrid.ItemsSource = (System.Collections.IEnumerable)result;

                }
                dropdownTask.ItemsSource = list3;


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void dailyTrackGrid_OnLoadingRow(object sender, DataGridRowEventArgs e)
        {
          
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

 

        public void UpdateTask(ModelTask mtask)
        {

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[6];
            Excel.Range range = workSheet.UsedRange;
            Excel.Range ra = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = ra.Rows.Count;
            try { 
            for (int i = 3; i <= rowCount; i++)
            {

                if (workSheet.Cells[i, "B"] != null && workSheet.Cells[i, "B"].Value != null && mtask.TaskId == workSheet.Cells[i, "B"].Value.ToString())
                {
                    Console.WriteLine("Id found");
                    workSheet.Cells[i, "A"].Value = txtTicket.Text;
                    workSheet.Cells[i, "B"].Value = mtask.TaskId;
                    workSheet.Cells[i, "C"].Value = mtask.TaskTitle;
                    workSheet.Cells[i, "D"].Value = mtask.TaskDescription;
                    workSheet.Cells[i, "E"].Value = mtask.TaskType;
                    workSheet.Cells[i, "F"].Value = mtask.State;
                    workSheet.Cells[i, "G"].Value = mtask.Priority;
                    workSheet.Cells[i, "H"].Value = mtask.AssignedTo;
                    workSheet.Cells[i, "I"].Value = mtask.Efforts;
                    workSheet.Cells[i, "J"].Value = mtask.PlannedStartDate;
                    workSheet.Cells[i, "K"].Value = mtask.PlannedEndDate;
                    workSheet.Cells[i, "L"].Value = mtask.ActualStartDate;
                    workSheet.Cells[i, "M"].Value = mtask.ActualEndDate;



                }
                else
                {
          

                }

            }
            workBook.Save();


            workBook.Close();

            xlApp.Quit();
            GC.Collect();

        }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        public void EdittTask(object sender, RoutedEventArgs e)
        {
            editUpdate.Visibility = Visibility.Visible;
            btnEditTask.Visibility = Visibility.Hidden;
            btnAddTask.Visibility = Visibility.Hidden;
            addtask.Visibility = Visibility.Hidden;
            Edit.Visibility = Visibility.Visible;
            upd.Visibility = Visibility.Visible;

        }
        public void AddTask(object sender, RoutedEventArgs e)
        {
            addtask.Visibility = Visibility.Visible;
            editUpdate.Visibility = Visibility.Visible;
            btnEditTask.Visibility = Visibility.Hidden;
            Edit.Visibility = Visibility.Hidden;
            btnAddTask.Visibility = Visibility.Hidden;
            upd.Visibility = Visibility.Hidden;
        }
        public void CancelAction(object sender,RoutedEventArgs e)
        {
            editUpdate.Visibility = Visibility.Hidden;
            btnAddTask.Visibility = Visibility.Visible;
            btnEditTask.Visibility = Visibility.Visible;
            txtTicket.Text = "";
            txtTaskId.Text = "";
            txtTaskTitle.Text = "";
            txtTaskType.Text = "";
            txtTaskDesc.Text = "";
            dropdownState.Text = "";
            dropdownPriority.Text = "";
            txtEfforts.Text = "";
            txtPSD.Text = "";
            txtPED.Text = "";
            txtASD.Text = "";
            txtAED.Text = "";



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
        public void EditTask(object sender, RoutedEventArgs e)
        {
            try
            {
                if (taskGrid.SelectedCells != null)
                {
                    DataGridCellInfo cell0 = taskGrid.SelectedCells[0];
                    txtTaskId.Text = ((TextBlock)cell0.Column.GetCellContent(cell0.Item)).Text;
                    DataGridCellInfo cell1 = taskGrid.SelectedCells[1];
                    txtTicket.Text = ((TextBlock)cell1.Column.GetCellContent(cell1.Item)).Text;
                    DataGridCellInfo cell2 = taskGrid.SelectedCells[2];
                    txtTaskTitle.Text = ((TextBlock)cell2.Column.GetCellContent(cell2.Item)).Text;
                    DataGridCellInfo cell3 = taskGrid.SelectedCells[3];
                    txtTaskDesc.Text = ((TextBlock)cell3.Column.GetCellContent(cell3.Item)).Text;
                    DataGridCellInfo cell4 = taskGrid.SelectedCells[4];
                    txtTaskType.Text = ((TextBlock)cell4.Column.GetCellContent(cell4.Item)).Text;
                    DataGridCellInfo cell5 = taskGrid.SelectedCells[5];
                    dropdownState.Text = ((TextBlock)cell5.Column.GetCellContent(cell5.Item)).Text;
                    DataGridCellInfo cell6 = taskGrid.SelectedCells[6];
                    dropdownPriority.Text = ((TextBlock)cell6.Column.GetCellContent(cell6.Item)).Text;     
                    DataGridCellInfo cell8 = taskGrid.SelectedCells[8];
                    txtEfforts.Text = ((TextBlock)cell8.Column.GetCellContent(cell8.Item)).Text;
                    DataGridCellInfo cell9 = taskGrid.SelectedCells[9];
                    txtPSD.Text= ((TextBlock)cell9.Column.GetCellContent(cell9.Item)).Text;
                    DataGridCellInfo cell10 = taskGrid.SelectedCells[10];
                    txtPED.Text = ((TextBlock)cell10.Column.GetCellContent(cell10.Item)).Text;
                    DataGridCellInfo cell11 = taskGrid.SelectedCells[11];
                    txtASD.Text = ((TextBlock)cell11.Column.GetCellContent(cell11.Item)).Text;
                    DataGridCellInfo cell12 = taskGrid.SelectedCells[12];
                    txtAED.Text = ((TextBlock)cell12.Column.GetCellContent(cell12.Item)).Text;
                }
                else
                {
                    MessageBox.Show("Please select a row");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Please select a row");

            }
           
        }


    }
}
