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

namespace TaskTrackerWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string filePath = ConfigurationManager.AppSettings["xlsxPath"];
        public static bool isAdmin = false;
        string id = "1";
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

                    //if ((workSheet.Cells[i, "H"] != null && workSheet.Cells[i, "H"].Value != null && workSheet.Cells[i, "H"].Value.ToString() == userId) || (isAdmin == true))
                    //{
                        ModelTask ta = new ModelTask();


                        if (workSheet.Cells[i, "A"] != null && workSheet.Cells[i, "A"].Value != null)
                        {

                            ta.TicketNumber = workSheet.Cells[i, "A"].Value.ToString();
                        }
                        else
                        {
                            ta.TicketNumber = "No Value";
                        }
                        if (workSheet.Cells[i, "B"] != null && workSheet.Cells[i, "B"].Value != null)
                        {
                            ta.TaskId = workSheet.Cells[i, "B"].Value.ToString();
                        }
                        else
                        {
                            ta.TaskId = "No Value";
                        }
                        if (workSheet.Cells[i, "C"] != null && workSheet.Cells[i, "C"].Value != null)
                        {
                            ta.TaskTitle = workSheet.Cells[i, "C"].Value.ToString();
                        }
                        else
                        {
                            ta.TaskTitle = "No Value";
                        }
                        if (workSheet.Cells[i, "D"] != null && workSheet.Cells[i, "D"].Value != null)
                        {
                            ta.TaskDescription = workSheet.Cells[i, "D"].Value.ToString();
                        }
                        else
                        {
                            ta.TaskDescription = "No Value";
                        }
                        if (workSheet.Cells[i, "E"] != null && workSheet.Cells[i, "E"].Value != null)
                        {
                            ta.TaskType = workSheet.Cells[i, "E"].Value.ToString();
                        }
                        else
                        {
                            ta.TaskType = "No Value";
                        }
                        if (workSheet.Cells[i, "F"] != null && workSheet.Cells[i, "F"].Value != null)
                        {
                            ta.State = workSheet.Cells[i, "F"].Value.ToString();
                        }
                        else
                        {
                            ta.State = "No Value";
                        }
                        if (workSheet.Cells[i, "G"] != null && workSheet.Cells[i, "G"].Value != null)
                        {
                            ta.Priority = workSheet.Cells[i, "G"].Value.ToString();
                        }
                        else
                        {
                            ta.Priority = "No Value";
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
                            ta.AssignedTo = "No Value";
                        }
                        if (workSheet.Cells[i, "I"] != null && workSheet.Cells[i, "I"].Value != null)
                        {
                            ta.Efforts = workSheet.Cells[i, "I"].Value.ToString();
                        }
                        else
                        {
                            ta.Efforts = "No Value";
                        }
                        if (workSheet.Cells[i, "J"] != null && workSheet.Cells[i, "J"].Value != null)
                        {
                            ta.PlannedStartDate = workSheet.Cells[i, "J"].Value.ToString();
                        }
                        else
                        {
                            ta.PlannedStartDate = "No Value";
                        }
                        if (workSheet.Cells[i, "K"] != null && workSheet.Cells[i, "K"].Value != null)
                        {
                            ta.PlannedEndDate = workSheet.Cells[i, "K"].Value.ToString();
                        }
                        else
                        {
                            ta.PlannedEndDate = "No Value";
                        }
                        if (workSheet.Cells[i, "L"] != null && workSheet.Cells[i, "L"].Value != null)
                        {
                            ta.ActualStartDate= workSheet.Cells[i, "L"].Value.ToString();
                        }
                        else
                        {
                            ta.ActualStartDate = "No Value";
                        }
                        if (workSheet.Cells[i, "M"] != null && workSheet.Cells[i, "M"].Value != null)
                        {
                            ta.ActualEndDate = workSheet.Cells[i, "M"].Value.ToString();
                        }
                        else
                        {
                            ta.ActualEndDate = "No Value";
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
                if (i > rowCount)
                {
                    Console.WriteLine("Data being inserted at:" + i);
                    workSheet.Cells[i, "A"].Value = txtTicket.Text.ToString();
                    workSheet.Cells[i, "B"].Value = txtTaskId.Text.ToString();
                    workSheet.Cells[i, "C"].Value = txtTaskTitle.Text.ToString();
                    workSheet.Cells[i, "D"].Value = txtTaskDesc.Text.ToString();
                    workSheet.Cells[i, "E"].Value = txtTaskType.Text.ToString();
                    workSheet.Cells[i, "F"].Value = dropdownState.SelectedItem.ToString();
                    workSheet.Cells[i, "G"].Value = dropdownPriority.SelectedItem.ToString();
                    workSheet.Cells[i, "H"].Value = id;
                    workSheet.Cells[i, "I"].Value = txtEfforts.Text.ToString();
                    workSheet.Cells[i, "J"].Value = CalenderPSD.SelectedDate.ToString();
                    workSheet.Cells[i, "K"].Value = CalenderPED.SelectedDate.ToString();
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
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
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

                    book.SaveAs("C:\\Users\\drksu\\Desktop\\MyReport.xlsx");
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
                    workSheet.Cells[i, "C"].Value = Calender1.SelectedDate.ToString();
                    workSheet.Cells[i, "D"].Value = dropdownTask.SelectedValue.ToString();
                    workSheet.Cells[i, "E"].Value = txtHours.Text.ToString();
                    workSheet.Cells[i, "F"].Value = txtRemarks.Text.ToString();

                }
                else if(i > rowCount && isAdmin == true)
                {
                    workSheet.Cells[i, "B"].Value = dropdownEmp.SelectedValue.ToString();
                    workSheet.Cells[i, "C"].Value = Calender1.SelectedDate.ToString();
                    workSheet.Cells[i, "D"].Value = dropdownTask.SelectedValue.ToString();
                    workSheet.Cells[i, "E"].Value = txtHours.Text.ToString();
                    workSheet.Cells[i, "F"].Value = txtRemarks.Text.ToString();
                }
             
                MessageBox.Show("Details added successfully");
                //txtDate.Text = "";
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
                Console.WriteLine(ex.Message);
            }

        }
        private void dailyTrackGrid_OnLoadingRow(object sender, DataGridRowEventArgs e)
        {
          
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }


     


    }
}
