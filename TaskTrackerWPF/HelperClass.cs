using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.Windows;

namespace TaskTrackerWPF
{
    public class HelperClass
    {
        string filePath = ConfigurationManager.AppSettings["xlsxPath"];

        public List<UserInfo> BindEmployeeData()
        {
            Excel.Application xlApp = new Excel.Application();
            List<UserInfo> userList = new List<UserInfo>();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[3];
            Excel.Range range = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = range.Rows.Count;
            try
            {
                for (int i = 2; i <= rowCount; i++)
                {
                    UserInfo userInformation = new UserInfo();
                    if (workSheet.Cells[i, "A"] != null && workSheet.Cells[i, "A"].Value != null)
                    {
                        userInformation.EmpName = workSheet.Cells[i, "A"].Value.ToString();
                    }
                    if (workSheet.Cells[i, "B"] != null && workSheet.Cells[i, "B"].Value != null)
                    {
                        userInformation.EmpId = workSheet.Cells[i, "B"].Value.ToString();
                    }
                    if (workSheet.Cells[i, "C"] != null && workSheet.Cells[i, "C"].Value != null)
                    {
                        userInformation.UaId = workSheet.Cells[i, "C"].Value.ToString();
                    }
                    if (workSheet.Cells[i, "D"] != null && workSheet.Cells[i, "D"].Value != null)
                    {
                        userInformation.Password = workSheet.Cells[i, "D"].Value.ToString();
                    }
                    if(workSheet.Cells[i,"E"]!=null && workSheet.Cells[i,"E"].Value!=null)
                    {
                        userInformation.AdminAccess = workSheet.Cells[i, "E"].Value.ToString();
                    }
                    userList.Add(userInformation);
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
        public List<ModelTask> GetTaskList(string userId, bool isAdmin)
        {
            Excel.Application xlApp = new Excel.Application();
            List<ModelTask> taskList = new List<ModelTask>();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[2];
            Excel.Worksheet workSheetEmp = workBook.Worksheets[3];
            Excel.Range range = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = range.Rows.Count;
            Excel.Range raEmp = workSheetEmp.UsedRange.Columns["A", Type.Missing];
            int rowCountemp = raEmp.Rows.Count;
            try
            {
                for (int i = 2; i <= rowCount; i++)
                {
                    ModelTask tasks = new ModelTask();
                    if (workSheet.Cells[i, "A"] != null && workSheet.Cells[i, "A"].Value != null)
                    {
                        tasks.TicketNumber = workSheet.Cells[i, "A"].Value.ToString();
                    }
                    if (workSheet.Cells[i, "B"] != null && workSheet.Cells[i, "B"].Value != null)
                    {
                        tasks.TaskId = workSheet.Cells[i, "B"].Value.ToString();
                    }
                    if (workSheet.Cells[i, "C"] != null && workSheet.Cells[i, "C"].Value != null)
                    {
                        tasks.TaskTitle = workSheet.Cells[i, "C"].Value.ToString();
                    }
                    if (workSheet.Cells[i, "D"] != null && workSheet.Cells[i, "D"].Value != null)
                    {
                        tasks.TaskDescription = workSheet.Cells[i, "D"].Value.ToString();
                    }
                    if (workSheet.Cells[i, "E"] != null && workSheet.Cells[i, "E"].Value != null)
                    {
                        tasks.TaskType = workSheet.Cells[i, "E"].Value.ToString();
                    }
                    if (workSheet.Cells[i, "F"] != null && workSheet.Cells[i, "F"].Value != null)
                    {
                        tasks.State = workSheet.Cells[i, "F"].Value.ToString();
                    }
                    if (workSheet.Cells[i, "G"] != null && workSheet.Cells[i, "G"].Value != null)
                    {
                        tasks.Priority = workSheet.Cells[i, "G"].Value.ToString();
                    }
                    if (workSheet.Cells[i, "H"] != null && workSheet.Cells[i, "H"].Value != null)
                    {
                        for (int j = 2; j <= rowCountemp; j++)
                        {
                            if (workSheetEmp.Cells[j, "B"] != null && workSheetEmp.Cells[j, "B"].Value != null)
                            {
                                if (workSheetEmp.Cells[j, "B"].Value.ToString() == workSheet.Cells[i, "H"].Value.ToString())
                                {
                                    tasks.AssignedTo = workSheetEmp.Cells[j, "A"].Value.ToString();
                                }
                            }
                        }
                    }
                    if (workSheet.Cells[i, "I"] != null && workSheet.Cells[i, "I"].Value != null)
                    {
                        tasks.Efforts = workSheet.Cells[i, "I"].Value.ToString();
                    }
                    if (workSheet.Cells[i, "J"] != null && workSheet.Cells[i, "J"].Value != null)
                    {
                        tasks.PlannedStartDate = workSheet.Cells[i, "J"].Value.ToString();
                    }
                    if (workSheet.Cells[i, "K"] != null && workSheet.Cells[i, "K"].Value != null)
                    {
                        tasks.PlannedEndDate = workSheet.Cells[i, "K"].Value.ToString();
                    }
                    if (workSheet.Cells[i, "L"] != null && workSheet.Cells[i, "L"].Value != null)
                    {
                        tasks.ActualStartDate = workSheet.Cells[i, "L"].Value.ToString();
                    }
                    if (workSheet.Cells[i, "M"] != null && workSheet.Cells[i, "M"].Value != null)
                    {
                        tasks.ActualEndDate = workSheet.Cells[i, "M"].Value.ToString();
                    }
                    taskList.Add(tasks);
                }
                workBook.Close();
                xlApp.Quit();
                GC.Collect();
                return taskList;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return null;

        }
        public List<ModelTaskTracker> GetDailyTaskList()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[1];
            Excel.Range range = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = range.Rows.Count;
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
                    if (workSheet.Cells[i, "C"] != null && workSheet.Cells[i, "C"].Value != null)
                    {
                        track.Date = workSheet.Cells[i, "C"].Value.ToString();
                    }
                    if (workSheet.Cells[i, "D"] != null && workSheet.Cells[i, "D"].Value != null)
                    {
                        track.TaskId = workSheet.Cells[i, "D"].Value.ToString();
                    }
                    if (workSheet.Cells[i, "E"] != null && workSheet.Cells[i, "E"].Value != null)
                    {
                        track.HoursSpent = Convert.ToInt32(workSheet.Cells[i, "E"].Value);
                    }
                    if (workSheet.Cells[i, "F"] != null && workSheet.Cells[i, "F"].Value != null)
                    {
                        track.Remarks = workSheet.Cells[i, "F"].Value.ToString();
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
                MessageBox.Show(ex.Message);
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
        public List<ReportGenerationModel> GenerateReport(List<ModelTaskTracker> track, List<UserInfo> user, List<ModelTask> task, string startDate, string endDate)
        {
            try
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
                Excel.Worksheet workSheet = workBook.Worksheets[1];
                Excel.Worksheet workSheetemp = workBook.Worksheets[3];
                Excel.Range range = workSheet.UsedRange.Columns["A", Type.Missing];
                Excel.Range rangeTwo = workSheetemp.UsedRange.Columns["B", Type.Missing];
                int rowCountemp = rangeTwo.Rows.Count;
                int rowCount = range.Rows.Count;
                List<ReportGenerationModel> report = new List<ReportGenerationModel>();

                var masterList = (from t in track
                                  join ta in task
                                  on t.TaskId equals ta.TaskId
                                  join e in user
                                  on t.EmployeeId equals e.EmpId
                                  where DateTime.Parse(t.Date) >= DateTime.Parse(startDate) && DateTime.Parse(t.Date) <= DateTime.Parse(endDate)
                                  select new { t.EmployeeId, e.EmpName, t.TaskId, ta.TaskType, ta.State, t.Date }).ToList();
                var name = (from m in masterList
                            select m.EmpName).Distinct().ToList();


                foreach (var na in name)
                {
                    ReportGenerationModel mo = new ReportGenerationModel();
                    mo.EmpName = na;
                    var templist = (from p in track
                                    join q in task
                                    on p.TaskId equals q.TaskId
                                    select new { q.TaskId, q.State, q.TaskType, q.AssignedTo }).ToList();
                    var bugCount = (from t in templist
                                    where t.TaskType.Equals("Bug") && t.AssignedTo.Equals(na)
                                    select t.TaskId).Distinct().Count();
                    mo.BugCount = bugCount;
                    var featureCount = (from t in templist
                                        where t.TaskType.Equals("Feature") && t.AssignedTo.Equals(na)
                                        select t.TaskId).Distinct().Count();
                    mo.FeatureCount = featureCount;
                    var dailyTaskCount = (from t in templist
                                          where t.TaskType.Equals("Daily Task") && t.AssignedTo.Equals(na)
                                          select t.TaskId).Distinct().Count();
                    mo.DailyTaskCount = dailyTaskCount;
                    var weeklyTaskCount = (from t in templist
                                           where t.TaskType.Equals("Weekly Task") && t.AssignedTo.Equals(na)
                                           select t.TaskId).Distinct().Count();
                    mo.WeeklyTaskCount = weeklyTaskCount;
                    var monthlyTaskCount = (from t in templist
                                            where t.TaskType.Equals("Monthly Task") && t.AssignedTo.Equals(na)
                                            select t.TaskId).Distinct().Count();
                    mo.MonthlyTaskCount = monthlyTaskCount;
                    var otherCount = (from t in templist
                                      where t.TaskType.Equals("Other") && t.AssignedTo.Equals(na)
                                      select t.TaskId).Distinct().Count();
                    mo.OthersCount = otherCount;
                    var totalCount = bugCount + featureCount + dailyTaskCount + weeklyTaskCount + monthlyTaskCount + otherCount;
                    mo.TotalCount = totalCount;
                    var completeCount = (from t in templist
                                         where t.State.Equals("COMPLETED") && t.AssignedTo.Equals(na)
                                         select t.TaskId).Distinct().Count();
                    mo.TotalCompletedCount = completeCount;

                    report.Add(mo);
                }
                workBook.Close();
                xlApp.Quit();
                GC.Collect();
                return report;
            }
            catch
            {
                MessageBox.Show("Something went wrong");
            }
            return null;
        }
        public void UpdateTask(ModelTask mtask)
        {

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[2];
            Excel.Range range = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = range.Rows.Count;
            try
            {
                for (int i = 2; i <= rowCount; i++)
                {
                    if (workSheet.Cells[i, "B"] != null && workSheet.Cells[i, "B"].Value != null && mtask.TaskId == workSheet.Cells[i, "B"].Value.ToString())
                    {
                        MainWindow window = Application.Current.Windows.OfType<MainWindow>().FirstOrDefault();
                        workSheet.Cells[i, "A"].Value = window.txtTicket.Text.ToString();
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
                }
                workBook.Save();
                workBook.Close();

                xlApp.Quit();
                GC.Collect();

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
    }
}
