using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel=Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Security.Cryptography.X509Certificates;
using Microsoft.SqlServer.Server;
using System.Configuration;

namespace TaskTracker
{
    public class BusinessLogic
    {
        
        string filePath = ConfigurationManager.AppSettings["xlsxPath"];
   

        public void AddTask(ModelTask m)
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
                    workSheet.Cells[i, "A"].Value = m.TaskId;
                    workSheet.Cells[i, "B"].Value = m.TaskTitle;
                    workSheet.Cells[i, "C"].Value = m.TaskDescription;
                    workSheet.Cells[i, "D"].Value = m.TaskType;
                    workSheet.Cells[i, "E"].Value = m.State;
                    workSheet.Cells[i, "F"].Value = m.Priority;
                    workSheet.Cells[i, "G"].Value = m.AssignedTo;
                    workSheet.Cells[i, "H"].Value = m.Efforts;
                    workSheet.Cells[i, "I"].Value = m.PlannedStartDate;
                    workSheet.Cells[i, "J"].Value = m.PlannedEndDate;
                    workSheet.Cells[i, "K"].Value = m.ActualStartDate;
                    workSheet.Cells[i, "L"].Value = m.ActualEndDate;
                    workSheet.Cells[i, "M"].Value = m.HoursSpent;
                    workSheet.Cells[i, "N"].Value = m.HoursRemaining;
                    workSheet.Cells[i, "O"].Value = m.ExtraHours;
                }

                workBook.Save();

          
                workBook.Close();
              
                xlApp.Quit();

                GC.Collect();
            }
            catch(Exception ex)
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
            Excel.Range range = workSheet.UsedRange;
            Excel.Range ra = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = ra.Rows.Count;
           
            try
            {
                for (int i = 3; i <= rowCount; i++)
                {

                    if ((workSheet.Cells[i, "G"] != null && workSheet.Cells[i, "G"].Value != null && workSheet.Cells[i, "G"].Value.ToString() == userId) || (isAdmin == true))
                    {
                        ModelTask ta = new ModelTask();
                       

                        if (workSheet.Cells[i, "A"] != null && workSheet.Cells[i, "A"].Value != null)
                        {

                            ta.TaskId = workSheet.Cells[i, "A"].Value.ToString();
                        }
                        else
                        {
                            ta.TaskId = "No Value";
                        }
                        if (workSheet.Cells[i, "B"] != null && workSheet.Cells[i, "B"].Value != null)
                        {
                            ta.TaskTitle = workSheet.Cells[i, "B"].Value.ToString();
                        }
                        else
                        {
                            ta.TaskTitle = "No Value";
                        }
                        if (workSheet.Cells[i, "C"] != null && workSheet.Cells[i, "C"].Value != null)
                        {
                            ta.TaskDescription = workSheet.Cells[i, "C"].Value.ToString();
                        }
                        else
                        {
                            ta.TaskDescription = "No Value";
                        }
                        if (workSheet.Cells[i, "D"] != null && workSheet.Cells[i, "D"].Value != null)
                        {
                            ta.TaskType = workSheet.Cells[i, "D"].Value.ToString();
                        }
                        else
                        {
                            ta.TaskType = "No Value";
                        }
                        if (workSheet.Cells[i, "E"] != null && workSheet.Cells[i, "E"].Value != null)
                        {
                            ta.State = workSheet.Cells[i, "E"].Value.ToString();
                        }
                        else
                        {
                            ta.State = "No Value";
                        }
                        if (workSheet.Cells[i, "F"] != null && workSheet.Cells[i, "F"].Value != null)
                        {
                            ta.Priority = workSheet.Cells[i, "F"].Value.ToString();
                        }
                        else
                        {
                            ta.Priority = "No Value";
                        }
                        if (workSheet.Cells[i, "G"] != null && workSheet.Cells[i, "G"].Value != null)
                        {
                            ta.AssignedTo = workSheet.Cells[i, "G"].Value.ToString();
                        }
                        else
                        {
                            ta.AssignedTo = "No Value";
                        }
                        if (workSheet.Cells[i, "H"] != null && workSheet.Cells[i, "H"].Value != null)
                        {
                            ta.Efforts = workSheet.Cells[i, "H"].Value.ToString();
                        }
                        else
                        {
                            ta.Efforts = "No Value";
                        }
                        if (workSheet.Cells[i, "I"] != null && workSheet.Cells[i, "I"].Value != null)
                        {
                            ta.PlannedStartDate = workSheet.Cells[i, "I"].Value.ToString();
                        }
                        else
                        {
                            ta.PlannedStartDate = "No Value";
                        }
                        if (workSheet.Cells[i, "J"] != null && workSheet.Cells[i, "J"].Value != null)
                        {
                            ta.PlannedEndDate = workSheet.Cells[i, "J"].Value.ToString();
                        }
                        else
                        {
                            ta.PlannedEndDate = "No Value";
                        }
                        if (workSheet.Cells[i, "K"] != null && workSheet.Cells[i, "K"].Value != null)
                        {
                            ta.ActualStartDate = workSheet.Cells[i, "K"].Value.ToString();
                        }
                        else
                        {
                            ta.ActualStartDate = "No Value";
                        }
                        if (workSheet.Cells[i, "L"] != null && workSheet.Cells[i, "L"].Value != null)
                        {
                            ta.ActualEndDate = workSheet.Cells[i, "L"].Value.ToString();
                        }
                        else
                        {
                            ta.ActualEndDate = "No Value";
                        }
                        if (workSheet.Cells[i, "M"] != null && workSheet.Cells[i, "M"].Value != null)
                        {
                            ta.HoursSpent = workSheet.Cells[i, "M"].Value.ToString();
                        }
                        else
                        {
                            ta.HoursSpent = "No Value";
                        }
                        if (workSheet.Cells[i, "N"] != null && workSheet.Cells[i, "N"].Value != null)
                        {
                            ta.HoursRemaining = workSheet.Cells[i, "N"].Value.ToString();
                        }
                        else
                        {
                            ta.HoursRemaining = "No Value";
                        }
                        if (workSheet.Cells[i, "O"] != null && workSheet.Cells[i, "O"].Value != null)
                        {
                            ta.ExtraHours = workSheet.Cells[i, "O"].Value.ToString();
                        }
                        else
                        {
                            ta.ExtraHours = "No Value";
                        }
                        taskList.Add(ta);
                    }
                }


                //foreach (var record in taskList)
                //{

                //    Console.WriteLine(Environment.NewLine);
                //    Console.Write(record.TaskId);
                //    Console.Write(record.TaskTitle);
                //    Console.Write(record.TaskDescription);
                //    Console.Write(record.TaskType);
                //    Console.Write(record.State);
                //    Console.Write(record.Priority);
                //    Console.Write(record.AssignedTo);
                //    Console.Write(record.Efforts);
                //    Console.Write(record.PlannedStartDate);
                //    Console.Write(record.PlannedEndDate);
                //    Console.Write(record.AcualStartDate);
                //    Console.Write(record.ActualEndDate);
                //    Console.Write(record.HoursSpent);
                //    Console.Write(record.HoursRemaining);
                //    Console.Write(record.ExtraHours);

                //}
              
         
                workBook.Close();
            
                xlApp.Quit();
                GC.Collect();


                return taskList;
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return null;
                  
        }
        public void UpdateTask(ModelTask mtask )
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[6];
            Excel.Range range = workSheet.UsedRange;
            Excel.Range ra = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = ra.Rows.Count;
            try
            {
                for (int i = 3; i <= rowCount; i++)
                {

                    if (workSheet.Cells[i, "A"] != null && workSheet.Cells[i, "A"].Value != null && mtask.TaskId == workSheet.Cells[i, "A"].Value.ToString())
                    {
                        Console.WriteLine("Id found");
                        workSheet.Cells[i, "A"].Value = mtask.TaskId;
                        workSheet.Cells[i, "B"].Value = mtask.TaskTitle;
                        workSheet.Cells[i, "C"].Value = mtask.TaskDescription;
                        workSheet.Cells[i, "D"].Value = mtask.TaskType;
                        workSheet.Cells[i, "E"].Value = mtask.State;
                        workSheet.Cells[i, "F"].Value = mtask.Priority;
                        workSheet.Cells[i, "G"].Value = mtask.AssignedTo;
                        workSheet.Cells[i, "H"].Value = mtask.Efforts;
                        workSheet.Cells[i, "I"].Value = mtask.PlannedStartDate;
                        workSheet.Cells[i, "J"].Value = mtask.PlannedEndDate;
                        workSheet.Cells[i, "K"].Value = mtask.ActualStartDate;
                        workSheet.Cells[i, "L"].Value = mtask.ActualEndDate;
                        workSheet.Cells[i, "M"].Value = mtask.HoursSpent;
                        workSheet.Cells[i, "N"].Value = mtask.HoursRemaining;
                        workSheet.Cells[i, "O"].Value = mtask.ExtraHours;


                    }
                    else
                    {
                        Console.WriteLine("Id not found");

                    }

                }
                workBook.Save();

          
                workBook.Close();
               
                xlApp.Quit();
                GC.Collect();

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            

        }
        public void DeleteTask(ModelTask t)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[6];
            Excel.Range range = workSheet.UsedRange;
            Excel.Range ra = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = ra.Rows.Count;
            try
            {
                for (int i = 3; i <= rowCount; i++)
                {
                    if (workSheet.Cells[i, "A"] != null && workSheet.Cells[i, "A"].Value != null && t.TaskId == workSheet.Cells[i, "A"].Value.ToString())
                    {
                        Console.WriteLine("Deleting at:" +i);
                        workSheet.Cells[i, "A"].Value = null;
                        workSheet.Cells[i, "B"].Value = null;
                        workSheet.Cells[i, "C"].Value = null;
                        workSheet.Cells[i, "D"].Value = null;
                        workSheet.Cells[i, "E"].Value = null;
                        workSheet.Cells[i, "F"].Value = null;
                        workSheet.Cells[i, "G"].Value = null;
                        workSheet.Cells[i, "H"].Value = null;
                        workSheet.Cells[i, "I"].Value = null;
                        workSheet.Cells[i, "J"].Value = null;
                        workSheet.Cells[i, "K"].Value = null;
                        workSheet.Cells[i, "L"].Value = null;
                        workSheet.Cells[i, "M"].Value = null;
                        workSheet.Cells[i, "N"].Value = null;
                        workSheet.Cells[i, "O"].Value = null;

                    }
                    else
                    {
                        Console.WriteLine("Id not found");

                    }

                }
                workBook.Save();

     
                workBook.Close();
            
                xlApp.Quit();
                GC.Collect();

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
                

        }

        

        public List<UserInfo> GetEmployeeList()
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
                        ua.EmpName = "No Value";
                    }
                    if (workSheet.Cells[i, "B"] != null && workSheet.Cells[i, "B"].Value != null)
                    {
                        ua.EmpId = workSheet.Cells[i, "B"].Value.ToString();
                    }
                    else
                    {
                        ua.EmpId = "No Value";
                    }
                    if (workSheet.Cells[i, "C"] != null && workSheet.Cells[i, "C"].Value != null)
                    {
                        ua.UaId = workSheet.Cells[i, "C"].Value.ToString();
                    }
                    else
                    {
                        ua.UaId = "No Value";
                    }
                    if (workSheet.Cells[i, "D"] != null && workSheet.Cells[i, "D"].Value != null)
                    {
                        ua.Password = workSheet.Cells[i, "D"].Value.ToString();
                    }
                    else
                    {
                        ua.Password = "No Value";
                    }
                    userList.Add(ua);
                }
                //foreach (var user in userList)
                //{
                //    Console.Write(user.EmpName);
                //    Console.Write(user.EmpId);
                //    Console.Write(user.UaId);
                //    Console.Write(user.Password);
                //    Console.WriteLine(Environment.NewLine);
                //}
       
    
                workBook.Close();
            
                xlApp.Quit();
                GC.Collect();


                return userList;
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return null;

        }
        public void AddEmployee(UserInfo user)
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
                    workSheet.Cells[i, "A"].Value = user.EmpName;
                    workSheet.Cells[i, "B"].Value = user.EmpId;
                    workSheet.Cells[i, "C"].Value = user.UaId;
                    workSheet.Cells[i, "D"].Value = user.Password;


                }
                workBook.Save();

            
   
                workBook.Close();
              
                xlApp.Quit();
                GC.Collect();

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        public void UpdateEmployee(UserInfo user)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[7];
            Excel.Range range = workSheet.UsedRange;
            Excel.Range ra = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = ra.Rows.Count;
            try
            {
                for (int i = 2; i <=rowCount; i++)
                {
                    if (workSheet.Cells[i, "B"] != null && workSheet.Cells[i, "B"].Value != null && user.EmpId == workSheet.Cells[i, "B"].Value.ToString())
                    {
                        Console.WriteLine("Id Found");
                        workSheet.Cells[i, "A"].Value = user.EmpName;
                        workSheet.Cells[i, "B"].Value = user.EmpId;
                        workSheet.Cells[i, "C"].Value = user.UaId;
                        workSheet.Cells[i, "D"].Value = user.Password;

                    }
                    else
                    {
                        Console.WriteLine("Id not found");

                    }

                }
                workBook.Save();

          
         
                workBook.Close();
               
                xlApp.Quit();
                GC.Collect();

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        public void DeleteEmployee(UserInfo user)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook workBook = xlApp.Workbooks.Open(filePath);
            Excel.Worksheet workSheet = workBook.Worksheets[7];
            Excel.Range range = workSheet.UsedRange;
            Excel.Range ra = workSheet.UsedRange.Columns["A", Type.Missing];
            int rowCount = ra.Rows.Count;
            try
            {
                for (int i = 2; i <= rowCount; i++)
                {
                    if (workSheet.Cells[i, "B"] != null && workSheet.Cells[i, "B"].Value != null && user.EmpId == workSheet.Cells[i, "B"].Value.ToString())
                    {
                        Console.WriteLine("Deleting record at:" + i);
                        workSheet.Cells[i, "A"].Value = null;
                        workSheet.Cells[i, "B"].Value = null;
                        workSheet.Cells[i, "C"].Value = null;
                        workSheet.Cells[i, "D"].Value = null;

                    }
                    else
                    {
                        Console.WriteLine("Id not found");

                    }

                }
                workBook.Save();

              
          
                workBook.Close();
         
                xlApp.Quit();
                GC.Collect();

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public List<ModelTaskTracker> GetDailyTaskList(string EmpId,bool isAdmin)
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
                    if ((workSheet.Cells[i, "B"] != null && workSheet.Cells[i, "B"].Value != null && workSheet.Cells[i, "B"].Value.ToString() == EmpId) || (isAdmin == true))
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

                }
                //foreach (var track in trackerList)
                //{
                //    Console.WriteLine(Environment.NewLine);
                //    Console.Write(track.EmployeeId);
                //    Console.Write(track.Date);
                //    Console.Write(track.TaskId);
                //    Console.Write(track.HoursSpent);
                //    Console.Write(track.Remarks);

                //}
           
              
                workBook.Close();
           
                xlApp.Quit();
                GC.Collect();



                return trackerList;
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return null;
        }
        public object DailyTaskList(List<ModelTaskTracker> trackerList,List<UserInfo> us,List<ModelTask> tasklist,bool isAdmin)
        {
            if (isAdmin == false)
            {
                var DailyTaskList1 = (from m in trackerList
                                      join n in tasklist
                                      on m.TaskId equals n.TaskId
                                      select new {m.EmployeeId,m.Date, m.TaskId, n.TaskTitle, n.TaskType, n.State, n.Priority, m.HoursSpent, m.Remarks }).ToList();

                //foreach (var d in DailyTaskList1)
                //{
                //    Console.WriteLine(Environment.NewLine);
                //    Console.Write(d.EmployeeId);
                //    Console.Write(d.Date);
                //    Console.Write(d.TaskId);
                //    Console.Write(d.TaskTitle);
                //    Console.Write(d.TaskType);
                //    Console.Write(d.State);
                //    Console.Write(d.Priority);
                //    Console.Write(d.HoursSpent);
                //    Console.Write(d.Remarks);
                //}


                return DailyTaskList1;
            }

            else
            {
                var List1 = (from m in trackerList
                                      join n in tasklist
                                      on m.TaskId equals n.TaskId
                                      select new { m.EmployeeId, m.Date, m.TaskId, n.TaskTitle, n.TaskType, n.State, n.Priority, m.HoursSpent, m.Remarks }).ToList();
                var List2 = (from m in trackerList
                                      join n in us
                                      on m.EmployeeId equals n.EmpId
                                      select new {m.TaskId, n.EmpId,n.EmpName }).ToList();
                var DailyTaskList2 = (from p in List1
                                      join q in List2
                                      on p.TaskId equals q.TaskId
                                      select new { p.EmployeeId, q.EmpName, p.Date, p.TaskId, p.TaskTitle, p.TaskType, p.State, p.Priority, p.HoursSpent, p.Remarks }).ToList();

                //foreach(var d in DailyTaskList2)
                //{
                //    Console.WriteLine(Environment.NewLine);
                //    Console.Write(d.EmployeeId);
                //    Console.Write(d.EmpName);
                //    Console.Write(d.Date);
                //    Console.Write(d.TaskId);
                //    Console.Write(d.TaskTitle);
                //    Console.Write(d.TaskType);
                //    Console.Write(d.State);
                //    Console.Write(d.Priority);
                //    Console.Write(d.HoursSpent);
                //    Console.Write(d.Remarks);

                //}

                return DailyTaskList2;
            }
          

           
                           
        }
        public void AddDailyTask(ModelTaskTracker mo)
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
                if (i > rowCount)
                {
                    
                    Console.WriteLine("Data being inserted at:" + i);
                    workSheet.Cells[i, "B"].Value = mo.EmployeeId;       
                    workSheet.Cells[i, "C"].Value = mo.Date;
                    workSheet.Cells[i, "D"].Value = mo.TaskId;          
                    workSheet.Cells[i, "E"].Value = mo.HoursSpent;
                    workSheet.Cells[i, "F"].Value = mo.Remarks;

                }
                workBook.Save();

                
            
                workBook.Close();
              
                xlApp.Quit();
                GC.Collect();

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


        }
    


        }

        }
    

