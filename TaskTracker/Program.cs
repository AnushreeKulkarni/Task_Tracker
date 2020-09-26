using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace TaskTracker
{
   public class Program
    {
        public static bool isAdmin = false;
        static void Main(string[] args)
        {
            BusinessLogic logic = new BusinessLogic();



            Console.WriteLine("Enter UserID");
            string id = Console.ReadLine();
            logic.GetTaskList(id, isAdmin);            //GetTaskListMethod



            //ModelTask model = new ModelTask();
            //model.TaskId = "T123459";
            //model.TaskTitle = "TEPTASK";
            //model.TaskDescription = "ABCXYZ";
            //model.TaskType = "BUG";
            //model.State = "BLOCKED";
            //model.State = "LOW";
            //model.AssignedTo = "1";
            //model.Efforts = "26";
            //model.PlannedStartDate = "4/16/2020";
            //model.PlannedEndDate = "4/21/2020";
            //model.AcualStartDate = "4/19/2020";
            //model.ActualEndDate = "4/26/2020";
            //model.HoursSpent = "20";
            //model.HoursRemaining = "6";
            //model.ExtraHours = "0";
            //logic.UpdateTask(model);                //UpdateTaskMethod

            /* logic.AddTask(model);*/                  //AddTaskMethod

            //logic.DeleteTask(model);        //DeleteTaskMethod


            logic.GetEmployeeList();       //GetEmployeeListMethod

            //UserInfo user = new UserInfo();
            //user.EmpName = "Anushree K";
            //user.EmpId = "806782";
            //user.UaId = "UA60105";
            //user.Password = "qwerty";
            //logic.AddEmployee(user);            //AddEmployeeMethod

            /*logic.UpdateEmployee(user); */            //UpdateEmployeeMethod

            //logic.DeleteEmployee(user);          //DeleteEmployeeMethod

            /* logic.GetDailyTaskList(id, isAdmin);*/    //GetDailyTaskListMethod //Not of any use here

            //ModelTaskTracker mo = new ModelTaskTracker();
            //mo.EmployeeId = "806782";
            //mo.Date = "08/03/2020";
            //mo.TaskId = "45";
            //mo.HoursSpent = 12;
            //mo.Remarks = "Some Remarks";
            //Console.WriteLine("Starting");
            //logic.AddDailyTask(mo);                          //AddDailyTaskMethod
            //Console.WriteLine("Finished");

            List<UserInfo> list2;
            list2 = logic.GetEmployeeList();
            List<ModelTaskTracker> list1;
            list1 = logic.GetDailyTaskList(id, isAdmin);
            List<ModelTask> list3;
            list3 = logic.GetTaskList(id, isAdmin);
            logic.DailyTaskList(list1, list2, list3, isAdmin);


        }
    }
}
