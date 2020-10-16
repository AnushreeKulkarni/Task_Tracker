using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskTrackerWPF
{
   public  class ModelTaskTracker
    {
        public string EmployeeId { get; set; }
        public string Date { get; set; }
        public string TaskId { get; set; }
        public int HoursSpent { get; set; }
        public string Remarks { get; set; }
    }
}
