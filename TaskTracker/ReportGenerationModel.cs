using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskTracker
{
    public class ReportGenerationModel
    {
        public string EmpName { get; set; }
        public int BugCount { get; set; }
        public int FeatureCount { get; set; }
        public int DailyTaskCount { get; set; }
        public int WeeklyTaskCount { get; set; }
        public int MonthlyTaskCount { get; set; }
        public int OthersCount { get; set; }
        public int TotalCount { get; set; }
        public int TotalCompletedCount { get; set; }
    }
}
