using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskTrackerWPF
{
    public class ModelTask
    {
        public string TaskId { get; set; }
        public string TicketNumber { get; set; }
        public string TaskTitle { get; set; }
        public string TaskDescription { get; set; }
        public string TaskType { get; set; }
        public string State { get; set; }
        public string Priority { get; set; }
        public string AssignedTo { get; set; }
        public string Efforts { get; set; }
        public string PlannedStartDate { get; set; }
        public string PlannedEndDate { get; set; }
        public string ActualStartDate { get; set; }
        public string ActualEndDate { get; set; }
        public string HoursSpent { get; set; }
        public string HoursRemaining { get; set; }
        public string ExtraHours { get; set; }
    }
}
