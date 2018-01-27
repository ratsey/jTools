using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JTools.MLO
{
    public class MLOTaskTimingAndReminder : MLOTaskProperty
    {
        public bool InheritParentDates { get; set; }
        public bool HasStartDate { get; set; }
        public DateTime StartDate { get; set; }
        public bool HasDueDate { get; set; }
        public DateTime DueDateTime { get; set; }
        public LeadTime LeadTime { get; set; }
        public bool UseTime { get; set; }
        public int IsLockPeriod { get; set; }
    }
}