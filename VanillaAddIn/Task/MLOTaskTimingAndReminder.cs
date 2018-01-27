using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;

namespace JTools.MLO
{
    public class MLOTaskTimingAndReminder : MLOTaskProperty
    {
        public bool InheritParentDates { get; set; }
        public bool HasStartDate { get; set; }
        public DateTime StartDate { get; set; }
        public bool HasDueDate { get; set; }
        public DateTime DueDateTime { get; set; }

        [JsonConverter(typeof(StringEnumConverter))]
        public LeadTime LeadTime { get; set; }

        public bool UseTime { get; set; }
        public int IsLockPeriod { get; set; }
    }
}