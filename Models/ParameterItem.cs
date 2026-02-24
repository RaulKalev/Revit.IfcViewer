using System;

namespace ProSchedules.Models
{
    public class ParameterItem
    {
        public string Name { get; set; }
        public object Id { get; set; } // Can be SchedulableField (Available) or ScheduleFieldId (Scheduled)
        public bool IsScheduled { get; set; }
        public bool IsFieldId { get; set; } // True if Id is ScheduleFieldId, False if SchedulableField
    }
}
