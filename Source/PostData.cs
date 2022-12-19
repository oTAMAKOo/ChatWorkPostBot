
using System;

namespace ChatWorkPostBot
{
    [Serializable]
    public sealed class PostData
    {
        public string hash;

        public string roomId;
            
        public string message;

        public int dayInterval;

        public DayOfWeek[] dayOfWeeks;
            
        public TimeSpan postTime;

        public bool postHoliday;
    }
}
