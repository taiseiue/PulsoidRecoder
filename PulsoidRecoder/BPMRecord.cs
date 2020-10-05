using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PulsoidRecoder
{
    [Serializable()]
    class BPMRecord
    {
        public int Month { get; set; }

        public int Day { get; set; }

        public int Hour { get; set; }
        public int Minute { get; set; }
        public int Second { get; set; }
        public int BPM { get; set; }

        public BPMRecord(DateTime dateTime,int bpm)
        {
            this.Month = dateTime.Month;
            this.Day = dateTime.Day;
            this.Hour = dateTime.Hour;
            this.Minute = dateTime.Minute;
            this.Second = dateTime.Second;
            this.BPM = bpm;
        }
    }
}
