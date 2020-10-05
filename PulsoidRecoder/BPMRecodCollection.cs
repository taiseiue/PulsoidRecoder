using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PulsoidRecoder
{
    [Serializable()]
    class BPMRecordCollection
    {
        public BPMRecord[] Records;

        public BPMRecordCollection(BPMRecord[] recods,int max,int min,int avg)
        {
            this.Records = recods;
            this.Maxbpm = max;
            this.Minbpm = min;
            this.Avgbpm=avg;
        }

        public int Maxbpm { get; set; }
        public int Minbpm { get; set; }
        public int Avgbpm { get; set; }
    }
}
