using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StatureTool
{
    public class OLV_MappingDetails
    {
        public int index { get; set; }
        public string mrn { get; set; }
        public string course { get; set; }
        public string plan { get; set; }
        //DateTime planDate { get; set; }
        public string structureName { get; set; }
        public double volume { get; set; }
        public double d95 { get; set; }
        public double dMean { get; set; }
        public double dMedian { get; set; }
        public string mappingStatus { get; set; }
        public string templateName { get; set; }
        public int tnMatchCount { get; set; }
        public string group { get; set; }

        //public OLV_MappingDetails(string mrn, string course, string plan, DateTime planDate, string structureName, double volume, double d95, string mappingStatus, string templateName, int tnMatchCount, string group)
        public OLV_MappingDetails(int index, string mrn, string course, string plan, string structureName, double volume, double d95, double dMean, double dMedian, string mappingStatus, string templateName, int tnMatchCount, string group)
        {
            this.index = index;
            this.mrn = mrn;
            this.course = course;
            this.plan = plan;
            //this.planDate = planDate;
            this.structureName = structureName;
            this.volume = volume;
            this.d95 = d95;
            this.dMean = dMean;
            this.dMedian = dMedian;
            this.mappingStatus = mappingStatus;
            this.templateName = templateName;
            this.tnMatchCount = tnMatchCount;
            this.group = group;
        }
    }
}
