using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace PDF_Column_Report_Generator
{
    public interface IReport
    {
        DataTable ReportData { get; set; } 
        string ReportTitle { get; set; }
    }
}
