using System;
using System.Data;

namespace ReportGenerator
{
    public interface IReport
    {
        /// <summary>
        /// Path where the output report file will be stored
        /// </summary>
        string OutputPath { get; set; }

        /// <summary>
        /// Name which will be given to the output report
        /// NOTE: Make sure to add '.pdf' as an extension to any file name
        /// </summary>
        string OutputFileName { get; set; }

        /// <summary>
        /// Data that will be used to generate the report
        /// NOTE: Name column headers manually in order to have a human readable output
        /// on report document
        /// </summary>
        DataTable ReportData { get; set; }

        /// <summary>
        /// Text that will appear at the top of the report as a header
        /// </summary>
        string HeaderText { get; set; }

        /// <summary>
        /// Line under the report header that will indicate the report's title
        /// </summary>
        string ReportTitle { get; set; }

        /// <summary>
        /// Boolean that determines whether the report header will contain the report generation date
        /// </summary>
        bool UseReportDate { get; set; }

        /// <summary>
        /// Columns that should be treated specifically as <see cref="DateTime"/>
        /// NOTE: This will represent the column index with the <see cref="ReportData"/> <see cref="DataTable"/>
        /// </summary>
        int[] ColumnsToTreatAsDateTime { get; set; }
    }
}
