using System.Data;
using System.Text;

// REPORT ENGINEER V0.20
// UTILIZES ITEXTSHARP V 5.5.1 - PLEASE REFERENCE ITEXTSHARP EULA
// COPYRIGHT JOSHUA H. SANTIAGO, 2015

namespace PDF_Column_Report_Generator
{
    public partial class ReportEngineer
    {
        // Generate the Document, Set it's path, set it's file name
        public void CreateReportDocument(string _SavePath, string _FileName)
        {
            DocumentGenerator(_SavePath, _FileName);        
        }
        // Sets the header for the document
        public void AddHeader(StringBuilder sb, bool GenerateReportDate)
        {
            HeaderGenerator(sb, GenerateReportDate);    
        }
        // Sets the Content for the Document
        public void AddContent(int _NumberOfColumns, DataTable _DataForReport)
        {
            AddBodyContent(_NumberOfColumns, _DataForReport);        
        }
        // Finalizes and saves the document
        public void FinalizeDocument()
        {
            MainDocument.Close();        
        }


    }
}
