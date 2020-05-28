using System;
using System.Data;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;

namespace PDF_Column_Report_Generator
{
    partial class ReportEngineer
    {


        // Document Generator
        private void DocumentGenerator(string Path, string documentname)
        {
            // Create a standard US report, Letter size, landscape
            MainDocument = new Document(PageSize.LETTER.Rotate());
            MainDocument.SetMargins(25f, 25f, 35f, 25f);
            PdfWriter.GetInstance(MainDocument, new FileStream(Path + documentname, FileMode.Create));
            MainDocument.Open();            
        }

        // Header Generator
        private void HeaderGenerator(StringBuilder _HeaderContent, bool _ReportDateVisible)
        {
            // Determines whether or not the header content will contain a date
            if (_ReportDateVisible == true)
            {
                _HeaderContent.AppendLine("");
                _HeaderContent.AppendLine("Date of Report: " + DateTime.Today.ToShortDateString());
            }
            // Adds header content
            
            PdfPTable HeaderContent = new PdfPTable(1);
            HeaderContent.DefaultCell.BorderColor = BaseColor.WHITE;            
            PdfPCell HeaderContentText = new PdfPCell(new Paragraph(_HeaderContent.ToString(), CourierLarge));
            HeaderContent.WidthPercentage = 90;
            HeaderContentText.HorizontalAlignment = Element.ALIGN_CENTER;
            HeaderContentText.VerticalAlignment = Element.ALIGN_MIDDLE;
            HeaderContentText.BorderColor = BaseColor.WHITE;
            HeaderContent.AddCell(HeaderContentText);
            HeaderContent.SpacingAfter = 20;
            MainDocument.Add(HeaderContent);        
        }

        private void AddBodyContent(int NoOfColumns, DataTable TableContent)
        {
            // Creates a table to hold the Body of the report
            PdfPTable BodyContent = new PdfPTable(NoOfColumns);       
            // Based on the datatable - creates a header row
            foreach (DataColumn HeaderCol in TableContent.Columns)
            {
                PdfPHeaderCell NewHeaderCell = new PdfPHeaderCell();
                NewHeaderCell.BorderColor = BaseColor.WHITE;
                NewHeaderCell.BackgroundColor = BaseColor.LIGHT_GRAY;
                NewHeaderCell.Phrase = new Phrase(HeaderCol.ColumnName.ToString().ToUpper(), CourierBold);
                BodyContent.AddCell(NewHeaderCell);
            }

            // Creates each row based on data in table
            foreach (DataRow dr in TableContent.Rows)
            {
                for (int i = 0; i <= NoOfColumns - 1; i++)
                {
                    DateTime TempTestDate;
                    if (DateTime.TryParse(dr[i].ToString(), out TempTestDate))
                    {
                        PdfPCell CellContent = new PdfPCell(new Phrase(DateTime.Parse(dr[i].ToString()).ToShortDateString(), Courier));
                        CellContent.BorderColor = BaseColor.WHITE;
                        BodyContent.AddCell(CellContent); 
                    }
                    else
                    {
                        PdfPCell CellContent = new PdfPCell(new Phrase(dr[i].ToString(), Courier));
                        CellContent.BorderColor = BaseColor.WHITE;
                        BodyContent.AddCell(CellContent);
                    }
                }
            }
            // Sets the table width
            BodyContent.WidthPercentage = 100;
            // Sets the header row
            BodyContent.HeaderRows = 1;
            // Sets the spacing after the table
            BodyContent.SpacingAfter = 5;
            // Adds the content to the document
            MainDocument.Add(BodyContent);
            // Adds END OF RECORD statement to the end of the document
            Paragraph EndOfRecord = new Paragraph("*** END OF RECORD ***", CourierBold);
            EndOfRecord.Alignment = Element.ALIGN_CENTER;
            MainDocument.Add(EndOfRecord);  

        }// End of AddContent


    }
}
