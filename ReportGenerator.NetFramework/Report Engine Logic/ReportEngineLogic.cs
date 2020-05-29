using iText.IO.Font.Constants;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using System;
using System.Data;
using System.Linq;
using System.Text;

namespace ReportGenerator
{
    public class ReportEngineLogic
    {
        #region Private Members 

        // Fonts 
        private PdfFont m_FONT_Courier = PdfFontFactory.CreateFont(StandardFonts.COURIER); 
        private PdfFont m_FONT_CourierBold = PdfFontFactory.CreateFont(StandardFonts.COURIER_BOLD);

        #endregion

        #region Properties 

        /// <summary>
        /// Report object that is the subject of the report to be generated 
        /// </summary>
        public IReport Report { get; private set; }

        /// <summary>
        /// Main document that will be constructed
        /// </summary>
        public Document MainDocument { get; private set; }

        /// <summary>
        /// Main PDF Document that will contain the underlying document
        /// </summary>
        public PdfDocument MainPDFDocument { get; private set; }

        #endregion

        #region Default Constructor

        public ReportEngineLogic(IReport report)
        {
            Report = report;
            if(Report.ColumnsToTreatAsDateTime == null)
            {
                Report.ColumnsToTreatAsDateTime = new int[0]; 
            }
            CreateDocumentManagementComponents(); 
        }

        #endregion

        #region Public methods 

        /// <summary>
        /// Call this method to construct the report document and return a file path
        /// </summary>
        /// <returns></returns>
        public string CreateDocument()
        {
            try
            {
                HeaderGenerator();
                AddBodyContent();
                AddEndOfRecordIndicator(); 
                MainPDFDocument.Close(); 
                return ConstructFileName();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message); 
            }
        }

        #endregion

        #region Private Methods 

        /// <summary>
        /// Creates a <see cref="PdfWriter"/> and a new <see cref="PdfDocument"/> that
        /// will be used to create the report
        /// </summary>
        private void CreateDocumentManagementComponents()
        {
            MainPDFDocument = new PdfDocument(new PdfWriter(ConstructFileName()));
            MainPDFDocument.SetDefaultPageSize(PageSize.LETTER.Rotate()); 
            MainDocument = new Document(MainPDFDocument); 
        }

        /// <summary>
        /// Constructs a file name based on the Report path and output file name components
        /// </summary>
        /// <returns></returns>
        private string ConstructFileName()
        {
            return $"{Report.OutputPath}{Report.OutputFileName}"; 
        }

        /// <summary>
        /// Modifies the report header to add a <see cref="DateTime"/> representing the report's creation date
        /// </summary>
        private void AddDateToHeader()
        {
            if (Report.UseReportDate)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(Report.HeaderText);
                sb.AppendLine();
                sb.AppendLine("Date of Report: " + DateTime.Today.ToShortDateString());
                Report.HeaderText = sb.ToString();
            }
        }

        /// <summary>
        /// Generates the header that will be assigned to our <see cref="MainDocument"/>
        /// </summary>
        private void HeaderGenerator()
        {
            // Adds date to header if the Report indicates that it should be done 
            AddDateToHeader();

            // Create table that will hold header
            Table HeaderContent = new Table(1);
            HeaderContent.SetWidth(UnitValue.CreatePercentValue(95)); 
            HeaderContent.SetBorder(Border.NO_BORDER);
            HeaderContent.SetHorizontalAlignment(HorizontalAlignment.CENTER); 
            HeaderContent.SetMarginBottom(20); 

            // Create cell that will hold header paragraph
            Cell HeaderContentTextCell = new Cell();
            HeaderContentTextCell.SetHorizontalAlignment(HorizontalAlignment.CENTER);
            HeaderContentTextCell.SetBorder(Border.NO_BORDER); 

            // Create paragraph that will hold header text
            Paragraph HeaderContentParagraph = new Paragraph();

            // Set up paragraph
            HeaderContentParagraph.SetBorder(Border.NO_BORDER);
            HeaderContentParagraph.SetTextAlignment(TextAlignment.CENTER); 
            HeaderContentParagraph.Add(Report.HeaderText);
            HeaderContentParagraph.SetHorizontalAlignment(HorizontalAlignment.CENTER);
            HeaderContentParagraph.SetVerticalAlignment(VerticalAlignment.MIDDLE);
            HeaderContentParagraph.SetFont(m_FONT_Courier);
            HeaderContentParagraph.SetFontSize(10);

            // Combine elements and add to main document
            HeaderContentTextCell.Add(HeaderContentParagraph);
            HeaderContent.AddCell(HeaderContentTextCell);
            MainDocument.Add(HeaderContent); 
        }

        /// <summary>
        /// Adds all body content to the report
        /// </summary>
        private void AddBodyContent()
        {
            // Create a table to hold the body of the report
            // NOTE: Set column count to that of the Report DataTable
            Table BodyContent = new Table(Report.ReportData.Columns.Count);

            // Sets the table width
            BodyContent.SetWidth(UnitValue.CreatePercentValue(95));

            // Set table positioning
            BodyContent.SetHorizontalAlignment(HorizontalAlignment.CENTER);

            // Sets the spacing after the table
            BodyContent.SetMarginBottom(5);


            // For each column in the Report Data, add a header cell with the column's name
            foreach (DataColumn column in Report.ReportData.Columns)
            {
                // Create cell
                Cell HeaderCell = new Cell();
                // Set cell background and border 
                HeaderCell.SetBackgroundColor(ColorConstants.LIGHT_GRAY);
                HeaderCell.SetBorder(Border.NO_BORDER); 
                // Create a paragraph to hold the header cell text
                Paragraph HeaderCellText = new Paragraph(column.ColumnName);
                // Set up the text formatting in the header cell paragraph
                HeaderCellText.SetFont(m_FONT_CourierBold);
                HeaderCellText.SetFontSize(8);
                HeaderCellText.SetTextAlignment(TextAlignment.CENTER); 
                // Add paragraph to cell
                HeaderCell.Add(HeaderCellText); 
                // Add cell to Header row in Body Content
                BodyContent.AddHeaderCell(HeaderCell);  
            }

            // Creates each row based on data in table
            foreach (DataRow dr in Report.ReportData.Rows)
            {
                for (int i = 0; i < Report.ReportData.Columns.Count; i++)
                {
                    // Create cell and paragraph to hold text
                    Cell CellContent = new Cell();
                    Paragraph CellContentText = new Paragraph();

                    // Set up cell for proper formatting
                    CellContent.SetBorder(Border.NO_BORDER);
                    CellContent.SetFontSize(8);
                    CellContent.SetHorizontalAlignment(HorizontalAlignment.LEFT);

                    if (Report.ColumnsToTreatAsDateTime.Contains(i))
                    {
                        DateTime outputDateTime;
                        if(DateTime.TryParse(dr[i].ToString(), out outputDateTime))
                        {
                            CellContentText.Add(outputDateTime.ToShortDateString());
                            CellContent.Add(CellContentText);
                            BodyContent.AddCell(CellContent); 
                        }
                        else
                        {
                            CellContentText.Add(dr[i].ToString());
                            CellContent.Add(CellContentText);
                            BodyContent.AddCell(CellContent); 
                        }
                    }
                    else
                    {
                        CellContentText.Add(dr[i].ToString());
                        CellContent.Add(CellContentText);
                        BodyContent.AddCell(CellContent);
                    }
                }
            }

            // Adds the body content to the document
            MainDocument.Add(BodyContent);
        }

        private void AddEndOfRecordIndicator()
        {
            // Adds END OF RECORD statement to the end of the document
            Paragraph EndOfRecord = new Paragraph();
            EndOfRecord.Add("*** END OF RECORD ***");
            EndOfRecord.SetHorizontalAlignment(HorizontalAlignment.CENTER);
            EndOfRecord.SetTextAlignment(TextAlignment.CENTER);
            EndOfRecord.SetFont(m_FONT_CourierBold);
            EndOfRecord.SetFontSize(10);
            MainDocument.Add(EndOfRecord);
        }

        #endregion 
    }
}
