using iTextSharp.text;
using iTextSharp.text.html;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ITextHtmlPdfDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            var html = File.ReadAllText("email-template.html");
            PdfToHTML(html);
        }

        public static void PdfToHTML(string html)
        {
            byte[] bytes;

            using (var ms = new MemoryStream())
            {
                using (var doc = new Document(PageSize.A3))
                {
                    doc.SetMargins(50, 50, 260, 50);
                    using (var writer = PdfWriter.GetInstance(doc, ms))
                    {
                        doc.Open();
                        writer.PageEvent = new ITextEvents(new object());

                        iTextSharp.text.Font baseFontBold = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 9f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                        iTextSharp.text.Font baseFontNormal = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 9f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);

                        Paragraph para;
                        PdfPCell pdfCell;

                        PdfPTable pdfTab = new PdfPTable(8);
                        float[] widths = new float[] { 5f, 10f, 10f,10f,10f,20f,10f,15f };
                        pdfTab.SetWidths(widths);

                        pdfTab.TotalWidth = doc.PageSize.Width - doc.LeftMargin - doc.RightMargin;
                        pdfTab.WidthPercentage = 100;
                        pdfTab.HeaderRows = 1;

                        pdfCell = new PdfPCell();
                        pdfCell.UseAscender = true;
                        pdfCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
                        pdfCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                        pdfCell.FixedHeight = 20f;
                        para = new Paragraph("#", baseFontBold);
                        para.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                        pdfCell.AddElement(para);
                        pdfTab.AddCell(pdfCell);

                        pdfCell = new PdfPCell();
                        pdfCell.UseAscender = true;
                        pdfCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
                        pdfCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                        para = new Paragraph("DISTRICT", baseFontBold);
                        para.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                        pdfCell.AddElement(para);
                        pdfTab.AddCell(pdfCell);

                        pdfCell = new PdfPCell();
                        pdfCell.UseAscender = true;
                        pdfCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
                        pdfCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                        para = new Paragraph("TEHSIL", baseFontBold);
                        para.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                        pdfCell.AddElement(para);
                        pdfTab.AddCell(pdfCell);

                        pdfCell = new PdfPCell();
                        pdfCell.UseAscender = true;
                        pdfCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
                        pdfCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                        para = new Paragraph("UC NAME", baseFontBold);
                        para.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                        pdfCell.AddElement(para);
                        pdfTab.AddCell(pdfCell);

                        pdfCell = new PdfPCell();
                        pdfCell.UseAscender = true;
                        pdfCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
                        pdfCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                        para = new Paragraph("SEMIS", baseFontBold);
                        para.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                        pdfCell.AddElement(para);
                        pdfTab.AddCell(pdfCell);

                        pdfCell = new PdfPCell();
                        pdfCell.UseAscender = true;
                        pdfCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
                        pdfCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                        para = new Paragraph("SCHOOL NAME", baseFontBold);
                        para.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                        pdfCell.AddElement(para);
                        pdfTab.AddCell(pdfCell);

                        pdfCell = new PdfPCell();
                        pdfCell.UseAscender = true;
                        pdfCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
                        pdfCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                        para = new Paragraph("TYPE", baseFontBold);
                        para.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                        pdfCell.AddElement(para);
                        pdfTab.AddCell(pdfCell);

                        pdfCell = new PdfPCell();
                        pdfCell.UseAscender = true;
                        pdfCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
                        pdfCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                        para = new Paragraph("ATTENDANCE DATE", baseFontBold);
                        para.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                        pdfCell.AddElement(para);
                        pdfTab.AddCell(pdfCell);

                        for(var i = 0; i < 80; i++)
                        {
                            pdfCell = new PdfPCell();
                            pdfCell.UseAscender = true;
                            pdfCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
                            pdfCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                            pdfCell.FixedHeight = 20f;
                            para = new Paragraph(""+i+10, baseFontNormal);
                            para.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                            pdfCell.AddElement(para);
                            pdfTab.AddCell(pdfCell);

                            pdfCell = new PdfPCell();
                            pdfCell.UseAscender = true;
                            pdfCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
                            pdfCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER;
                            para = new Paragraph("Larkana", baseFontNormal);
                            para.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                            pdfCell.AddElement(para);
                            pdfTab.AddCell(pdfCell);

                            pdfCell = new PdfPCell();
                            pdfCell.UseAscender = true;
                            pdfCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
                            pdfCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
                            para = new Paragraph("Dokri", baseFontNormal);
                            para.Alignment = iTextSharp.text.Element.ALIGN_LEFT;
                            pdfCell.AddElement(para);
                            pdfTab.AddCell(pdfCell);

                            pdfCell = new PdfPCell();
                            pdfCell.UseAscender = true;
                            pdfCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
                            pdfCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
                            para = new Paragraph("2-Badah-II", baseFontNormal);
                            para.Alignment = iTextSharp.text.Element.ALIGN_LEFT;
                            pdfCell.AddElement(para);
                            pdfTab.AddCell(pdfCell);

                            pdfCell = new PdfPCell();
                            pdfCell.UseAscender = true;
                            pdfCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
                            pdfCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
                            para = new Paragraph("413010013", baseFontNormal);
                            para.Alignment = iTextSharp.text.Element.ALIGN_LEFT;
                            pdfCell.AddElement(para);
                            pdfTab.AddCell(pdfCell);

                            pdfCell = new PdfPCell();
                            pdfCell.UseAscender = true;
                            pdfCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
                            pdfCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
                            para = new Paragraph("Gund.", baseFontNormal);
                            para.Alignment = iTextSharp.text.Element.ALIGN_LEFT;
                            pdfCell.AddElement(para);
                            pdfTab.AddCell(pdfCell);

                            pdfCell = new PdfPCell();
                            pdfCell.UseAscender = true;
                            pdfCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
                            pdfCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
                            para = new Paragraph("Present", baseFontNormal);
                            para.Alignment = iTextSharp.text.Element.ALIGN_LEFT;
                            pdfCell.AddElement(para);
                            pdfTab.AddCell(pdfCell);

                            pdfCell = new PdfPCell();
                            pdfCell.UseAscender = true;
                            pdfCell.VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE;
                            pdfCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT;
                            para = new Paragraph("13-Apr-19", baseFontNormal);
                            para.Alignment = iTextSharp.text.Element.ALIGN_RIGHT;
                            pdfCell.AddElement(para);
                            pdfTab.AddCell(pdfCell);
                        }

                        doc.Add(pdfTab);

                        doc.Close();
                    }
                }

                bytes = ms.ToArray();
            }

            using (MemoryStream stream = new MemoryStream())
            {
                PdfReader reader = new PdfReader(bytes);
                using (PdfStamper stamper = new PdfStamper(reader, stream))
                {
                    iTextSharp.text.Font baseFontNormal = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 9f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);
                    BaseFont bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                    int pages = reader.NumberOfPages;
                    for (int i = 1; i <= pages; i++)
                    {
                        String text = "Page " + i.ToString() + " of " + pages;
                        float len = bf.GetWidthPoint(text, 9);
                        var test = reader.GetPageSize(1);
                        var RightMargin = test.Width - len-23;
                        var phrase = new Phrase(text, baseFontNormal);
                        ColumnText.ShowTextAligned(stamper.GetUnderContent(i),
                        @Element.ALIGN_CENTER, phrase, RightMargin, 24f, 0);
                    }
                }

                bytes = stream.ToArray();

            }

            if (File.Exists("Output.pdf"))
            {
                File.Delete("Output.pdf");
            }

            File.WriteAllBytes("Output.pdf", bytes);
        }
    }

    public class ITextEvents : PdfPageEventHelper
    {
        // This is the contentbyte object of the writer
        PdfContentByte cb;

        // we will put the final number of pages in a template
        PdfTemplate headerTemplate,footerTemplate; 

         // this is the BaseFont we are going to use for the header / footer
         BaseFont bf = null;

        object objModal;

        // This keeps track of the creation time
        DateTime PrintTime = DateTime.Now;

        #region Fields
        private string _header;
        #endregion

        #region Properties
        public string Header
        {
            get { return _header; }
            set { _header = value; }
        }
        #endregion

        public ITextEvents(object obj)
        {
            objModal = obj;
        }

        public override void OnEndPage(iTextSharp.text.pdf.PdfWriter writer, iTextSharp.text.Document document)
        {
            base.OnEndPage(writer, document);
            try
            {
                PrintTime = DateTime.Now;
                bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                cb = writer.DirectContent;
                headerTemplate = cb.CreateTemplate(100, 200);
                footerTemplate = cb.CreateTemplate(50, 50);
            }
            catch (DocumentException de)
            {
            }
            catch (System.IO.IOException ioe)
            {
            }

            iTextSharp.text.Font baseFontBold = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 10f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
            iTextSharp.text.Font baseFontNormal = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);
            iTextSharp.text.Font baseFontBig = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 17f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);

            iTextSharp.text.Font greenFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 9f, iTextSharp.text.Font.BOLD, WebColors.GetRGBColor("#2daa2d"));
            Phrase p1Header = new Phrase("EMPLOYEE ATTENDANCE REPORT", baseFontBig);

            //Create PdfTable object
            PdfPTable pdfTab = new PdfPTable(3);
            PdfPTable pdfTab2 = new PdfPTable(2);

            pdfTab.DefaultCell.Border = Rectangle.NO_BORDER;
            pdfTab2.DefaultCell.Border = Rectangle.NO_BORDER;
            pdfTab.DefaultCell.BorderWidth = 0f;
            pdfTab2.DefaultCell.BorderWidth = 0f;

            float[] widths = new float[] { 20f, 50f,30f};
            float[] widths2 = new float[] { 13f, 87f};
            pdfTab.SetWidths(widths);
            pdfTab2.SetWidths(widths2);

            //We will have to create separate cells to include image logo and 2 separate strings
            //Row 1
            PdfPCell pdfCell1 = new PdfPCell();
            Image logo = Image.GetInstance("https://mne.seld.gos.pk/images/sindh-logo.png");
            pdfCell1.Border = Rectangle.NO_BORDER;
            logo.ScaleAbsolute(45.75f, 45.75f);
            pdfCell1.AddElement(logo);
            PdfPCell pdfCell2 = new PdfPCell(p1Header);
            pdfCell2.Border = Rectangle.NO_BORDER;
            PdfPCell pdfCell3 = new PdfPCell();
            pdfCell3.Border = Rectangle.NO_BORDER;
            var paraTop = new Paragraph("Directorate General of");
            paraTop.Font = greenFont;
            paraTop.Alignment= iTextSharp.text.Element.ALIGN_RIGHT;
            paraTop.SpacingAfter = 1f;
            pdfCell3.AddElement(paraTop);
            paraTop = new Paragraph("Monitoring and Evaluation");
            paraTop.Font = greenFont;
            paraTop.Alignment = iTextSharp.text.Element.ALIGN_RIGHT;
            paraTop.SpacingAfter = 1f;
            pdfCell3.AddElement(paraTop);
            paraTop = new Paragraph("School Education & Literacy Department");
            paraTop.Font = greenFont;
            paraTop.Alignment = iTextSharp.text.Element.ALIGN_RIGHT;
            paraTop.SpacingAfter = 1f;
            pdfCell3.AddElement(paraTop);
            paraTop = new Paragraph("www.sindheducation.gov.pk");
            paraTop.Font = greenFont;
            paraTop.Alignment = iTextSharp.text.Element.ALIGN_RIGHT;
            paraTop.SpacingAfter = 1f;
            pdfCell3.AddElement(paraTop);
            pdfCell3.UseAscender=(true);
            pdfCell3.UseDescender=(true);
            pdfCell3.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT;

            pdfTab.AddCell(pdfCell1);
            pdfTab.AddCell(pdfCell2);
            pdfTab.AddCell(pdfCell3);

            pdfCell1 = new PdfPCell();
            pdfCell1.Colspan = 3;
            string sampleText = "Employee name";
            paraTop = new Paragraph(sampleText.ToUpper(), baseFontBold);
            paraTop.Alignment = iTextSharp.text.Element.ALIGN_RIGHT;
            PdfPCell innerCell = new PdfPCell(paraTop);
            innerCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT;
            innerCell.Border = Rectangle.NO_BORDER;
            pdfTab2.AddCell(innerCell);
            sampleText = "ABDUL HAMEED / M.JUMAN";
            paraTop = new Paragraph(sampleText.ToUpper(), baseFontNormal);
            paraTop.SetLeading(12f, 0);
            paraTop.PaddingTop = 0;
            innerCell = new PdfPCell(paraTop);
            innerCell.UseAscender = true;
            innerCell.UseDescender = true;
            innerCell.PaddingTop=5;
            innerCell.PaddingLeft=10;
            innerCell.PaddingBottom=10;
            innerCell.Border = Rectangle.NO_BORDER;
            pdfTab2.AddCell(innerCell);

            sampleText = "FATHER NAME";
            paraTop = new Paragraph(sampleText.ToUpper(), baseFontBold);
            paraTop.Alignment = iTextSharp.text.Element.ALIGN_RIGHT;
            innerCell = new PdfPCell(paraTop);
            innerCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT;
            innerCell.Border = Rectangle.NO_BORDER;
            pdfTab2.AddCell(innerCell);

            sampleText = "MUHAMMAD JUMAN SARIO";
            paraTop = new Paragraph(sampleText.ToUpper(), baseFontNormal);
            paraTop.SetLeading(12f, 0);
            paraTop.PaddingTop = 0;
            innerCell = new PdfPCell(paraTop);
            innerCell.UseAscender = true;
            innerCell.UseDescender = true;
            innerCell.PaddingTop = 5;
            innerCell.PaddingLeft = 10;
            innerCell.PaddingBottom = 10;
            innerCell.Border = Rectangle.NO_BORDER;
            pdfTab2.AddCell(innerCell);
            
            sampleText = "CNIC";
            paraTop = new Paragraph(sampleText.ToUpper(), baseFontBold);
            paraTop.Alignment = iTextSharp.text.Element.ALIGN_RIGHT;
            innerCell = new PdfPCell(paraTop);
            innerCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT;
            innerCell.Border = Rectangle.NO_BORDER;
            pdfTab2.AddCell(innerCell);

            sampleText = "4320174263291";
            paraTop = new Paragraph(sampleText.ToUpper(), baseFontNormal);
            paraTop.SetLeading(12f, 0);
            paraTop.PaddingTop = 0;
            innerCell = new PdfPCell(paraTop);
            innerCell.UseAscender = true;
            innerCell.UseDescender = true;
            innerCell.PaddingTop = 5;
            innerCell.PaddingLeft = 10;
            innerCell.PaddingBottom = 10;
            innerCell.Border = Rectangle.NO_BORDER;
            pdfTab2.AddCell(innerCell);
            
            sampleText = "PERSONNEL ID";
            paraTop = new Paragraph(sampleText.ToUpper(), baseFontBold);
            paraTop.Alignment = iTextSharp.text.Element.ALIGN_RIGHT;
            innerCell = new PdfPCell(paraTop);
            innerCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT;
            innerCell.Border = Rectangle.NO_BORDER;
            pdfTab2.AddCell(innerCell);

            sampleText = "10208052";
            paraTop = new Paragraph(sampleText.ToUpper(), baseFontNormal);
            paraTop.SetLeading(12f, 0);
            paraTop.PaddingTop = 0;
            innerCell = new PdfPCell(paraTop);
            innerCell.UseAscender = true;
            innerCell.UseDescender = true;
            innerCell.PaddingTop = 5;
            innerCell.PaddingLeft = 10;
            innerCell.PaddingBottom = 10;
            innerCell.Border = Rectangle.NO_BORDER;
            pdfTab2.AddCell(innerCell);
            
            sampleText = "DESIGNATION";
            paraTop = new Paragraph(sampleText.ToUpper(), baseFontBold);
            paraTop.Alignment = iTextSharp.text.Element.ALIGN_RIGHT;
            innerCell = new PdfPCell(paraTop);
            innerCell.HorizontalAlignment = iTextSharp.text.Element.ALIGN_RIGHT;
            innerCell.Border = Rectangle.NO_BORDER;
            pdfTab2.AddCell(innerCell);

            sampleText = "PST";
            paraTop = new Paragraph(sampleText.ToUpper(), baseFontNormal);
            paraTop.SetLeading(12f, 0);
            paraTop.PaddingTop = 0;
            innerCell = new PdfPCell(paraTop);
            innerCell.UseAscender = true;
            innerCell.UseDescender = true;
            innerCell.PaddingTop = 5;
            innerCell.PaddingLeft = 10;
            innerCell.PaddingBottom = 10;
            innerCell.Border = Rectangle.NO_BORDER;
            pdfTab2.AddCell(innerCell);

            pdfTab2.TotalWidth = 100f;
            pdfTab2.WidthPercentage = 100;
            pdfCell1.AddElement(pdfTab2);

            pdfCell1.Border = Rectangle.NO_BORDER;
            pdfTab.AddCell(pdfCell1);

            pdfTab.TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
            pdfTab.WidthPercentage = 100;
            pdfTab.WriteSelectedRows(0, -1, document.LeftMargin, document.PageSize.Height - 80, writer.DirectContent);
            //set pdfContent value
            String text = "Page " + writer.PageNumber + " of " +document.PageNumber;
            //{
            //    cb.BeginText();
            //    cb.SetFontAndSize(bf, 12);
            //    cb.SetTextMatrix(document.PageSize.GetRight(180), document.PageSize.GetBottom(30));
            //    cb.ShowText(text);
            //    cb.EndText();
                float len = bf.GetWidthPoint(text, 12);
              cb.AddTemplate(footerTemplate, document.PageSize.GetRight(180) + len, document.PageSize.GetBottom(30));
            //}

            //set pdfContent value

            //Move the pointer and draw line to separate footer section from rest of page
            //cb.MoveTo(40, document.PageSize.GetBottom(50));
            //cb.LineTo(document.PageSize.Width - 40, document.PageSize.GetBottom(50));
            //cb.Stroke();
        }

        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            base.OnCloseDocument(writer, document);

            //footerTemplate.BeginText();
            //footerTemplate.SetFontAndSize(bf, 12);
            //footerTemplate.SetTextMatrix(0, 0);
            //footerTemplate.ShowText((writer.PageNumber).ToString());
            //footerTemplate.EndText();
        }
    }
}
