using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace pdf_invoice
{
    public partial class Default : System.Web.UI.Page
    {
        protected Font normalFont = FontFactory.GetFont(FontFactory.HELVETICA, 9);
        protected Font boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12);

        protected Font boldFontSmall = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 9);
        protected void Page_Load(object sender, EventArgs e)
        {
            // Step 1: get the DataTable.
            DataTable table = SetUpData();
            // Step 4: print the first cell.
            double tax = 0.062;
            double subtotal = (double)table.Compute("Sum(AMOUNT)", "True");
            double SalesTax = subtotal*tax;
            double Total = subtotal + SalesTax;
            Response.Write(Total);
        }
        protected void Button1_Click(object sender, EventArgs e)
        {
            MemoryStream stream = new MemoryStream();
            stream = WriteDocument();
            Response.Clear();
            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", "attachment;filename=invoice.pdf");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            Response.BinaryWrite(stream.ToArray());
            Response.Flush();
            Response.Close();
            Response.End();
        }
        DataTable SetUpData()
        {
            //here we create a DataTable.
            // We add 4 columns, each with a Type.
            DataTable table = new DataTable();
            table.Columns.Add("QTY", typeof(int));
            table.Columns.Add("DESCRIPTION", typeof(string));
            table.Columns.Add("UNIT PRICE", typeof(double));
            table.Columns.Add("AMOUNT", typeof(double));

            //here we add 5 rows.
            table.Rows.Add(10, "Ferrero Rocher", 10, 10*10);
            table.Rows.Add(5, "Ghirardelli", 12, 5*12);
            table.Rows.Add(5, "Lindt & Sprüngli",4, 5*4);
            table.Rows.Add(2, "Toblerone",5, 2*5);
            table.Rows.Add(3, "Cadbury",2, 3*2);
            return table;
        }
        public MemoryStream WriteDocument()
        {
            MemoryStream stream = new MemoryStream();
            Document document = new Document(PageSize.A4, 25, 25, 25,25);
            PdfWriter writer;
            writer = PdfWriter.GetInstance(document, stream);
            document.Open();
            document.NewPage();
            document.Add(HeaderTable());
            document.Add(AdressTable());
            document.Add(ProductsTable());
            document.Close();

            return stream;
        }
        private PdfPTable HeaderTable()
        {
            PdfPTable table = new PdfPTable(2);
            table.WidthPercentage = 90f;
            int[] firstTablecellwidth = { 45, 45 };
            table.SetWidths(firstTablecellwidth);

            Phrase Ptitle = new Phrase();

            Ptitle.Add(new Chunk("Organisation name Inc \n \n", boldFont));
            Ptitle.Add(new Chunk("address 1 \n", normalFont));
            Ptitle.Add(new Chunk("address 2", normalFont));

            PdfPCell cell1 = new PdfPCell(new Phrase(Ptitle));
            cell1.SetLeading(15f, 0f);
            PdfPCell cell2 = new PdfPCell(new Phrase("Invoice", boldFont));

            cell1.Border = Rectangle.NO_BORDER;
            cell2.Border = Rectangle.NO_BORDER;

            cell2.HorizontalAlignment = Rectangle.ALIGN_RIGHT;
            table.AddCell(cell1);
            table.AddCell(cell2);
            table.SpacingAfter = 20f;
            return table;
        }
        private PdfPTable AdressTable()
        {
            PdfPTable table = new PdfPTable(4);
            table.WidthPercentage = 90f;
            int[] firstTablecellwidth = { 30, 30, 15,15 };
            table.SetWidths(firstTablecellwidth);

            //column1
            Phrase BillTo = new Phrase(6f);
            BillTo.Add(new Chunk("Bill To \n", boldFontSmall));
            BillTo.Add(new Chunk("address 1 \n", normalFont));
            BillTo.Add(new Chunk("address 2 \n", normalFont));
            BillTo.Add(new Chunk("address 2", normalFont));
            PdfPCell cell1 = new PdfPCell(BillTo);
            cell1.SetLeading(15f, 0f);
            cell1.Border = Rectangle.NO_BORDER;
            cell1.FixedHeight = 75f;
            //column2
            Phrase ShipTo = new Phrase();
            ShipTo.Add(new Chunk("Ship To \n", boldFontSmall));
            ShipTo.Add(new Chunk("address 1 \n", normalFont));
            ShipTo.Add(new Chunk("address 2 \n", normalFont));
            ShipTo.Add(new Chunk("address 2", normalFont));
            PdfPCell cell2 = new PdfPCell(ShipTo);
            cell2.SetLeading(15f, 0f);
            cell2.Border = Rectangle.NO_BORDER;
            cell2.FixedHeight = 75f;
            //column1
            Phrase Invoice = new Phrase();
            Invoice.Add(new Chunk("Invoice# \n", boldFontSmall));
            Invoice.Add(new Chunk("address 1 \n", normalFont));
            Invoice.Add(new Chunk("address 2 \n", normalFont));
            Invoice.Add(new Chunk("address 2", normalFont));
            PdfPCell cell3 = new PdfPCell(new Phrase(Invoice));
            cell3.SetLeading(15f, 0f);
            cell3.Border = Rectangle.NO_BORDER;
            cell3.FixedHeight = 75f;
            cell3.HorizontalAlignment = Rectangle.ALIGN_RIGHT;
            //column1
            Phrase US = new Phrase();
            US.Add(new Chunk("US-001 \n", boldFontSmall));
            US.Add(new Chunk("address 1 \n", normalFont));
            US.Add(new Chunk("address 2 \n", normalFont));
            US.Add(new Chunk("address 2", normalFont));
            PdfPCell cell4 = new PdfPCell(new Phrase(US));
            cell4.SetLeading(15f, 0f);
            cell4.Border = Rectangle.NO_BORDER;
            cell4.FixedHeight = 75f;
            cell4.HorizontalAlignment = Rectangle.ALIGN_RIGHT;

            table.AddCell(cell1);
            table.AddCell(cell2);
            table.AddCell(cell3);
            table.AddCell(cell4);
            table.SpacingAfter = 12.5f;
            return table;
        }
        private PdfPTable ProductsTable()
        {
            PdfPTable table = new PdfPTable(4);
            table.WidthPercentage = 90f;
            int[] firstTablecellwidth = { 10, 45, 15, 20 };
            table.SetWidths(firstTablecellwidth);

            //header columns
            PdfPCell cell1 = new PdfPCell(new Phrase("QTY", boldFontSmall));
            cell1.HorizontalAlignment = Rectangle.ALIGN_CENTER;
            cell1.BackgroundColor = new BaseColor(217, 217, 217);
            PdfPCell cell2 = new PdfPCell(new Phrase("DESCRIPTION", boldFontSmall));
            cell2.HorizontalAlignment = Rectangle.ALIGN_CENTER;
            cell2.BackgroundColor = new BaseColor(217, 217, 217);
            PdfPCell cell3 = new PdfPCell(new Phrase("UNIT PRICE", boldFontSmall));
            cell3.HorizontalAlignment = Rectangle.ALIGN_CENTER;
            cell3.BackgroundColor = new BaseColor(217, 217, 217);
            PdfPCell cell4 = new PdfPCell(new Phrase("AMOUNT", boldFontSmall));
            cell4.HorizontalAlignment = Rectangle.ALIGN_CENTER;
            cell4.BackgroundColor = new BaseColor(217, 217, 217);

            cell1.Padding = 5;
            cell2.Padding = 5;
            cell3.Padding = 5;
            cell4.Padding = 5;

            table.AddCell(cell1);
            table.AddCell(cell2);
            table.AddCell(cell3);
            table.AddCell(cell4);

            // Step 1: get the DataTable.
            DataTable Dtable = SetUpData();
            // Step 4: print the first cell.
            double tax = 0.062;
            double subtotal = (double)Dtable.Compute("Sum(AMOUNT)", "True");
            double SalesTax = subtotal * tax;
            double Total = subtotal + SalesTax;

            for (int i = 0; i < Dtable.Rows.Count; i++)
            {
                for (int j = 0; j < Dtable.Columns.Count ; j++){
                    PdfPCell p = new PdfPCell(new Phrase(Dtable.Rows[i][j].ToString(), normalFont));
                    if (j == 0) p.HorizontalAlignment = Rectangle.ALIGN_CENTER;
                    if (j > 1) p.HorizontalAlignment = Rectangle.ALIGN_RIGHT;
                    p.Padding = 5;
                    table.AddCell(p);
                }
            }

            cell1 = new PdfPCell(new Phrase("subtotal", normalFont));
            cell1.Colspan = 3;
            cell1.Border = PdfPCell.NO_BORDER;
            cell1.Border = PdfPCell.LEFT_BORDER | PdfPCell.RIGHT_BORDER | PdfPCell.TOP_BORDER ;
            cell2 = new PdfPCell(new Phrase(subtotal.ToString(), normalFont));
            cell2.Border = PdfPCell.NO_BORDER;
            cell2.Border = PdfPCell.LEFT_BORDER | PdfPCell.RIGHT_BORDER | PdfPCell.TOP_BORDER;
            cell1.HorizontalAlignment = Rectangle.ALIGN_RIGHT;
            cell2.HorizontalAlignment = Rectangle.ALIGN_RIGHT;
            cell1.Padding = 5;
            cell2.Padding = 5;
            table.AddCell(cell1);
            table.AddCell(cell2);

            cell1 = new PdfPCell(new Phrase("Sales Tax 6.25%", normalFont));
            cell1.Colspan = 3;
            cell1.Border = PdfPCell.NO_BORDER;
            cell1.Border = PdfPCell.LEFT_BORDER | PdfPCell.RIGHT_BORDER | PdfPCell.BOTTOM_BORDER;
            cell2 = new PdfPCell(new Phrase(SalesTax.ToString(), normalFont));
            cell2.Border = PdfPCell.NO_BORDER;
            cell2.Border = PdfPCell.LEFT_BORDER | PdfPCell.RIGHT_BORDER | PdfPCell.BOTTOM_BORDER;
            cell1.HorizontalAlignment = Rectangle.ALIGN_RIGHT;
            cell2.HorizontalAlignment = Rectangle.ALIGN_RIGHT;
            cell1.Padding = 5;
            cell2.Padding = 5;
            table.AddCell(cell1);
            table.AddCell(cell2);

            cell1 = new PdfPCell(new Phrase("Total", boldFontSmall));
            cell1.Colspan = 3;
            cell2 = new PdfPCell(new Phrase(Total.ToString(),boldFontSmall));
            cell1.HorizontalAlignment = Rectangle.ALIGN_RIGHT;
            cell2.HorizontalAlignment = Rectangle.ALIGN_RIGHT;
            cell1.Padding = 5;
            cell2.Padding = 5;
            table.AddCell(cell1);
            table.AddCell(cell2);
            
            return table;
        }
    }
}