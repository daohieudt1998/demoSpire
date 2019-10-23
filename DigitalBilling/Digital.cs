using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Text;

namespace DigitalBilling
{
    public class Digital
    {
        public string create()
        {
            string nameCompany = "Công ty định vị Bách Khoa";
            string taxCompany = "101010101010";
            string addressCompany = "560 Nguyễn Bỉnh Khiêm Hải Phòng";
            string phoneCompany = "0775 225 222";
            string accountNumber = "12121212121";
            //data info Customer
            string customerName = "Trần Tăng Đoan";
            string customerCompany = "Định vị bách khoa";
            string taxCustomerCompany = "13131313131";
            string customerCompanyAddress = "Big C, Nguyễn Bỉnh Khiêm, Hải Phòng";
            string kingOfPayment = "Visa";
            string customerAccountNumber = "19001009";
            //Code search eBilling
            string codeSearch = Guid.NewGuid().ToString();
            Document doc = new Document();

            //Gọi file mẫu
            doc.LoadFromFile(@"D:\Demo\demoSpire\demoSpire\Content\demo.docx");

            //info Company sale
            doc.Replace("<nameCompany>", nameCompany.ToUpper(), true, true);
            doc.Replace("<taxCompany>", taxCompany, true, true);
            doc.Replace("<addressCompany>", addressCompany, true, true);
            doc.Replace("<phoneCompany>", phoneCompany, true, true);
            doc.Replace("<accountNumber>", accountNumber, true, true);
            //info Customer 
            doc.Replace("<customerName>", customerName, true, true);
            doc.Replace("<customerCompany>", customerCompany, true, true);
            doc.Replace("<taxCustomerCompany>", taxCustomerCompany, true, true);
            doc.Replace("<customerCompanyAddress>", customerCompanyAddress, true, true);
            doc.Replace("<kingOfPayment>", kingOfPayment, true, true);
            doc.Replace("<customerAccountNumber>", customerAccountNumber, true, true);
            //Code search eBilling
            doc.Replace("<codeSearch>", codeSearch, true, true);
            //create Table
            Table table = new Table(doc, true);
            DataTable dt = new DataTable();
            PreferredWidth width = new PreferredWidth(WidthType.Percentage, 100);
            table.PreferredWidth = width;
            dt.Columns.Add("STT", typeof(string));
            dt.Columns.Add("Tên hang hóa, dịch vụ", typeof(string));
            dt.Columns.Add("Đơn vị tính", typeof(string));
            dt.Columns.Add("Số lượng", typeof(string));
            dt.Columns.Add("Đơn giá", typeof(string));
            dt.Columns.Add("Thành tiền", typeof(string));
            dt.Rows.Add(new string[] { "STT", "Tên hàng hóa, dịch vụ", "Đơn vị tính", "Số lượng", "Đơn giá", "Thành tiền" });
            int n = 5;
            for (int i = 0; i < n; i++)
            {
                dt.Rows.Add(new String[] { i.ToString(), i.ToString(), i.ToString(), i.ToString(), i.ToString(), i.ToString()
});
            }
            dt.Rows.Add(new string[] { "Cộng tiền hàng: " });
            dt.Rows.Add(new string[] { "Tiền thuế giá trị gia tăng: " });
            dt.Rows.Add(new string[] { "Tổng tiền thanh toán: " });
            table.ResetCells(dt.Rows.Count, dt.Columns.Count);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    table.Rows[i].Cells[j].AddParagraph().AppendText(dt.Rows[i][j].ToString());
                }
            }
            table.Rows[0].Cells[0].SetCellWidth(7, CellWidthType.Percentage);
            table.Rows[0].Cells[1].SetCellWidth(30, CellWidthType.Percentage);
            table.Rows[0].Cells[2].SetCellWidth(15, CellWidthType.Percentage);
            table.Rows[0].Cells[3].SetCellWidth(10, CellWidthType.Percentage);
            table.Rows[0].Cells[4].SetCellWidth(18, CellWidthType.Percentage);
            table.Rows[0].Cells[5].SetCellWidth(20, CellWidthType.Percentage);
            table.ApplyHorizontalMerge(n + 1, 0, 5);
            table.ApplyHorizontalMerge(n + 2, 0, 5);
            table.ApplyHorizontalMerge(n + 3, 0, 5);
            Section section = doc.Sections[0];
            TextSelection selection = doc.FindString("[billing_Datatable]", true, true); //find position to customize
            TextRange range = selection.GetAsOneRange();
            Paragraph paragraph = range.OwnerParagraph;
            Body body = paragraph.OwnerTextBody;
            int index = body.ChildObjects.IndexOf(paragraph);
            body.ChildObjects.Remove(paragraph); //remove string
            body.ChildObjects.Insert(index, table); //insert datatable to position

            //save file
            doc.SaveToFile(@"D:\Demo\demoSpire\demoSpire\wwwroot\test4.pdf", Spire.Doc.FileFormat.PDF);
            PdfDocument pdf = new PdfDocument();
            string pathFilePdf = @"D:\Demo\demoSpire\demoSpire\wwwroot\test4.pdf";
            pdf.LoadFromFile(pathFilePdf);
            PdfPageBase page = pdf.Pages[0];
            Image backgroundImage = Image.FromFile(@"D:\Demo\demoSpire\demoSpire\wwwroot\background.jpg");
            page.BackgroundImage = backgroundImage;
            pdf.SaveToFile(pathFilePdf);
            return pathFilePdf;
        }

        public string create1()
        {
            //data info Customer
            string customerName = "Trần Tăng Đoan";
            string customerCompany = "Định vị bách khoa";
            string taxCustomerCompany = "13131313131";
            string customerCompanyAddress = "Big C, Nguyễn Bỉnh Khiêm, Hải Phòng";
            string kingOfPayment = "Visa";
            string customerAccountNumber = "19001009";
            var dd = DateTime.Now.ToString("MM");
            var mm = DateTime.Now.ToString("MM");
            var yyyy = DateTime.Now.ToString("yyyy");
            //Code search eBilling
            string codeSearch = Guid.NewGuid().ToString();
            int rateTax = 2;
            Document doc = new Document();

            //Gọi file mẫu
            doc.LoadFromFile(@"D:\Demo\demoSpire\demoSpire\wwwroot\sample1.doc");


            //info Customer 
            doc.Replace("<dd>", dd, true, true);
            doc.Replace("<mm>", mm, true, true);
            doc.Replace("<yy>", yyyy, true, true);
            doc.Replace("<customerName>", customerName, true, true);
            doc.Replace("<customerCompany>", customerCompany, true, true);
            doc.Replace("<taxCustomerCompany>", taxCustomerCompany, true, true);
            doc.Replace("<addressCustomer>", customerCompanyAddress, true, true);
            doc.Replace("<Payment>", kingOfPayment, true, true);
            doc.Replace("<accountNumber>", customerAccountNumber, true, true);
            //Code search eBilling
            doc.Replace("<codeSearch>", codeSearch, true, true);
            //create Table
            Table table = new Table(doc, true);
            DataTable dt = new DataTable();
            PreferredWidth width = new PreferredWidth(WidthType.Percentage, 100);
            table.PreferredWidth = width;
            dt.Columns.Add("STT", typeof(string));
            dt.Columns.Add("Tên hang hóa, dịch vụ", typeof(string));
            dt.Columns.Add("Đơn vị tính", typeof(string));
            dt.Columns.Add("Số lượng", typeof(string));
            dt.Columns.Add("Đơn giá", typeof(string));
            dt.Columns.Add("Thành tiền", typeof(string));
            dt.Rows.Add(new string[] { "STT", "Tên hàng hóa, dịch vụ", "Đơn vị tính", "Số lượng", "Đơn giá", "Thành tiền" });
            int n = 5;
            for (int i = 0; i < n; i++)
            {
                dt.Rows.Add(new String[] { i.ToString(), i.ToString(), i.ToString(), i.ToString(), i.ToString(), i.ToString()
});
            }
            dt.Rows.Add(new string[] { "Cộng tiền hàng (Total amount):" });
            dt.Rows.Add(new string[] { "Thuế suất GTGT (VAT rate):" + rateTax + "%" });
            dt.Rows.Add(new string[] { "Tiền thuế GTGT (VAT amount):"});
            dt.Rows.Add(new string[] { "Tổng cộng tiền thanh toán (Total payment):" });
            table.ResetCells(dt.Rows.Count, dt.Columns.Count);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    table.Rows[i].Cells[j].AddParagraph().AppendText(dt.Rows[i][j].ToString());
                }
            }
            table.Rows[0].Cells[0].SetCellWidth(7, CellWidthType.Percentage);
            table.Rows[0].Cells[1].SetCellWidth(30, CellWidthType.Percentage);
            table.Rows[0].Cells[2].SetCellWidth(15, CellWidthType.Percentage);
            table.Rows[0].Cells[3].SetCellWidth(10, CellWidthType.Percentage);
            table.Rows[0].Cells[4].SetCellWidth(18, CellWidthType.Percentage);
            table.Rows[0].Cells[5].SetCellWidth(20, CellWidthType.Percentage);
            table.ApplyHorizontalMerge(n + 1, 0, 5);
            table.ApplyHorizontalMerge(n + 2, 0, 5);
            table.ApplyHorizontalMerge(n + 3, 0, 5);
            table.ApplyHorizontalMerge(n + 4, 0, 5);
            Section section = doc.Sections[0];
            TextSelection selection = doc.FindString("[billing_Datatable]", true, true); //find position to customize
            TextRange range = selection.GetAsOneRange();
            Paragraph paragraph = range.OwnerParagraph;
            Body body = paragraph.OwnerTextBody;
            int index = body.ChildObjects.IndexOf(paragraph);
            body.ChildObjects.Remove(paragraph); //remove string
            body.ChildObjects.Insert(index, table); //insert datatable to position

            //save file
            doc.SaveToFile(@"D:\Demo\demoSpire\demoSpire\wwwroot\test3.pdf", Spire.Doc.FileFormat.PDF);
            PdfDocument pdf = new PdfDocument();
            string pathFilePdf = @"D:\Demo\demoSpire\demoSpire\wwwroot\test3.pdf";
            pdf.LoadFromFile(pathFilePdf);
            PdfPageBase page = pdf.Pages[0];
            Image backgroundImage = Image.FromFile(@"D:\Demo\demoSpire\demoSpire\wwwroot\background.jpg");
            page.BackgroundImage = backgroundImage;
            pdf.SaveToFile(pathFilePdf);
            return pathFilePdf;
        }

    }
}
