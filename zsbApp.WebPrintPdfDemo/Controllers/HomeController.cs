
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace zsbApp.WebPrintPdfDemo.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult PdfViewer()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }


        public FileStreamResult CreatePdf()
        {
            //return this.getlocalFile("123.pdf");
            //return createFile();

            return createExcelFile();
        }


        private FileStreamResult createExcelFile()
        {
            string basePath = System.IO.Path.Combine(Server.MapPath("~"), "pdf");
            string template = System.IO.Path.Combine(basePath, "template.xlsx");

            string saveFilePath = System.IO.Path.Combine(basePath, $"{System.DateTime.Now.ToString("yyyyMMdd-HHmmss")}.pdf");

            var dt = this.getDataTable();

            Aspose.Cells.Workbook book = new Aspose.Cells.Workbook(template);
            Aspose.Cells.Worksheet sheet = book.Worksheets[0];
            
            for (int y = 0; y < 100; y++)
            {
                for (int x = 0; x < 10; x++)
                {
                    var obj = sheet.Cells[y, x].Value;
                    if (obj == null)
                        continue;
                    string cellValue = obj.ToString();
                    if (!cellValue.StartsWith("$"))
                        continue;
                    cellValue = cellValue.Replace("$", "");
                    if (cellValue == "序号")
                    {
                        this.processDetail(sheet, y);
                        break;
                    }
                    if (dt.Columns.Contains(cellValue))
                    sheet.Cells[y, x].PutValue(dt.Rows[0][cellValue].ToString());
                    
                }
            }
            book.Save(saveFilePath, Aspose.Cells.SaveFormat.Pdf);

            return getlocalFileFullPath(saveFilePath);
        }

        private void processDetail(Aspose.Cells.Worksheet sheet, int rowIndex)
        {
            var dtDetail = this.getDataTableDetail();
            var header = new List<string>();
            for (int i = 0; i < 10; i++)
            {
                var obj = sheet.Cells[rowIndex, i].Value;
                if (obj == null)
                    break;
                header.Add(obj.ToString().Replace("$", ""));
            }
            var header_datatable_index = new int[header.Count];
            for (int i = 0; i < header.Count; i++)
            {
                header_datatable_index[i] = dtDetail.Columns.IndexOf(header[i]);
            }
            int index = rowIndex;
            for (int i = 0; i < dtDetail.Rows.Count; i++)
            {
                for (int j = 0; j < header_datatable_index.Length; j++)
                {
                    int table_column_index = header_datatable_index[j];
                    if (table_column_index == -1) continue;
                    sheet.Cells[index, j].PutValue(dtDetail.Rows[i][table_column_index].ToString());
                }
                index++;
            }
        }

        private System.Data.DataTable getDataTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("委托单位", typeof(string));
            dt.Columns.Add("工程名称", typeof(string));
            dt.Columns.Add("结算日期", typeof(string));
            dt.Rows.Add("中润高铁实验室中润大厦试验部", "中润高铁实验室", System.DateTime.Now.AddDays(-10).ToString("yyyy-MM-dd"));
            return dt;
        }

        private System.Data.DataTable getDataTableDetail()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("序号", typeof(int));
            dt.Columns.Add("委托日期", typeof(string));
            dt.Columns.Add("报告编号", typeof(string));
            dt.Columns.Add("检测项目", typeof(string));
            dt.Columns.Add("单位", typeof(string));
            dt.Columns.Add("数量", typeof(int));
            dt.Columns.Add("单价", typeof(int));
            dt.Columns.Add("辅助费用", typeof(int));
            dt.Columns.Add("实收金额", typeof(int));

            string[] arrayItems = new string[] { "水泥", "混泥土抗压", "粉煤灰", "环境水", "钢筋原材" };
            int[] arrayPrice = new int[] { 100, 200, 300, 400, 500 };
            int[] arrayPriceOther = new int[] { 50, 60, 70, 80, 90 };

            System.Random random = new Random();
            var date = System.DateTime.Now.AddDays(-100);

            for (int i = 1; i < 10; i++)
            {
                var dr = dt.NewRow();
                var itemIndex = random.Next(0, 4);
                int count = random.Next(1, 10);

                dr["序号"] = i;
                dr["委托日期"] = date.ToString("yyyy-MM-dd");
                dr["报告编号"] = $"TJSN18{random.Next(100000, 999999)}";
                dr["检测项目"] = arrayItems[itemIndex];
                dr["单位"] = "件";
                dr["数量"] = count;
                dr["单价"] = arrayPrice[itemIndex];
                dr["辅助费用"] = arrayPriceOther[itemIndex];
                dr["实收金额"] = count * arrayPrice[itemIndex] + arrayPriceOther[itemIndex];
                dt.Rows.Add(dr);
            }

            return dt;

        }


        private FileStreamResult createFile()
        {
            string fileName = $"{System.DateTime.Now.ToString("yyyyMMdd-HHmmss")}.pdf";
            string filePath = System.IO.Path.Combine(Server.MapPath("~"), "pdf");
            filePath = System.IO.Path.Combine(filePath, fileName);

            Document doc = new Document(PageSize.A4, 0, 0, 0, 0);
            PdfWriter write = PdfWriter.GetInstance(doc, new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write));
            doc.Open();

            //中文字体
            string chinese = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "KAIU.TTF");
            BaseFont bsFont = BaseFont.CreateFont(@"C:\Windows\Fonts\simsun.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            var font = new iTextSharp.text.Font(bsFont, 12);
            doc.Add(new Paragraph("第一个PDF文件", font));//将一句短语写入PDF中

            doc.Close();//关闭

            return getlocalFile(fileName);
        }

        private FileStreamResult getlocalFile(string fileName)
        {
            string filePath = System.IO.Path.Combine(Server.MapPath("~"), "pdf");
            filePath = System.IO.Path.Combine(filePath, fileName);
            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            return File(fs, "application/pdf");
        }

        private FileStreamResult getlocalFileFullPath(string fileName)
        {
            FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            return File(fs, "application/pdf");
        }
    }
}