using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectOpenXML
{
    public partial class Form1 : Form
    {
        string[] headerArray = { "h1", "h2", "h3" };
        string[] rowArray = { "1", "2", "3" };

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(
                System.IO.Path.GetTempPath() + "HelloWorld.docx",
                WordprocessingDocumentType.Document)
            )
            {
                TableGrid grid = new TableGrid();
                int maxColumnNum = 3;
                for (int index = 0; index < maxColumnNum; index++)
                {
                    grid.Append(new TableGrid());
                }

                // 设置表格边框
                TableProperties tblProp = new TableProperties(
                new TableBorders(
                new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 },
                new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 },
                new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 },
                new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 },
                new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 },
                new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 }
                )
                );

                Table table = new Table();
                table.Append(tblProp);

                // 添加表头. 其实有TableHeader对象的,小弟用不来.
                TableRow headerRow = new TableRow();
                foreach (string headerStr in headerArray)
                {
                    TableCell cell = new TableCell();
                    cell.Append(new Paragraph(new Run(new Text(headerStr))));
                    headerRow.Append(cell);
                }
                table.Append(headerRow);

                // 添加数据
                for (int i = 0; i < 3; i++)
                {
                    TableRow row = new TableRow();
                    foreach (string strCell in rowArray)
                    {
                        TableCell cell = new TableCell();
                        cell.Append(new Paragraph(new Run(new Text(strCell))));
                        row.Append(cell);
                    }
                    table.Append(row);
                }

                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                Paragraph p = mainPart.Document.Body.AppendChild(new Paragraph());
                p.AppendChild(new Run(new Text("Test")));

                body.Append(new Paragraph(new Run(table)));

                doc.MainDocumentPart.Document.Save();
            }
            using (System.Diagnostics.Process xx = System.Diagnostics.Process.Start(
                System.IO.Path.GetTempPath() + "HelloWorld.docx"))
            {
            }
        }
    }
}
