using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace SqlDocument.Models
{
    public class OenpXmlGenDoc
    {
        public void GenerateWord(List<SpInfo> spInfos)
        {
            string docPath = @"D:\\temp\\SpDoc22.docx";
            string docPath2 = @"D:\\temp\\temp.docx";
            //using (WordprocessingDocument doc = WordprocessingDocument.Create(docPath2, WordprocessingDocumentType.Document))
            using (WordprocessingDocument doc = WordprocessingDocument.Open(docPath2, true))
            {
                //doc.AddMainDocumentPart();

                TableGrid grid = new TableGrid();
                int maxColumnNum = 2;
                for (int index = 0; index < maxColumnNum; index++)
                {
                    grid.Append(new TableGrid());
                }


                // 設置表格邊框
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
               
                foreach (var spInfo in spInfos)
                {
                    Table table = new Table();
                    table.Append(tblProp);
                    // 添加表頭. 其實有TableHeader對象的,小弟用不來.
                    //TableRow headerRow = new TableRow();
                    //foreach (string headerStr in headerArray)
                    //{
                    //    TableCell cell = new TableCell();
                    //    cell.Append(new Paragraph(new Run(new Text(headerStr))));
                    //    headerRow.Append(cell);
                    //}
                    //table.Append(headerRow);
                    // 添加數據

                    TableRow rowA = new TableRow();
                    TableCell cellA = new TableCell();
                    cellA.Append(new Paragraph(new Run(new Text("預存程序名稱"))));
                    rowA.Append(cellA);
                    TableCell cellB = new TableCell();
                    cellB.Append(new Paragraph(new Run(new Text(spInfo.SpName))));
                    rowA.Append(cellB);
                    table.Append(rowA);


                    doc.MainDocumentPart.Document.Body.Append(new Paragraph(new Run(table)));
                    doc.Save();
                    //TableRow rowB = new TableRow();
                    //TableCell cellA2 = new TableCell();
                    //cellA2.Append(new Paragraph(new Run(new Text("AAAA"))));
                    //rowB.Append(cellA2);
                    //TableCell cellB2 = new TableCell();
                    //cellB2.Append(new Paragraph(new Run(new Text("DDDDD"))));
                    //rowB.Append(cellB2);

                    //table.Append(rowB);
                    doc.MainDocumentPart.Document.Body.Append(new Paragraph(new Run(table)));
                }
                //先試著產出文件, 看看是不是自己要的
                doc.Save();
            }
           
        }
    }
}
