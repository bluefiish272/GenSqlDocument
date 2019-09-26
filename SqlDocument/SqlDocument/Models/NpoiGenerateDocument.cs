using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace SqlDocument.Models
{
    public class NpoiGenerateDocument
    {
        public void GenerateWord(List<SpInfo> spInfos)
        {
            XWPFDocument doc = new XWPFDocument();

            for (int i = 0; i < spInfos.Count; i++)
            {
                XWPFTable table = doc.CreateTable();
                var firstRow = table.GetRow(0);

                firstRow.GetCell(0).SetText("預存程序名稱");
                firstRow.AddNewTableCell().SetText(spInfos[i].SpName);
                CT_TcPr m_Pr = firstRow.GetCell(0).GetCTTc().AddNewTcPr();
                m_Pr = Set(m_Pr);
               
                var secondRow = table.CreateRow();
                secondRow.GetCell(0).SetText("建立日期");
                secondRow.GetCell(1).SetText($"{spInfos[i].CreateDate:yyyy/MM/dd}");
             
                var thirdRow = table.CreateRow();
                thirdRow.GetCell(0).SetText("修改日期");
                thirdRow.GetCell(1).SetText($"{spInfos[i].ModifyDate:yyyy/MM/dd}");

                var forthRow = table.CreateRow();
                forthRow.GetCell(0).SetText("說明");
                forthRow.GetCell(1).SetText(spInfos[i].Description);

                var fifthRow = table.CreateRow();
                fifthRow.GetCell(0).SetText("Script");
                fifthRow.GetCell(1).SetText(spInfos[i].Script);

                doc.CreateParagraph();
            }

            // 產出文件
            // path可以自己換
            using (FileStream fileStream = new FileStream(@"D:\\temp\\SpDoc.docx", FileMode.Create))
            {
                doc.Write(fileStream);
            }
        }

        private static CT_TcPr Set(CT_TcPr cT_TcPr)
        {
            cT_TcPr.tcW = new CT_TblWidth();
            cT_TcPr.tcW.w = "2100";
            cT_TcPr.tcW.type = ST_TblWidth.dxa; //设置单元格宽度
            return cT_TcPr;
        }
    }
}
