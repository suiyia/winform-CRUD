using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Formatting;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StatisticalTools.utils
{
    class Test
    {
        public void fun() {

            //创建Word文档 
            Document document = new Document();
            //添加section
            Section section = document.AddSection();
            // 
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("订单编号");
            dataTable.Columns.Add("客户名称");
            dataTable.Columns.Add("型号");
            dataTable.Columns.Add("种类");
            dataTable.Columns.Add("颜色");
            dataTable.Columns.Add("数量");
            dataTable.Columns.Add("单价");
            dataTable.Columns.Add("总金额");
            dataTable.Columns.Add("其它");
            for (int i = 0; i < 10; i++) {
                DataRow dr = dataTable.NewRow();
                dr["订单编号"] = "12312312312"; dr["客户名称"] = "周杰伦"; dr["型号"] = "123141";
                dr["种类"] = "123141"; dr["颜色"] = "123141"; dr["数量"] = 123141;
                dr["单价"] = 123141.99; dr["总金额"] = 123141.99; dr["其它"] = "123141";
                dataTable.Rows.Add(dr);
            }

            Paragraph para1 = section.AddParagraph();
            para1.AppendText("硕麒拉链码庄送货单");
            para1.Format.HorizontalAlignment = HorizontalAlignment.Center;
            para1.AppendBreak(BreakType.LineBreak);
            ParagraphStyle style1 = new ParagraphStyle(document);
            style1.Name = "titleStyle";
            style1.CharacterFormat.Bold = true;
            style1.CharacterFormat.TextColor = Color.Black;
            style1.CharacterFormat.FontName = "黑体";
            style1.CharacterFormat.FontSize = 26f;
            style1.CharacterFormat.CharacterSpacing = document.Styles.Add(style1);
            para1.ApplyStyle("titleStyle");


            //   添加表格  表格样式
            Table table = section.AddTable(true);
            //指定表格的行数和列数（）
            table.ResetCells(9 + dataTable.Rows.Count, 9);
           

            //固定项 样式
            CharacterFormat charactertitle = new CharacterFormat(document);
            charactertitle.Bold = true;
            charactertitle.FontSize = 15;
            charactertitle.TextColor = Color.Black;
            charactertitle.FontName = "黑体";

            TextRange range = table[0, 0].AddParagraph().AppendText("客户姓名");
            range.ApplyCharacterFormat(charactertitle);
            TextRange range1 = table[1, 0].AddParagraph().AppendText("电话");
            range1.ApplyCharacterFormat(charactertitle);
            TextRange range3 = table[0, 5].AddParagraph().AppendText("单号");
            range3.ApplyCharacterFormat(charactertitle);
            TextRange range4 = table[1, 5].AddParagraph().AppendText("开票日期");
            range4.ApplyCharacterFormat(charactertitle);
            TextRange ra1 = table[3, 0].AddParagraph().AppendText("序号");
            ra1.ApplyCharacterFormat(charactertitle);
            TextRange ra2 = table[3, 1].AddParagraph().AppendText("系统编号");
            ra2.ApplyCharacterFormat(charactertitle);
            TextRange ra3 = table[3, 2].AddParagraph().AppendText("型号");
            ra3.ApplyCharacterFormat(charactertitle);
            TextRange ra4 = table[3, 3].AddParagraph().AppendText("种类");
            ra4.ApplyCharacterFormat(charactertitle);
            TextRange ra5 = table[3, 4].AddParagraph().AppendText("颜色");
            ra5.ApplyCharacterFormat(charactertitle);
            TextRange ra6 = table[3, 5].AddParagraph().AppendText("数量");
            ra6.ApplyCharacterFormat(charactertitle);
            TextRange ra7 = table[3, 6].AddParagraph().AppendText("单价(元)");
            ra7.ApplyCharacterFormat(charactertitle);
            TextRange ra8 = table[3, 7].AddParagraph().AppendText("总金额(元)");
            ra8.ApplyCharacterFormat(charactertitle);
            TextRange ra9 = table[3, 8].AddParagraph().AppendText("备注");
            ra9.ApplyCharacterFormat(charactertitle);
            // 空一行出来
            table.ApplyHorizontalMerge(2, 0, 8);

            // 尾部 第一行3列
            table.ApplyHorizontalMerge(dataTable.Rows.Count + 5, 2, 8);
            TextRange ra10 = table[dataTable.Rows.Count + 5, 0].AddParagraph().AppendText("未付款");
            ra10.ApplyCharacterFormat(charactertitle);
            TextRange ra11 = table[dataTable.Rows.Count + 5, 1].AddParagraph().AppendText("合计");
            ra11.ApplyCharacterFormat(charactertitle);
            // 总金额  —— 数字形式
            TextRange ra12 = table[dataTable.Rows.Count + 5, 2].AddParagraph().AppendText("99999  （元）");
            ra12.ApplyCharacterFormat(charactertitle);

            // 尾部第二行 2列
            table.ApplyHorizontalMerge(dataTable.Rows.Count + 6, 0,1);
            TextRange ra13 = table[dataTable.Rows.Count + 6, 0].AddParagraph().AppendText("大写金额");
            ra13.ApplyCharacterFormat(charactertitle);
            // 待写入大写金额
            ExportToWord word = new ExportToWord();
            table.ApplyHorizontalMerge(dataTable.Rows.Count + 6, 2, 8);
            TextRange ra14 = table[dataTable.Rows.Count + 6, 2].AddParagraph().AppendText(word.ConvertSum("8888.88"));
            Console.WriteLine(word.ConvertSum("8888.88"));
            ra14.ApplyCharacterFormat(charactertitle);

            // 尾部第3行 1列
            table.ApplyHorizontalMerge(dataTable.Rows.Count + 7, 0, 8);
            TextRange ra15 = table[dataTable.Rows.Count + 7, 0].AddParagraph().AppendText("地址： 汉川市新河汉正大道思嘉工业园       电话 ： 0712-8105088");
            ra15.ApplyCharacterFormat(charactertitle);

            // 尾部第4行 2列
            table.ApplyHorizontalMerge(dataTable.Rows.Count + 8, 0, 4);
            table.ApplyHorizontalMerge(dataTable.Rows.Count + 8, 5, 8);
            TextRange ra16 = table[dataTable.Rows.Count + 8, 0].AddParagraph().AppendText("制单： ");
            ra16.ApplyCharacterFormat(charactertitle);
            TextRange ra17 = table[dataTable.Rows.Count + 8, 5].AddParagraph().AppendText("客户签名： ");
            ra17.ApplyCharacterFormat(charactertitle);


            // 待填充部分
            // 姓名
            table.ApplyHorizontalMerge(0, 1, 4);
            TextRange r1 = table[0, 1].AddParagraph().AppendText("周杰伦");
            r1.ApplyCharacterFormat(charactertitle);
            // 单号
            table.ApplyHorizontalMerge(0, 6, 8);
            TextRange r2 = table[0, 6].AddParagraph().AppendText("CL-2018-01-01");
            r2.ApplyCharacterFormat(charactertitle);
            // 电话
            table.ApplyHorizontalMerge(1, 1, 4);
            TextRange r3 = table[1, 1].AddParagraph().AppendText("15072464246");
            r3.ApplyCharacterFormat(charactertitle);
            // 开票日期
            table.ApplyHorizontalMerge(1, 6, 8);
            TextRange r4 = table[1, 6].AddParagraph().AppendText("2018-01-01");
            r4.ApplyCharacterFormat(charactertitle);
            


            // 添加的数据完毕
            for (int i = 0; i < dataTable.Rows.Count; i++) {
                table[i+4, 0].AddParagraph().AppendText(i+"");
                DataRow rows = dataTable.Rows[i];
                int temp = 1;
                for (int j = 0; j < 9; j++) {
                    if (j != 1) { 
                        table[i + 4, temp].AddParagraph().AppendText(rows[j].ToString());
                        temp++;
                    }
                }
            }

            // 单元格样式
            table[0, 0].CellFormat.FitText = true;
            table[5 + dataTable.Rows.Count, 0].CellFormat.FitText = true;  // 未付款
            //table[8 + dataTable.Rows.Count, 5].CellFormat.FitText = true;
            for (int i = 3; i < 5 + dataTable.Rows.Count; i++)
            {
                for (int j = 0; j < 9; j++)
                {
                    table[i, j].CellFormat.FitText = true;
                    table.DefaultRowHeight = 25;
                }
            }

            //保存文档
            document.SaveToFile("C:\\Users\\answer\\Desktop\\Table.docx");
        }

    }
}
