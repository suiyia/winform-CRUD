using PapersStatisticTools;
using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using System.Data;
using Spire.Doc.Fields;
using Spire.Doc.Formatting;
using NPinyin;
using System.Text;

namespace StatisticalTools.utils
{
    class ExportToWord
    {
        // 客户名  总金额
        public void ExportToWordFun(System.Data.DataTable dataTable, String name, double sum)
        {
            System.Windows.Forms.SaveFileDialog saveDia = new SaveFileDialog();
            saveDia.Filter = "Word|*.docx";
            saveDia.Title = "导出为Word文件";
            saveDia.FileName = DateTime.Now.ToString("yyyyMMddhhss") + name;
            if (saveDia.ShowDialog() == System.Windows.Forms.DialogResult.OK
             && !string.Empty.Equals(saveDia.FileName))
            {

                MysqlManager mysqlManager = new MysqlManager();
                String phone = mysqlManager.getPhone(name);
                if (phone == null) {
                    phone = "00000000000";
                }

                //创建Word文档 
                Document document = new Document();
                //添加section
                Section section = document.AddSection();

                section.PageSetup.PageSize = new System.Drawing.SizeF(600, 420);
                //section.PageSetup.PageSize = PageSize.A4;
                section.PageSetup.Orientation = PageOrientation.Landscape;

                Paragraph para1 = section.AddParagraph();
                para1.AppendText("硕麒拉链码庄送货单");
                para1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                para1.AppendBreak(BreakType.LineBreak);
                ParagraphStyle style1 = new ParagraphStyle(document);
                style1.Name = "titleStyle";
                style1.CharacterFormat.Bold = true;
                style1.CharacterFormat.TextColor = Color.Black;
                style1.CharacterFormat.FontName = "黑体";
                style1.CharacterFormat.FontSize = 18f;
                style1.CharacterFormat.CharacterSpacing = document.Styles.Add(style1);
                para1.ApplyStyle("titleStyle");


                //   添加表格  表格样式
                Table table = section.AddTable(true);
                //指定表格的行数和列数（）
                table.ResetCells(9 + dataTable.Rows.Count, 9);


                //固定项 样式
                CharacterFormat charactertitle = new CharacterFormat(document);
                charactertitle.Bold = true;
                charactertitle.FontSize = 13;
                charactertitle.TextColor = Color.Black;
                charactertitle.FontName = "黑体";

                TextRange range = table[0, 0].AddParagraph().AppendText("客户姓名");
                range.ApplyCharacterFormat(charactertitle);
                TextRange range1 = table[1, 0].AddParagraph().AppendText("客户电话");
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
                TextRange ra8 = table[3, 7].AddParagraph().AppendText("金额(元)");
                ra8.ApplyCharacterFormat(charactertitle);
                TextRange ra9 = table[3, 8].AddParagraph().AppendText("备注");
                ra9.ApplyCharacterFormat(charactertitle);
                // 空一行出来
                table.ApplyHorizontalMerge(2, 0, 8);

                // 尾部 第一行3列
                table.ApplyHorizontalMerge(dataTable.Rows.Count + 5, 2, 8);
                TextRange ra10 = table[dataTable.Rows.Count + 5, 0].AddParagraph().AppendText("未付款");
                ra10.ApplyCharacterFormat(charactertitle);
                TextRange ra11 = table[dataTable.Rows.Count + 5, 1].AddParagraph().AppendText("合计金额");
                ra11.ApplyCharacterFormat(charactertitle);
                //                待填入   总金额  —— 数字形式
                TextRange ra12 = table[dataTable.Rows.Count + 5, 2].AddParagraph().AppendText(sum+"（元）");
                ra12.ApplyCharacterFormat(charactertitle);

                // 尾部第二行 2列
                table.ApplyHorizontalMerge(dataTable.Rows.Count + 6, 0, 1);
                TextRange ra13 = table[dataTable.Rows.Count + 6, 0].AddParagraph().AppendText("大写金额");
                ra13.ApplyCharacterFormat(charactertitle);
                //                   待填入   写入大写金额
                ExportToWord word = new ExportToWord();
                table.ApplyHorizontalMerge(dataTable.Rows.Count + 6, 2, 8);
                TextRange ra14 = table[dataTable.Rows.Count + 6, 2].AddParagraph().AppendText(word.ConvertSum(sum.ToString()));
                //Console.WriteLine(word.ConvertSum("8888.88"));
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
                TextRange r1 = table[0, 1].AddParagraph().AppendText(name);
                r1.ApplyCharacterFormat(charactertitle);
                // 单号
                table.ApplyHorizontalMerge(0, 6, 8);
                Encoding gb2312 = Encoding.GetEncoding("GB2312");
                String s = Pinyin.ConvertEncoding(name, Encoding.UTF8, gb2312);
                String s1 = Pinyin.GetInitials(s, gb2312);
                TextRange r2 = table[0, 6].AddParagraph().AppendText(s1+"-"+ DateTime.Now.Year+"-"+DateTime.Now.Month+"-"+DateTime.Now.Day+"-"+mysqlManager.getGuestOrderCount(name));
                r2.ApplyCharacterFormat(charactertitle);
                // 电话
                table.ApplyHorizontalMerge(1, 1, 4);
                TextRange r3 = table[1, 1].AddParagraph().AppendText(phone);
                r3.ApplyCharacterFormat(charactertitle);
                // 开票日期
                table.ApplyHorizontalMerge(1, 6, 8);
                TextRange r4 = table[1, 6].AddParagraph().AppendText(DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day+" "+DateTime.Now.Hour + ":" + DateTime.Now.Minute);
                r4.ApplyCharacterFormat(charactertitle);
                // 添加的数据完毕
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    table[i + 4, 0].AddParagraph().AppendText((i+1) + "");
                    DataRow rows = dataTable.Rows[i];
                    int temp = 1;
                    for (int j = 0; j < 9; j++)
                    {
                        if (j != 1)
                        {
                            table[i + 4, temp].AddParagraph().AppendText(rows[j].ToString());
                            temp++;
                        }
                    }
                }
                // 单元格样式
                table[0, 0].CellFormat.FitText = true;
                table[5 + dataTable.Rows.Count, 0].CellFormat.FitText = true;  // 未付款
                // 单元格高度设定
                //for (int i = 0; i < 9 + dataTable.Rows.Count; i++) {
                //    table.Rows[i].Height = 30;
                //}

                for (int i = 3; i < 5 + dataTable.Rows.Count; i++)
                {
                    
                    for (int j = 0; j < 9; j++)
                    {
                        table[i, 1].CellFormat.FitText = true;
                    }
                }
                try
                {
                    //保存文档
                    document.SaveToFile(saveDia.FileName);
                    //doc.SaveToFile(saveDia.FileName, FileFormat.Docx2013);
                    MessageBox.Show("导出成功!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出文件时出错,文件可能正被打开！\n" + ex.Message);
                    return;
                }
            }
        }


        //作者：fanz2000
        //Email:fanz2000@sohu.com
        /// <summary>
        /// 转换数字金额主函数（包括小数）
        /// </summary>
        /// <param name="str">数字字符串</param>
        /// <returns>转换成中文大写后的字符串或者出错信息提示字符串</returns>
        public string ConvertSum(string str)
        {
            if (!IsPositveDecimal(str))
                return "输入的不是正数字！";
            if (Double.Parse(str) > 999999999999.99)
                return "数字太大，无法换算，请输入一万亿元以下的金额";
            char[] ch = new char[1];
            ch[0] = '.'; //小数点
            string[] splitstr = null; //定义按小数点分割后的字符串数组
            splitstr = str.Split(ch[0]);//按小数点分割字符串
            if (splitstr.Length == 1) //只有整数部分
                return ConvertData(str) + "圆整";
            else //有小数部分
            {
                string rstr;
                rstr = ConvertData(splitstr[0]) + "圆";//转换整数部分
                rstr += ConvertXiaoShu(splitstr[1]);//转换小数部分
                return rstr;
            }

        }
        /// <summary>
        /// 判断是否是正数字字符串
        /// </summary>
        /// <param name="str"> 判断字符串</param>
        /// <returns>如果是数字，返回true，否则返回false</returns>
        public bool IsPositveDecimal(string str)
        {
            Decimal d;
            try
            {
                d = Decimal.Parse(str);
            }
            catch (Exception)
            {
                return false;
            }
            if (d > 0)
                return true;
            else
                return false;
        }
        /// <summary>
        /// 转换数字（整数）
        /// </summary>
        /// <param name="str">需要转换的整数数字字符串</param>
        /// <returns>转换成中文大写后的字符串</returns>
        public string ConvertData(string str)
        {
            string tmpstr = "";
            string rstr = "";
            int strlen = str.Length;
            if (strlen <= 4)//数字长度小于四位
            {
                rstr = ConvertDigit(str);
            }
            else
            {

                if (strlen <= 8)//数字长度大于四位，小于八位
                {
                    tmpstr = str.Substring(strlen - 4, 4);//先截取最后四位数字
                    rstr = ConvertDigit(tmpstr);//转换最后四位数字
                    tmpstr = str.Substring(0, strlen - 4);//截取其余数字
                                                          //将两次转换的数字加上萬后相连接
                    rstr = String.Concat(ConvertDigit(tmpstr) + "萬", rstr);
                    rstr = rstr.Replace("零萬", "萬");
                    rstr = rstr.Replace("零零", "零");

                }
                else
                 if (strlen <= 12)//数字长度大于八位，小于十二位
                {
                    tmpstr = str.Substring(strlen - 4, 4);//先截取最后四位数字
                    rstr = ConvertDigit(tmpstr);//转换最后四位数字
                    tmpstr = str.Substring(strlen - 8, 4);//再截取四位数字
                    rstr = String.Concat(ConvertDigit(tmpstr) + "萬", rstr);
                    tmpstr = str.Substring(0, strlen - 8);
                    rstr = String.Concat(ConvertDigit(tmpstr) + "億", rstr);
                    rstr = rstr.Replace("零億", "億");
                    rstr = rstr.Replace("零萬", "零");
                    rstr = rstr.Replace("零零", "零");
                    rstr = rstr.Replace("零零", "零");
                }
            }
            strlen = rstr.Length;
            if (strlen >= 2)
            {
                switch (rstr.Substring(strlen - 2, 2))
                {
                    case "佰零": rstr = rstr.Substring(0, strlen - 2) + "佰"; break;
                    case "仟零": rstr = rstr.Substring(0, strlen - 2) + "仟"; break;
                    case "萬零": rstr = rstr.Substring(0, strlen - 2) + "萬"; break;
                    case "億零": rstr = rstr.Substring(0, strlen - 2) + "億"; break;

                }
            }

            return rstr;
        }
        /// <summary>
        /// 转换数字（小数部分）
        /// </summary>
        /// <param name="str">需要转换的小数部分数字字符串</param>
        /// <returns>转换成中文大写后的字符串</returns>
        public string ConvertXiaoShu(string str)
        {
            int strlen = str.Length;
            string rstr;
            if (strlen == 1)
            {
                rstr = ConvertChinese(str) + "角";
                return rstr;
            }
            else
            {
                string tmpstr = str.Substring(0, 1);
                rstr = ConvertChinese(tmpstr) + "角";
                tmpstr = str.Substring(1, 1);
                rstr += ConvertChinese(tmpstr) + "分";
                rstr = rstr.Replace("零分", "");
                rstr = rstr.Replace("零角", "");
                return rstr;
            }


        }

        /// <summary>
        /// 转换数字
        /// </summary>
        /// <param name="str">转换的字符串（四位以内）</param>
        /// <returns></returns>
        public string ConvertDigit(string str)
        {
            int strlen = str.Length;
            string rstr = "";
            switch (strlen)
            {
                case 1: rstr = ConvertChinese(str); break;
                case 2: rstr = Convert2Digit(str); break;
                case 3: rstr = Convert3Digit(str); break;
                case 4: rstr = Convert4Digit(str); break;
            }
            rstr = rstr.Replace("拾零", "拾");
            strlen = rstr.Length;

            return rstr;
        }


        /// <summary>
        /// 转换四位数字
        /// </summary>
        public string Convert4Digit(string str)
        {
            string str1 = str.Substring(0, 1);
            string str2 = str.Substring(1, 1);
            string str3 = str.Substring(2, 1);
            string str4 = str.Substring(3, 1);
            string rstring = "";
            rstring += ConvertChinese(str1) + "仟";
            rstring += ConvertChinese(str2) + "佰";
            rstring += ConvertChinese(str3) + "拾";
            rstring += ConvertChinese(str4);
            rstring = rstring.Replace("零仟", "零");
            rstring = rstring.Replace("零佰", "零");
            rstring = rstring.Replace("零拾", "零");
            rstring = rstring.Replace("零零", "零");
            rstring = rstring.Replace("零零", "零");
            rstring = rstring.Replace("零零", "零");
            return rstring;
        }
        /// <summary>
        /// 转换三位数字
        /// </summary>
        public string Convert3Digit(string str)
        {
            string str1 = str.Substring(0, 1);
            string str2 = str.Substring(1, 1);
            string str3 = str.Substring(2, 1);
            string rstring = "";
            rstring += ConvertChinese(str1) + "佰";
            rstring += ConvertChinese(str2) + "拾";
            rstring += ConvertChinese(str3);
            rstring = rstring.Replace("零佰", "零");
            rstring = rstring.Replace("零拾", "零");
            rstring = rstring.Replace("零零", "零");
            rstring = rstring.Replace("零零", "零");
            return rstring;
        }
        /// <summary>
        /// 转换二位数字
        /// </summary>
        public string Convert2Digit(string str)
        {
            string str1 = str.Substring(0, 1);
            string str2 = str.Substring(1, 1);
            string rstring = "";
            rstring += ConvertChinese(str1) + "拾";
            rstring += ConvertChinese(str2);
            rstring = rstring.Replace("零拾", "零");
            rstring = rstring.Replace("零零", "零");
            return rstring;
        }
        /// <summary>
        /// 将一位数字转换成中文大写数字
        /// </summary>
        public string ConvertChinese(string str)
        {
            //"零壹贰叁肆伍陆柒捌玖拾佰仟萬億圆整角分"
            string cstr = "";
            switch (str)
            {
                case "0": cstr = "零"; break;
                case "1": cstr = "壹"; break;
                case "2": cstr = "贰"; break;
                case "3": cstr = "叁"; break;
                case "4": cstr = "肆"; break;
                case "5": cstr = "伍"; break;
                case "6": cstr = "陆"; break;
                case "7": cstr = "柒"; break;
                case "8": cstr = "捌"; break;
                case "9": cstr = "玖"; break;
            }
            return (cstr);
        }

    }
}
