using Microsoft.Office.Interop.Excel;
using NPinyin;
using PapersStatisticTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace StatisticalTools.utils
{
    class ExportToExcel
    {
        // 客户名  总金额
        public void ExportToExcelFun(System.Data.DataTable dt,String name,double sum)
        {
            MysqlManager mysqlManager = new MysqlManager();
            String phone = mysqlManager.getPhone(name);

            if (dt == null) return;
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("无法创建Excel对象，可能您的电脑未安装Excel");
                return;
            }

            System.Windows.Forms.SaveFileDialog saveDia = new SaveFileDialog();
            saveDia.Filter = "Excel|*.xlsx";
            saveDia.Title = "导出为Excel文件";
            saveDia.FileName = DateTime.Now.ToString("yyyyMMddhhss") + name;
            if (saveDia.ShowDialog() == System.Windows.Forms.DialogResult.OK
             && !string.Empty.Equals(saveDia.FileName))
            {
                Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
                //Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\Template.xlsx");
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];//取得sheet1
                Microsoft.Office.Interop.Excel.Range range = null;

                long totalCount = dt.Rows.Count;
                long rowRead = 0;
                float percent = 0;
                string fileName = saveDia.FileName;

                ////写入标题
                //for (int i = 0; i < dt.Columns.Count; i++)
                //{
                //    worksheet.Cells[2, i + 1] = dt.Columns[i].ColumnName;
                //    range = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[2, i + 1];
                //    //range.Interior.ColorIndex = 15;//背景颜色
                //    range.Font.Bold = true;//粗体
                //    range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;//居中
                //                                                                                       //加边框
                //    range.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, null);

                //    //range.ColumnWidth = 4.63;//设置列宽
                //    //range.EntireColumn.AutoFit();//自动调整列宽
                //    //r1.EntireRow.AutoFit();//自动调整行高
                //}

                //写入内容
                // 客户名称
                worksheet.Cells[2, 2] = name;
                // 电话
                worksheet.Cells[3, 2] = phone;
                // 单号
                Encoding gb2312 = Encoding.GetEncoding("GB2312");
                string s = Pinyin.ConvertEncoding(name, Encoding.UTF8, gb2312);
                worksheet.Cells[2, 8] = s+"-"+DateTime.Now.Year+"-"+ DateTime.Now.Month+"-"+ DateTime.Now.Day;

                // 开票日期
                worksheet.Cells[3, 8] = DateTime.Now.ToString();

                int r = 0;
                for (r = 0; r < dt.DefaultView.Count; r++)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        worksheet.Cells[r + 6, 1] = r + 1;
                        worksheet.Cells[r + 6, i + 2] = dt.DefaultView[r][i];
                        //range = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[r + 6, i + 1];
                        //range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        //range.Font.Size = 12;//字体大小
                        //                    //加边框
                        //range.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, null);
                        //range.EntireColumn.AutoFit();//自动调整列宽
                    }

                    rowRead++;
                    percent = ((float)(100 * rowRead)) / totalCount;
                    System.Windows.Forms.Application.DoEvents();
                }

                range = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[r+6,7];
                range.Font.Size = 15;
                range.Font.Bold = true;
                worksheet.Cells[r + 6, 7] = "总计： ";

              

                range = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[r + 6, 8];
                range.Font.Size = 15;
                range.Font.Bold = true;
                worksheet.Cells[r + 6, 8] = sum + " （元）";

                range = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[r + 8, 8];
                range.Font.Size = 17;
                range.Font.Bold = true;
                worksheet.Cells[r + 8, 8] = "经手人： ";



                range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                if (dt.Columns.Count > 1)
                {
                    range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                }

                try
                {
                    workbook.Saved = true;
                    workbook.SaveCopyAs(fileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出文件时出错,文件可能正被打开！\n" + ex.Message);
                    return;
                }

                workbooks.Close();
                if (xlApp != null)
                {
                    xlApp.Workbooks.Close();
                    xlApp.Quit();
                    int generation = System.GC.GetGeneration(xlApp);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                    xlApp = null;
                    System.GC.Collect(generation);

                }

                GC.Collect();//强行销毁
                #region 强行杀死最近打开的Excel进程

                System.Diagnostics.Process[] excelProc = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                System.DateTime startTime = new DateTime();
                int m, killId = 0;
                for (m = 0; m < excelProc.Length; m++)
                {
                    if (startTime < excelProc[m].StartTime)
                    {
                        startTime = excelProc[m].StartTime;
                        killId = m;
                    }
                }
                if (excelProc[killId].HasExited == false)
                {
                    excelProc[killId].Kill();
                }

                #endregion
                MessageBox.Show("导出成功!");
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
