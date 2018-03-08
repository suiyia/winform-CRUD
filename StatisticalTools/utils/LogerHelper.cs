using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Globalization;


namespace PapersStatisticTools
{
    /// <summary>
    /// 日志文件工具类
    /// </summary>
    public class LogerHelper
    {
        #region  
        /// <summary>
        /// 向日志文件写记录
        /// </summary>
        /// <param name="AEICode"></param>
        /// <param name="message"></param>
        public static void CreateLogTxt(string AEICode, string message)
        {
            string strPath;                                                   //文件的路径
            DateTime dt = DateTime.Now;
            try
            {
                strPath = System.AppDomain.CurrentDomain.BaseDirectory + "Log";          //winform工程\bin\目录下 创建日志文件夹 
                if (Directory.Exists(strPath) == false)                          //工程目录下 Log目录 '目录是否存在,为true则没有此目录
                {
                    Directory.CreateDirectory(strPath);                       //建立目录　Directory为目录对象
                }
                //年目录
                strPath = strPath + "\\" + dt.ToString("yyyy");
                if (Directory.Exists(strPath) == false)
                {
                    Directory.CreateDirectory(strPath);
                }
                //文件命名方式：年月日
                strPath = strPath + "\\" + dt.ToString("yyyyMMdd") + ".txt";
                StreamWriter FileWriter = new StreamWriter(strPath, true);           //创建日志文件
                FileWriter.WriteLine(AEICode + "->[" + dt.ToString("yyyy-MM-dd HH:mm:ss") + "]  " + message);
                FileWriter.Close();
            }
            catch (Exception ex)
            {
                string str = ex.Message.ToString();
                //写入日志失败
            }
        }
        
        /// <summary>
        /// 向日志文件写记录
        /// </summary>
        /// <param name="message"></param>
        public static void CreateLogTxt(string message)
        {
            string strPath;                                                   //文件的路径
            DateTime dt = DateTime.Now;
            try
            {
                strPath = System.AppDomain.CurrentDomain.BaseDirectory + "Log";          //winform工程\bin\目录下 创建日志文件夹 
                if (Directory.Exists(strPath) == false)                          //工程目录下 Log目录 '目录是否存在,为true则没有此目录
                {
                    Directory.CreateDirectory(strPath);                       //建立目录　Directory为目录对象
                }
                strPath = strPath + "\\" + dt.ToString("yyyy");
                if (Directory.Exists(strPath) == false)
                {
                    Directory.CreateDirectory(strPath);
                }
                strPath = strPath + "\\" + dt.ToString("yyyyMMdd") + ".txt";
                StreamWriter FileWriter = new StreamWriter(strPath, true);           //创建日志文件
                FileWriter.WriteLine("[" + dt.ToString("yyyy-MM-dd HH:mm:ss") + "]  " + message);
                FileWriter.Close();                                                 //关闭StreamWriter对象
            }
            catch (Exception ex)
            {
                string str = ex.Message.ToString();
                //写入日志失败
            }
        }

        /// <summary>
        /// 将curtime之前的日志文件删除
        /// </summary>
        /// <param name="curtime"></param>
        public static void DeleteLogText(DateTime curtime)
        {
            string sPath = System.AppDomain.CurrentDomain.BaseDirectory + "Log";
            //年目录
            DateTime dt_now = DateTime.Now;
            sPath = sPath + "\\" + dt_now.ToString("yyyy");
            //
            try
            {
                if (Directory.Exists(sPath))
                {
                    DirectoryInfo dir = new DirectoryInfo(sPath);
                    FileInfo[] files = dir.GetFiles("*.txt");
                    foreach (FileInfo item in files)
                    {
                        string dateStr = item.Name;
                        dateStr = dateStr.Substring(0, dateStr.IndexOf(".txt"));
                        DateTime dt = DateTime.ParseExact(dateStr, "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture);

                        if (dt.CompareTo(curtime) < 0)
                        {
                            item.Delete();
                        }
                    }
                }
            }
            catch (Exception exp)
            {
                return;
            }
        }
        
        
        //iTimerGap为天数
        /// <summary>
        /// 删除某天之前的所有日志文件
        /// </summary>
        /// <param name="iTimerGap"></param>
        public static void DeleteLogTxt(int iTimerGap)
        {
            DateTime CurTime = DateTime.Now;
            TimeSpan DelGap = new TimeSpan(iTimerGap, 0, 0, 0);
            DateTime PreDateTime = CurTime - DelGap;

            int iPreYear = 0;
            string sTemp = "";
            string sYearPath = "";
            string sPath = System.AppDomain.CurrentDomain.BaseDirectory + "Log";
            System.IO.DirectoryInfo dir = new DirectoryInfo(sPath);

            //防止删除日志文件时出错
            try
            {
                if (Directory.Exists(sPath))
                {
                    string[] fileList = Directory.GetFileSystemEntries(sPath);
                    foreach (string item in fileList)
                    {
                        string sFolder = item;
                        //这个文件下的文件直接删除
                        if (File.Exists(sFolder))
                        {
                            File.Delete(sFolder);
                        }
                        else if (Directory.Exists(sFolder))
                        {
                            sTemp = sFolder.Substring(sFolder.LastIndexOf('\\') + 1, sFolder.Length - sFolder.LastIndexOf('\\') - 1);
                            try
                            {
                                iPreYear = Convert.ToInt32(sTemp);
                                if (iPreYear < PreDateTime.Year)
                                {
                                    //这种情况下直接删掉文件夹
                                    DeleteFolder(sFolder);
                                }
                                else if (iPreYear == PreDateTime.Year)
                                {
                                    sYearPath = sFolder;
                                    string[] fileYearList = Directory.GetFiles(sYearPath);
                                    foreach (string sFileTmp in fileYearList)
                                    {
                                        if (Directory.Exists(sFileTmp))
                                        {
                                            //文件夹直接删除
                                            try
                                            {
                                                DeleteFolder(sFileTmp);
                                            }
                                            catch (System.Exception ex)
                                            {
                                                continue;
                                            }

                                        }
                                        else if (File.Exists(sFileTmp))
                                        {
                                            string FileName = Path.GetFileName(sFileTmp);
                                            string sDateTime = FileName.Substring(0, FileName.IndexOf('.'));
                                            int iYear = 0;
                                            int iMonth = 0;
                                            int iDay = 0;
                                            if (sDateTime.Length >= 8)
                                            {
                                                try
                                                {
                                                    iYear = Convert.ToInt32(sDateTime.Substring(0, 4));
                                                    iMonth = Convert.ToInt32(sDateTime.Substring(4, 2));
                                                    iDay = Convert.ToInt32(sDateTime.Substring(6, 2));
                                                    //获取文件的名字
                                                    DateTime dtTmp = new DateTime(iYear, iMonth, iDay);
                                                    if (dtTmp < PreDateTime)
                                                    {
                                                        File.Delete(sFileTmp);
                                                    }
                                                }
                                                catch (System.Exception ex)
                                                {
                                                    File.Delete(sFileTmp);
                                                }
                                            }
                                            else
                                            {
                                                File.Delete(sFileTmp);
                                            }


                                        }
                                        else
                                        {

                                        }
                                    }
                                }
                                else
                                {

                                }

                            }
                            catch (System.Exception ex)
                            {
                                //不是以时间命名的文件夹直接删掉
                                DeleteFolder(sFolder);
                            }
                        }
                        else
                        {

                        }

                    }
                }
            }
            catch (System.Exception ex)
            {
                return;
            }
        }
        /// <summary>
        /// 删除文件夹
        /// </summary>
        /// <param name="sPath"></param>
        public static void DeleteFolder(string sPath)
        {
            if (Directory.Exists(sPath)) //如果存在这个文件夹删除之 
            {
                foreach (string d in Directory.GetFileSystemEntries(sPath))
                {
                    if (File.Exists(d))
                        File.Delete(d); //直接删除其中的文件 
                    else
                        DeleteFolder(d); //递归删除子文件夹 
                }
                Directory.Delete(sPath); //删除已空文件夹 
            }
            else
            {
                //Response.Write(dir + " 该文件夹不存在"); //如果文件夹不存在则提示
            }
        }

        #endregion
    }
}
