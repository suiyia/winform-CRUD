using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows.Forms;
//这个文件用来读写Ini文件中的串口参数
namespace PapersStatisticTools
{
    public class ReadAndWriteIniFile
    {
        public static string  sFileName = System.AppDomain.CurrentDomain.BaseDirectory + "ExeIniConfig.ini";

        public ReadAndWriteIniFile()
        {
            
        }
        //声明读写INI文件的API函数
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        
       /// <summary>
       /// 从配置文件中获取信息
       /// </summary>
       /// <param name="section">根目录</param>
       /// <param name="key">关键字</param>
       /// <param name="def">默认值</param>
       /// <returns></returns>
       public static string GetInfoFromIniFile(string section, string key, string def)
        {
            StringBuilder stringBuilder = new StringBuilder(256);
            if (File.Exists(sFileName))
            {

                //获取软件名字
                GetPrivateProfileString(section, key, def, stringBuilder, 256, sFileName);
                return stringBuilder.ToString();
            }
            else
            {
                LogerHelper.CreateLogTxt("GetInfoFromIniFile", "配置文件不存在");
                return "";
            }
        }

       public static bool SetInfoToIniFile(string section, string key, string value)
       {
           if (File.Exists(sFileName))
           {
               try {
                   WritePrivateProfileString(section, key, value, sFileName);               
               }catch(Exception exp)
               {
                   LogerHelper.CreateLogTxt("SetInfoToIniFile", "配置文件存在,写参数报错:" + section + "--" + key + "--" + value);
                   return false;
               }
               return true;
           }
           else
           {
               LogerHelper.CreateLogTxt("SetInfoToIniFile", "配置文件不存在");
               return false;
           }
       }

       

    }
}