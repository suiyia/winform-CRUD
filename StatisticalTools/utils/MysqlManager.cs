using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using System.Data;
using StatisticalTools.po;
using StatisticalTools;
using System.Windows.Forms;
using StatisticalTools.utils;

namespace PapersStatisticTools
{
    /*
    * 数据库操作类-- 业务逻辑层
    */
    class MysqlManager
    {
        /// <summary>
        /// 初始化数据库中数据访问层的连接字符串信息
        /// </summary>
        public static void SettingConnectStr()
        {
            //String Server = ReadAndWriteIniFile.GetInfoFromIniFile("Run", "Server", "");
            //String Database = ReadAndWriteIniFile.GetInfoFromIniFile("Run", "Database", "");
            //String Uid = ReadAndWriteIniFile.GetInfoFromIniFile("Run", "Uid", "");
            //String Pwd = ReadAndWriteIniFile.GetInfoFromIniFile("Run", "Pwd", "");
            //String Charset = ReadAndWriteIniFile.GetInfoFromIniFile("Run", "Charset", "");
            String conn = ConfigHelper.GetAppConfig("connectionstring");

            //string connectStr = "Server = 127.0.0.1;Database=ssm;Uid=root;Pwd=123456;Charset=utf8;";
            //string connectStr = "Server=" + Server.Trim() + ";Database=" + Database.Trim() + ";Uid=" + Uid.Trim() + ";Pwd=" + Pwd.Trim() + ";Charset=" + Charset.Trim();

            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            //指定数据库连接字符串信息
            mysqlHelper.ConnStr = conn;
        }

        /// <summary>
        /// 获取所有的订单列表(从指定页面开始)
        /// </summary>
        public DataTable getOrderInfo(int startIndex)
        {
            String sqlStr = "SELECT billid as '单据编号', guestName as '顾客名称',xinghao as '型号',zhonglei as '种类',color as '颜色',num as '数量',singlePrice as '单价',totalPrice as '总金额',kaipiaoDate as '开票日期',otherText as '备注' from orderlist order by id desc LIMIT " + startIndex + ", 50 ;";
            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            return mysqlHelper.ExcuteDataTable(sqlStr);
        }

        /// <summary>
        /// 获取订单的总数（计算：共XXX页）
        /// </summary>
        public int getOrderCount()
        {
            String sqlStr = "SELECT count(*) as numbers from orderlist;";

            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            DataTable dt = mysqlHelper.ExcuteDataTable(sqlStr);
            if (dt.Rows.Count == 0)
            {
                return 0;
            }

            DataRow row = dt.Rows[0];
            object oj = row["numbers"];

            return Convert.ToInt32(oj);
        }

        # region  
        /// 插入新客户  添加
        public int insertGuest(Guest guest)
        {
            String sqlStr = "INSERT INTO guest(name,phone) VALUES (?name,?phone);";
            MySqlParameter[] parameters = {
                new MySqlParameter("?name",MySqlDbType.String),
                new MySqlParameter("?phone", MySqlDbType.String)
            };
            parameters[0].Value = guest.Name;
            parameters[1].Value = guest.Tel;
            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            return mysqlHelper.ExecuteNonquery(sqlStr, parameters);
        }

        // 查看用户存不存在
        public Boolean checkGuest(String name) {
            String sql = "select count(*) as count1 from guest where name = '" + name + "'";
            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            int count =  Convert.ToInt32(mysqlHelper.ExcuteDataTable(sql).Rows[0]["count1"]);
            if (count != 0)
            {
                return false;
            }
            else {
                return true;
            }
        }

        // 根据用户姓名获取电话号码
        public String getPhone(String name)
        {
            String sqlStr = "select phone from guest where name = '" + name + "'";
            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            return mysqlHelper.ExcuteDataTable(sqlStr).Rows[0]["phone"].ToString();
        }

        // 根据用户姓名获取他下的单的数量
        public int getGuestOrderCount(String name) {
            String sqlStr = "select count(*) as count1 from orderlist where guestName = '" + name + "'";
            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            int count = Convert.ToInt32(mysqlHelper.ExcuteDataTable(sqlStr).Rows[0]["count1"].ToString());
            if (count == 0) {
                return 1;
            }
            else
            {
                return count;
            }
        }

        // 删除客户
        public int deleteGuest(Guest guest) {
            String sqlStr = "delete from guest where name = ?guestName";
            MySqlParameter[] parameters = {
                new MySqlParameter("?guestName",MySqlDbType.String),
            };
            parameters[0].Value = guest.Name;
            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            return mysqlHelper.ExecuteNonquery(sqlStr, parameters);
        }
        #endregion

        //添加订单
        public int insertBill(BillPO billPO)
        {
            String sqlStr = "INSERT INTO orderlist";
            sqlStr += "(billid,guestName,xinghao,zhonglei,color,danwei,num,singlePrice,totalPrice,";
            sqlStr += "kaipiaor,jinshour,picPath,kaipiaoDate,address,otherText) VALUES ";
            sqlStr += "(?billid,?guestName,?xinghao,?zhonglei,?color,?danwei,?num,?singlePrice,?totalPrice,";
            sqlStr += "?kaipiaor,?jinshour,?picPath,?kaipiaoDate,?address,?otherText)";
            MySqlParameter[] parameters = {
                new MySqlParameter("?billid",MySqlDbType.String),
                new MySqlParameter("?guestName", MySqlDbType.String),
                new MySqlParameter("?xinghao", MySqlDbType.String),
                new MySqlParameter("?zhonglei", MySqlDbType.String),
                new MySqlParameter("?color", MySqlDbType.String),
                new MySqlParameter("?danwei", MySqlDbType.String),
                new MySqlParameter("?num", MySqlDbType.Int32),
                new MySqlParameter("?singlePrice", MySqlDbType.Double),
                new MySqlParameter("?totalPrice", MySqlDbType.Double),
                new MySqlParameter("?kaipiaor", MySqlDbType.String),
                new MySqlParameter("?jinshour", MySqlDbType.String),
                new MySqlParameter("?picPath", MySqlDbType.String),
                new MySqlParameter("?kaipiaoDate", MySqlDbType.String),
                new MySqlParameter("?address", MySqlDbType.String),
                new MySqlParameter("?otherText", MySqlDbType.String),
            };
            parameters[0].Value = billPO.Billid;
            parameters[1].Value = billPO.GuestName;
            parameters[2].Value = billPO.Xinghao;
            parameters[3].Value = billPO.Zhonglei;
            parameters[4].Value = billPO.Color;
            parameters[5].Value = billPO.Danwei;
            parameters[6].Value = billPO.Num;
            parameters[7].Value = billPO.SinglePrice;
            parameters[8].Value = billPO.TotalPrice;
            parameters[9].Value = billPO.Kaipiaor;
            parameters[10].Value = billPO.Jinshour;
            parameters[11].Value = billPO.PicPath;
            parameters[12].Value = billPO.KaipiaoDate;
            parameters[13].Value = billPO.Address;
            parameters[14].Value = billPO.OtherText;
            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            return mysqlHelper.ExecuteNonquery(sqlStr, parameters);
        }

        /// 插入新颜色   没用
        public int insertColor(Color color)
        {
            String sqlStr = "insert into color(cid,color) values(?cid,?color);";
            MySqlParameter[] parameters = {
                new MySqlParameter("?cid",MySqlDbType.String),
                new MySqlParameter("?color", MySqlDbType.String)
            };
            parameters[0].Value = color.Cid;
            parameters[1].Value = color.Mcolor;
            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            return mysqlHelper.ExecuteNonquery(sqlStr, parameters);
        }
        /// 插入新型号  没用
        public int insertXinghao(Xinghao xinghao)
        {
            String sqlStr = "insert into xinghao(kind) values(?kind);";
            MySqlParameter[] parameters = {
                new MySqlParameter("?kind",MySqlDbType.String),
            };
            parameters[0].Value = xinghao.Kind;
            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            return mysqlHelper.ExecuteNonquery(sqlStr, parameters);
        }
        /// 插入新种类  没用
        public int insertZhonglei(Zhonglei zhonglei)
        {
            String sqlStr = "insert into zhonglei(color) values(?color);";
            MySqlParameter[] parameters = {
                new MySqlParameter("?color",MySqlDbType.String),
            };
            parameters[0].Value = zhonglei.Color;
            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            return mysqlHelper.ExecuteNonquery(sqlStr, parameters);
        }
        //查找数据库所有型号  // 没用
        public DataTable getXinghao()
        {
            String sqlStr = "SELECT * from xinghao;";
            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            return mysqlHelper.ExcuteDataTable(sqlStr);
        }
        //查找数据库所有种类  // 没用 
        public DataTable getZhonglei()
        {
            String sqlStr = "SELECT * from zhonglei;";
            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            return mysqlHelper.ExcuteDataTable(sqlStr);
        }

        //获取客户名称 combox
        public DataTable getGuestNames()
        {
            String sqlStr = "select name from guest ORDER BY CONVERT(name USING gbk)";
            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            return mysqlHelper.ExcuteDataTable(sqlStr);
        }

        // 条件查询所有订单
        public DataTable getOrderlist(String sql)
        {
            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            return mysqlHelper.ExcuteDataTable(sql);
        }

        // 根据订单编号查询订单，用于修改订单页面
        public BillPO getBillPObyBillid(String billid) {
            BillPO billPO = new BillPO();
            String sqlStr = "SELECT id as '序号', billid as '单据编号', guestName as '顾客名称',xinghao as '型号',zhonglei as '种类',color as '颜色',danwei as '单位',num as '数量',singlePrice as '单价',totalPrice as '总金额',kaipiaor as '开票人',jinshour as '经手人',picPath as '图片地址',kaipiaoDate as '开票日期',address as '厂址',otherText as '备注' from orderlist where billid='" + billid + "'";
            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            DataTable table = mysqlHelper.ExcuteDataTable(sqlStr);
            // 根据订单没查到 或者 多条相同的订单编号
            if (table.Rows.Count !=  1)
            {
                return null;
            }
            else {
                billPO.Billid = table.Rows[0]["单据编号"].ToString();
                billPO.GuestName = table.Rows[0]["顾客名称"].ToString();
                billPO.Xinghao = table.Rows[0]["型号"].ToString();
                billPO.Zhonglei = table.Rows[0]["种类"].ToString();
                billPO.Color = table.Rows[0]["颜色"].ToString();
                billPO.Danwei = table.Rows[0]["单位"].ToString();
                billPO.Num = Convert.ToInt32(table.Rows[0]["数量"].ToString());
                billPO.SinglePrice = Convert.ToDouble(table.Rows[0]["单价"].ToString());
                billPO.TotalPrice = Convert.ToDouble(table.Rows[0]["总金额"].ToString());
                billPO.Kaipiaor = table.Rows[0]["开票人"].ToString();
                billPO.OtherText = table.Rows[0]["备注"].ToString();
                return billPO;
            }
        }
        
        
        // 更新订单,传入 订单对象
        public int UpdateBill(BillPO billPO) {
            String sql = "update orderlist set xinghao=?xinghao,zhonglei=?zhonglei,color=?color,num=?num,singlePrice=?singlePrice,totalPrice=?totalPrice,otherText=?otherText where billid=?billid";
            MySqlParameter[] parameters = {
                new MySqlParameter("?xinghao",MySqlDbType.String),
                new MySqlParameter("?zhonglei",MySqlDbType.String),
                new MySqlParameter("?color",MySqlDbType.String),
                new MySqlParameter("?num",MySqlDbType.Int32),
                new MySqlParameter("?singlePrice",MySqlDbType.Double),
                new MySqlParameter("?totalPrice",MySqlDbType.Double),
                new MySqlParameter("?otherText",MySqlDbType.String),
                new MySqlParameter("?billid",MySqlDbType.String)
                
            };
            parameters[0].Value = billPO.Xinghao;
            parameters[1].Value = billPO.Zhonglei;
            parameters[2].Value = billPO.Color;
            parameters[3].Value = billPO.Num;
            parameters[4].Value = billPO.SinglePrice;
            parameters[5].Value = billPO.TotalPrice;
            parameters[6].Value = billPO.OtherText;
            parameters[7].Value = billPO.Billid;
            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            return mysqlHelper.ExecuteNonquery(sql, parameters);
        }
       

        // 根据订单 id 删除订单
        public int DeleteBill(String billid) {
            String sql = "delete from orderlist where billid='" + billid+"'";
            MySqlHelper mysqlHelper = MySqlHelper.Ins;
            return mysqlHelper.ExecuteNonquery(sql);
        }


    }
}