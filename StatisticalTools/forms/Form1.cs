using MaterialSkin;
using MaterialSkin.Controls;
using PapersStatisticTools;
using StatisticalTools.forms;
using StatisticalTools.utils;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace StatisticalTools
{
    public partial class Form1 : MaterialForm
    {
        private readonly MaterialSkinManager materialSkinManager;

        // 页面分页用
        private int tbm_nowDataMaxIndex = 0;
        private int papersCount = 0;
        private int currIndex = 0;
        private int num = 0;

        // 条件查询中的三个 combox
        private String xinghao = null;
        private String zhonglei = null;
        private String color = null;
        private String guestname = null;

        private DataTable dt = null;

        AutoSizeFormClass asc = new AutoSizeFormClass();

        // 刚开始选择的客户 默认为 空
        public int selectindex = -1;
        public String preColor = "900#灰";

        public Form1()
        {
            InitializeComponent();
            // Initialize MaterialSkinManager
            materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.LightBlue900, Primary.LightBlue900, Primary.Blue900, Accent.Blue700, TextShade.WHITE);
            MysqlManager.SettingConnectStr();
            flushDGW();
            init();
            
            asc.controllInitializeSize(this);

        }

        //刷新文献管理中表格
        private void flushDGW()
        {
            currIndex = 1;
            MysqlManager mysqlManager = new MysqlManager();
            dt = mysqlManager.getOrderInfo(0);  // index = 0~20
            tbm_nowDataMaxIndex = dt.Rows.Count;

            papersCount = mysqlManager.getOrderCount();
            num = papersCount % 50 == 0 ? papersCount / 50 : (papersCount / 50 + 1);

            labelcount.Text = "第 "+ currIndex +" / "+ num + " 页";

            dataGridView1.DataSource = dt;
            dataGridView1.Update();
            btnPre.Enabled = false;
            if (tbm_nowDataMaxIndex >= papersCount)
            {
                btnNext.Enabled = false;
            }
        }

        // 填充条件查询的数据
        public void init() {
            comboBoxUserName.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxXinghao.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxZhonglei.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox1color.DropDownStyle = ComboBoxStyle.DropDownList;
            xinghao = ConfigHelper.GetAppConfig("xinghao");
            zhonglei = ConfigHelper.GetAppConfig("zhonglei");
            color = ConfigHelper.GetAppConfig("color");
            // 从数据库读取用户
            MysqlManager manager = new MysqlManager();
            DataTable table = manager.getGuestNames();
            String[] guests = new string[table.Rows.Count];
            comboBoxUserName.Items.Add("所有客户");
            comboBoxXinghao.Items.Add("所有型号");
            comboBoxZhonglei.Items.Add("所有种类");
            comboBox1color.Items.Add("所有颜色");

            comboBoxUserName.SelectedIndex = 0;
            comboBoxXinghao.SelectedIndex = 0;
            comboBoxZhonglei.SelectedIndex = 0;
            comboBox1color.SelectedIndex = 0;

            for (int i = 0; i < table.Rows.Count; i++)
            {
                comboBoxUserName.Items.Add(table.Rows[i]["name"]);
            }


            String[] xinghaos = xinghao.Split('.');
            String[] zhongleis = zhonglei.Split('.');
            String[] colors = color.Split('.');
            for (int i = 0; i < xinghaos.Length; i++)
            {
                comboBoxXinghao.Items.Add(xinghaos[i]);
            }

            for (int i = 0; i < zhongleis.Length; i++)
            {
                comboBoxZhonglei.Items.Add(zhongleis[i]);
            }

            for (int i = 0; i < colors.Length; i++)
            {
                comboBox1color.Items.Add(colors[i]);
            }
        }

        // 添加顾客
        private void btn2NewGuest_Click(object sender, EventArgs e)
        {
            NewGuest newGuest = new NewGuest();
            newGuest.StartPosition = FormStartPosition.CenterParent;
            newGuest.ShowDialog();
        }

        //添加种类
        private void btn2Newzhonglei_Click(object sender, EventArgs e)
        {
            //Newzhonglei zhonglei = new Newzhonglei();
            //zhonglei.StartPosition = FormStartPosition.CenterParent;
            //zhonglei.ShowDialog();
        }

        //添加型号
        private void btn2Newxinghao_Click(object sender, EventArgs e)
        {
            //Newxinghao xinghao = new Newxinghao();
            //xinghao.StartPosition = FormStartPosition.CenterParent;
            //xinghao.ShowDialog();
        }

        //添加颜色
        private void btn2NewColor_Click(object sender, EventArgs e)
        {
            //Newcolor color = new Newcolor();
            //color.StartPosition = FormStartPosition.CenterParent;
            //color.ShowDialog();
        }


        //                     选择客户 第一步
        private void materialRaisedButton10_Click(object sender, EventArgs e)
        {
            SelectGuest selectGuest = new SelectGuest(this);
            selectGuest.StartPosition = FormStartPosition.CenterParent;
            selectGuest.ShowDialog();
        }

        //                       添加单据  
        private void materialRaisedButton1_Click(object sender, EventArgs e)
        {
            NewBill bill = new NewBill(this);
            bill.StartPosition = FormStartPosition.CenterParent;
            bill.ShowDialog();
            
        }

        // 还原按钮
        private void btnIndex_Click(object sender, EventArgs e)
        {
            flushDGW();
        }

        // 上一页按钮
        private void btnPre_Click(object sender, EventArgs e)
        {
            MysqlManager mysqlManager = new MysqlManager();
            int currentCount = (tbm_nowDataMaxIndex % 50 == 0 ? 50 : tbm_nowDataMaxIndex % 50);
            DataTable dt = mysqlManager.getOrderInfo(tbm_nowDataMaxIndex - (50 + currentCount));
            tbm_nowDataMaxIndex -= currentCount;
            dataGridView1.DataSource = dt;
            if (tbm_nowDataMaxIndex <= 50)
            {
                btnPre.Enabled = false;
            }
            else {
                currIndex--;
            }
            btnNext.Enabled = true;
            labelcount.Text = "第 " + currIndex + " / " + num + " 页";
        }

        // 下一页按钮
        private void btnNext_Click(object sender, EventArgs e)
        {
            MysqlManager mysqlManager = new MysqlManager();
            DataTable dt = mysqlManager.getOrderInfo(tbm_nowDataMaxIndex);
            tbm_nowDataMaxIndex += dt.Rows.Count;
            dataGridView1.DataSource = dt;
            if (tbm_nowDataMaxIndex <= 20)
            {
                btnNext.Enabled = false;
            }
            else {
                
                currIndex++;
            }
            btnPre.Enabled = true;
            if (currIndex > num) {
                currIndex = num;
            }
            labelcount.Text = "第 " + currIndex + " / " + num + " 页";
        }

        // 条件查询
        private void materialRaisedButton2_Click(object sender, EventArgs e)
        {
            String str = null;
            MysqlManager mysqlManager = new MysqlManager();
            String zhonglei = comboBoxZhonglei.SelectedItem.ToString();
            String xinghao = comboBoxXinghao.SelectedItem.ToString();
            String color = comboBox1color.SelectedItem.ToString();
            String guestName = comboBoxUserName.SelectedItem.ToString();
            String starttime = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            //String s = dateTimePicker1.Value.ToShortDateString();
            String endtime = dateTimePicker2.Value.ToString("yyyy-MM-dd");
            str = "SELECT billid as '单据编号', guestName as '顾客名称',xinghao as '型号',zhonglei as '种类',color as '颜色',num as '数量',singlePrice as '单价',totalPrice as '总金额',kaipiaoDate as '开票日期',otherText as '备注' from orderlist where ";
            str += "DATE_FORMAT( kaipiaoDate, '%Y-%m-%d') >= '" + starttime + "' and DATE_FORMAT( kaipiaoDate, '%Y-%m-%d') <= '" + endtime +"'";
            if (zhonglei != "所有种类") {
                str += "and zhonglei='" + zhonglei+"'";
            }
            if (xinghao != "所有型号")
            {
                str += "and xinghao='" + xinghao + "'";
            }
            if (color != "所有颜色")
            {
                str += "and color='" + color + "'";
            }
            if (guestName != "所有客户")
            {
                str += "and guestName='" + guestName + "'";
            }
            str += ";";
            dt = mysqlManager.getOrderlist(str);
            double temp = 0.0;
            int num_temp = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                temp += Double.Parse(dt.Rows[i]["总金额"].ToString());
                num_temp += Int32.Parse(dt.Rows[i]["数量"].ToString());
            }


            dataGridView1.DataSource = dt;
            dataGridView1.Update();

            textBox3.Text = guestName;
            textBox4.Text = zhonglei;
            textBox5.Text = xinghao;
            textBox6.Text = color;
            textBox7.Text = starttime;
            textBox8.Text = endtime;
            textBox9.Text = dt.Rows.Count.ToString();
            textBox10.Text = temp.ToString();
            textBox11.Text = num_temp.ToString();
        }

        // 查询页面 导出按钮
        private void materialRaisedButton3_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("是否导出 " + dt.Rows.Count + " 条数据 ?","是否导出",MessageBoxButtons.YesNo);
            ;
            if (dr == DialogResult.Yes)
            {
                ExportToExcelMain export = new ExportToExcelMain();
                export.ExportToExcelFun(dt);
            }
        }

        // 修改页面   订单查询按钮
        private void materialRaisedButton4_Click_1(object sender, EventArgs e)
        {
            MysqlManager manager = new MysqlManager();
            if (textBox1.Text == "")
            {
                MessageBox.Show("请输入订单编号!");
            }
            else {
                BillPO bill = manager.getBillPObyBillid(textBox1.Text);
                if (bill == null) {
                    MessageBox.Show("订单编号输入有误，请重新输入！");
                }
                else {
                    UpdateBill updateBill = new UpdateBill(textBox1.Text);
                    updateBill.StartPosition = FormStartPosition.CenterParent;
                    updateBill.ShowDialog();
                }
            }
        }

        //  删除页面  订单查询按钮
        private void materialRaisedButton5_Click(object sender, EventArgs e)
        {
            MysqlManager manager = new MysqlManager();
            if (textBox2.Text == "")
            {
                MessageBox.Show("请输入订单编号!");
            }
            else
            {
                BillPO bill = manager.getBillPObyBillid(textBox2.Text);
                if (bill == null)
                {
                    MessageBox.Show("订单编号输入有误，请重新输入！");
                }
                else
                {
                    DeleteBill deleteBill = new DeleteBill(textBox2.Text);
                    deleteBill.StartPosition = FormStartPosition.CenterParent;
                    deleteBill.ShowDialog();
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //label18.Text = System.DateTime.Now.ToString();
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            asc.controlAutoSize(this);
        }

        // 删除客户
        private void materialRaisedButton6_Click(object sender, EventArgs e)
        {
            DeleteGuest deleteGuest = new DeleteGuest();
            deleteGuest.StartPosition = FormStartPosition.CenterParent;
            deleteGuest.ShowDialog();
        }

        // 生成单据表格---订单打印，还没有存入数据库
        private void materialRaisedButton7_Click(object sender, EventArgs e)
        {
            DataTable dt = GetDgvToTable(dataGridView2);
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("请先添加数据");
            }
            else {
                double sum = 0.0;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sum += Convert.ToDouble(dt.Rows[i]["Column8"]);
                }
                DialogResult dr = MessageBox.Show("是否导出 " + dt.Rows.Count + " 条数据 ?", "是否导出", MessageBoxButtons.YesNo);
                ;
                if (dr == DialogResult.Yes)
                {
                    //ExportToExcel export = new ExportToExcel();
                    //export.ExportToExcelFun(dt, dt.Rows[0]["Column2"].ToString(), sum);
                    ExportToWord word = new ExportToWord();
                    word.ExportToWordFun(dt, dt.Rows[0]["Column2"].ToString(), sum);
                    //utils.Test t = new utils.Test();
                    //t.fun();
                    //orderprint o = new orderprint();
                    //Bitmap b = new Bitmap(o.Bounds.Width, o.Bounds.Height);
                    //this.DrawToBitmap(b, new Rectangle(0, 0, o.Width, o.Height));
                    //b.Save("C:\\Users\\answer\\Desktop\\a.jpg");
                }
            }
        }

        // 从datagridview插入订单到数据库
        private void materialRaisedButton8_Click(object sender, EventArgs e)
        {
            MysqlManager manager = new MysqlManager();
            DialogResult dr = MessageBox.Show("确认添加订单到数据库 ？ ", "提示", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                DataTable dt = GetDgvToTable(dataGridView2);
                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("没有数据可添加");
                }
                else {
                    for (int i = 0; i < dt.Rows.Count; i++)
                {
                    BillPO billPO = new BillPO();
                    billPO.Billid = dt.Rows[i]["Column1"].ToString();
                    billPO.GuestName = dt.Rows[i]["Column2"].ToString();
                    billPO.Xinghao = dt.Rows[i]["Column3"].ToString();
                    billPO.Zhonglei = dt.Rows[i]["Column4"].ToString();
                    billPO.Color = dt.Rows[i]["Column5"].ToString();
                    billPO.Num = Convert.ToInt32(dt.Rows[i]["Column6"]);
                    billPO.SinglePrice = Convert.ToDouble(dt.Rows[i]["Column7"]);
                    billPO.TotalPrice = Convert.ToDouble(dt.Rows[i]["Column8"]);
                    billPO.OtherText = dt.Rows[i]["Column9"].ToString();

                    billPO.KaipiaoDate = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString();
                    billPO.Danwei = "码 / 米";
                    billPO.Kaipiaor = "硕麒拉链码庄";
                    billPO.Jinshour = "无";
                    billPO.PicPath = "无";
                    billPO.Address = "无";
                    manager.insertBill(billPO);
                }
                MessageBox.Show("订单添加成功，请返回首页查看 ！");
                }
            }
        }

        public DataTable GetDgvToTable(DataGridView dgv)
        {
            DataTable dt = new DataTable();

            // 列强制转换
            for (int count = 0; count < dgv.Columns.Count; count++)
            {
                DataColumn dc = new DataColumn(dgv.Columns[count].Name.ToString());
                dt.Columns.Add(dc);
            }

            // 循环行
            for (int count = 0; count < dgv.Rows.Count; count++)
            {
                DataRow dr = dt.NewRow();
                for (int countsub = 0; countsub < dgv.Columns.Count; countsub++)
                {
                    dr[countsub] = Convert.ToString(dgv.Rows[count].Cells[countsub].Value);
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        // 情况 data grid view
        private void materialRaisedButton9_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
        }

    }
}
