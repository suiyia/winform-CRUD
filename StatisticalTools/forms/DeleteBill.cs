using PapersStatisticTools;
using StatisticalTools.po;
using StatisticalTools.utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StatisticalTools.forms
{
    public partial class DeleteBill : Form
    {
        private String billid = null;

        private String xinghao = null;
        private String zhonglei = null;
        private String color = null;

        public DeleteBill()
        {
            InitializeComponent();
            //MysqlManager.SettingConnectStr();
        }

        public DeleteBill(String text)
        {
            InitializeComponent();
            billid = text;
        }

        // 页面加载时  加载查询的数据
        private void DeleteBill_Load(object sender, EventArgs e)
        {
            xinghao = ConfigHelper.GetAppConfig("xinghao");
            zhonglei = ConfigHelper.GetAppConfig("zhonglei");
            color = ConfigHelper.GetAppConfig("color");
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox3.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox4.DropDownStyle = ComboBoxStyle.DropDownList;
            String[] xinghaos = xinghao.Split('.');
            String[] zhongleis = zhonglei.Split('.');
            String[] colors = color.Split('.');
            // 从数据库读取用户
            MysqlManager manager = new MysqlManager();
            DataTable table = manager.getGuestNames();
            String[] guests = new string[table.Rows.Count];
            for (int i = 0; i < table.Rows.Count; i++)
            {
                comboBox1.Items.Add(table.Rows[i]["name"]);
            }

            for (int i = 0; i < xinghaos.Length; i++)
            {
                comboBox2.Items.Add(xinghaos[i]);
            }

            for (int i = 0; i < zhongleis.Length; i++)
            {
                comboBox3.Items.Add(zhongleis[i]);
            }

            for (int i = 0; i < colors.Length; i++)
            {
                comboBox4.Items.Add(colors[i]);
            }

            // 表单填上原来的值
            BillPO bill = manager.getBillPObyBillid(billid);

            textBox1.Text = bill.Billid;
            comboBox1.SelectedIndex = comboBox1.FindString(bill.GuestName);
            comboBox2.SelectedIndex = comboBox2.FindString(bill.Xinghao);
            comboBox3.SelectedIndex = comboBox3.FindString(bill.Zhonglei);
            comboBox4.SelectedIndex = comboBox4.FindString(bill.Color);
            //textBox2.Text = bill.Num.ToString();
            textBox3.Text = bill.Num.ToString();
            textBox4.Text = bill.SinglePrice.ToString();
            textBox5.Text = bill.TotalPrice.ToString();
            textBox9.Text = bill.OtherText;

        }

        // 确认修改
        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("确认删除 ？ ", "提示", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                MysqlManager manager = new MysqlManager();
                manager.DeleteBill(textBox1.Text);
                MessageBox.Show("删除成功 ！");
                this.Close();
            }
        }

    }
}
