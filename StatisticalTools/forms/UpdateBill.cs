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
    public partial class UpdateBill : Form
    {
        private String billid = null;

        private String xinghao = null;
        private String zhonglei = null;
        private String color = null;

        public UpdateBill()
        {
            InitializeComponent();
            //MysqlManager.SettingConnectStr();
        }

        public UpdateBill(String text)
        {
            InitializeComponent();
            billid = text;
        }

        // 页面加载时  加载查询的数据
        private void UpdateBill_Load(object sender, EventArgs e)
        {
            xinghao = ConfigHelper.GetAppConfig("xinghao");
            zhonglei = ConfigHelper.GetAppConfig("zhonglei");
            color = ConfigHelper.GetAppConfig("color");
            //comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox3.DropDownStyle = ComboBoxStyle.DropDownList;
            //comboBox4.DropDownStyle = ComboBoxStyle.DropDownList;
            String[] xinghaos = xinghao.Split('.');
            String[] zhongleis = zhonglei.Split('.');
            String[] colors = color.Split('.');
            // 从数据库读取用户
            MysqlManager manager = new MysqlManager();
            //DataTable table = manager.getGuestNames();
            //String[] guests = new string[table.Rows.Count];
            //for (int i = 0; i < table.Rows.Count; i++)
            //{
            //    comboBox1.Items.Add(table.Rows[i]["name"]);
            //}

            for (int i = 0; i < xinghaos.Length; i++)
            {
                comboBox2.Items.Add(xinghaos[i]);
            }

            for (int i = 0; i < zhongleis.Length; i++)
            {
                comboBox3.Items.Add(zhongleis[i]);
            }

            //for (int i = 0; i < colors.Length; i++)
            //{
            //    comboBox4.Items.Add(colors[i]);
            //}

            // 表单填上原来的值
            BillPO bill = manager.getBillPObyBillid(billid);

            textBox1.Text = bill.Billid;
            //comboBox1.SelectedIndex = comboBox1.FindString(bill.GuestName);
            comboBox2.SelectedIndex = comboBox2.FindString(bill.Xinghao);
            comboBox3.SelectedIndex = comboBox3.FindString(bill.Zhonglei);
            //comboBox4.SelectedIndex = comboBox4.FindString(bill.Color);

            textBox10.Text = bill.Color.ToString();
            textBox3.Text = bill.Num.ToString();
            textBox4.Text = bill.SinglePrice.ToString();
            textBox5.Text = bill.TotalPrice.ToString();
            textBox9.Text = bill.OtherText;
        }

        // 确认修改
        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("确认修改 ？ ", "提示", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                BillPO billPO = new BillPO();
                billPO.Billid = textBox1.Text;
                //billPO.GuestName = comboBox1.SelectedItem.ToString();
                billPO.Xinghao = comboBox2.SelectedItem.ToString();
                billPO.Zhonglei = comboBox3.SelectedItem.ToString();
                billPO.Color = textBox10.Text;
                billPO.Num = Convert.ToInt32(textBox3.Text);
                billPO.SinglePrice = Convert.ToDouble(textBox4.Text);
                billPO.TotalPrice = Convert.ToDouble(textBox4.Text) * Convert.ToDouble(textBox3.Text);
                billPO.OtherText = textBox9.Text;
                MysqlManager manager = new MysqlManager();
                int result = manager.UpdateBill(billPO);
                if (result == 1) {
                    MessageBox.Show("修改成功 ！");
                }
                else
                {
                    MessageBox.Show("修改失败！");
                }
                this.Close();
            }
        }

        // 取消修改
        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // 数量只能是 整数
        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != '\b')//这是允许输入退格键
            {
                if ((e.KeyChar < '0') || (e.KeyChar > '9'))//这是允许输入0-9数字
                {
                    e.Handled = true;
                }
            }
        }
    }
}
