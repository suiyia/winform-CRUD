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
    public partial class NewBill : Form
    {
        private String xinghao = null;
        private String zhonglei = null;
        private String color = null;
        private Form1 form1;
        private BillPO bill = null;
        //private int index1 = -1;

        public NewBill(Form1 f)
        {
            InitializeComponent();
            //MysqlManager.SettingConnectStr();
            form1 = f;
        }

        public NewBill(Form1 f,int index)
        {
            InitializeComponent();
            //MysqlManager.SettingConnectStr();
            form1 = f;
        }

        // 加载 config 文件内容
        private void NewBill_Load(object sender, EventArgs e)
        {
            xinghao = ConfigHelper.GetAppConfig("xinghao");
            zhonglei = ConfigHelper.GetAppConfig("zhonglei");
            color = ConfigHelper.GetAppConfig("color");
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox3.DropDownStyle = ComboBoxStyle.DropDownList;
            textBox11.Text = form1.preColor;

            // 从数据库读取用户
            MysqlManager manager = new MysqlManager();
            DataTable table = manager.getGuestNames();
            String[] guests = new string[table.Rows.Count];
            for (int i = 0; i < table.Rows.Count; i++) {
                comboBox1.Items.Add(table.Rows[i]["name"]);
            }

            if (form1.selectindex == -1)
            {
                MessageBox.Show("您还没有完成 选择用户 功能");
            }
            else {
                // 之前选好用户；
                comboBox1.SelectedIndex = form1.selectindex;
            }


            String[] xinghaos = xinghao.Split('.');
            String[] zhongleis = zhonglei.Split('.');
            String[] colors = color.Split('.');
            for (int i = 0; i < xinghaos.Length; i++) {
                comboBox2.Items.Add(xinghaos[i]);
            }

            for (int i = 0; i < zhongleis.Length; i++)
            {
                comboBox3.Items.Add(zhongleis[i]);
            }

            
            textBox2.Text = "米 / 码";
            textBox6.Text = "硕麒拉链码庄订单管理系统";
            textBox7.Text = "无";
            textBox8.Text = "无";
        }

        //确定插入
        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == "")
            {
                MessageBox.Show("数量不能为空");
            }
            else if (textBox4.Text == "")
            {
                MessageBox.Show("单价不能为空");
            }
           
            else if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("请选择用户名称");
            }
            else if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("请选择型号");
            }
            else if (comboBox3.SelectedIndex == -1)
            {
                MessageBox.Show("请选择种类");
            }
            else if (textBox11.Text == "")
            {
                MessageBox.Show("请输入颜色");
            }
            else {

                form1.preColor = textBox11.Text;

                bill = new BillPO();
                bill.Billid = textBox1.Text = DateTime.Now.ToString("yyyyMMddHHmmss") + comboBox1.SelectedItem.ToString() + (form1.dataGridView2.Rows.Count + 1);
                // 客户名称
                bill.GuestName = comboBox1.SelectedItem.ToString();
                // 型号
                bill.Xinghao = comboBox2.SelectedItem.ToString();
                bill.Zhonglei = comboBox3.SelectedItem.ToString();
                //bill.Color = comboBox4.SelectedItem.ToString();
                bill.Color = textBox11.Text;
                bill.Danwei = textBox2.Text;
                bill.Num = int.Parse(textBox3.Text);
                bill.SinglePrice = double.Parse(textBox4.Text);

                // 单价 *  数量
                bill.TotalPrice = double.Parse(textBox3.Text) * double.Parse(textBox4.Text);
                //bill.TotalPrice = double.Parse(textBox5.Text);
                bill.Kaipiaor = textBox6.Text;
                bill.Jinshour = textBox7.Text;
                bill.PicPath = "无";
                //bill.KaipiaoDate = DateTime.Now.ToString("yyyyMMddHHmmss");
                bill.Address = textBox8.Text;  // 无
                bill.OtherText = textBox9.Text;

                // 插入datagridview 中
                DataGridViewRow row = new DataGridViewRow();
                int index = form1.dataGridView2.Rows.Add(row);
                form1.dataGridView2.Rows[index].Cells[0].Value = bill.Billid;
                form1.dataGridView2.Rows[index].Cells[1].Value = bill.GuestName;
                form1.dataGridView2.Rows[index].Cells[2].Value = bill.Xinghao;
                form1.dataGridView2.Rows[index].Cells[3].Value = bill.Zhonglei;
                form1.dataGridView2.Rows[index].Cells[4].Value = bill.Color;
                form1.dataGridView2.Rows[index].Cells[5].Value = bill.Num;
                form1.dataGridView2.Rows[index].Cells[6].Value = bill.SinglePrice;
                form1.dataGridView2.Rows[index].Cells[7].Value = bill.TotalPrice;
                form1.dataGridView2.Rows[index].Cells[8].Value = bill.OtherText;
                this.Close();
            }
        }
        //取消添加
        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // 数量只输入  整数
        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if (e.KeyChar != '\b')//这是允许输入退格键
            //{
            //    if ((e.KeyChar < '0') || (e.KeyChar > '9'))//这是允许输入0-9数字
            //    {
            //        e.Handled = true;
            //    }
            //}

            //允许输入数字、小数点、删除键和负号
            if ((e.KeyChar < 48 || e.KeyChar > 57) && e.KeyChar != 8 && e.KeyChar != (char)('.') && e.KeyChar != (char)('-'))
            {
                e.Handled = true;
            }
            if (e.KeyChar == (char)('-'))
            {
                if ((sender as TextBox).Text != "")
                {
                    e.Handled = true;
                }
            }
            //小数点只能输入一次
            if (e.KeyChar == (char)('.') && ((TextBox)sender).Text.IndexOf('.') != -1)
            {
                e.Handled = true;
            }
            //第一位不能为小数点
            if (e.KeyChar == (char)('.') && ((TextBox)sender).Text == "")
            {
                e.Handled = true;
            }
            //第一位是0，第二位必须为小数点
            if (e.KeyChar != (char)('.') && e.KeyChar != 8 && ((TextBox)sender).Text == "0")
            {
                e.Handled = true;
            }
            //第一位是负号，第二位不能为小数点
            if (((TextBox)sender).Text == "-" && e.KeyChar == (char)('.'))
            {
                e.Handled = true;
            }
            e.Handled = false;
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            //第一步：判断输入的是否是数字——char.IsNumber(e.KeyChar)
            //如果是数字，可以输入（e.Handled = false;）
            //如果不是数字，则判断是否是小数点
            if (char.IsNumber(e.KeyChar))
            {
                e.Handled = false;
            }
            else
            {
                //判断输入的是否是小数点，或中文状态下的句号，或者是退格键
                //如果是小数点，循环判断每个字符是不是小数点，如果存在不能输入，如果不存在允许输入
                //如果是退格键，允许输入——if (e.KeyChar == '\b')
                //如果不是小数点也不是退格键，不允许输入
                if (e.KeyChar == Convert.ToChar("。") || e.KeyChar == Convert.ToChar("."))
                {
                    int i_d = 0;
                    for (int i = 0; i < textBox4.Text.Length; i++)
                    {
                        if (textBox4.Text.Substring(i, 1) == ".")
                        {
                            e.Handled = true;
                            i_d++;
                            return;
                        }
                    }
                    if (i_d == 0)
                    {
                        e.KeyChar = Convert.ToChar(".");//设置按键输入的值为"."
                        e.Handled = false;
                    }
                }
                else if (e.KeyChar == '\b')
                {
                    e.Handled = false;
                }

                else
                {
                    e.Handled = true;
                }
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            //第一步：判断输入的是否是数字——char.IsNumber(e.KeyChar)
            //如果是数字，可以输入（e.Handled = false;）
            //如果不是数字，则判断是否是小数点
            if (char.IsNumber(e.KeyChar))
            {
                e.Handled = false;
            }
            else
            {
                //判断输入的是否是小数点，或中文状态下的句号，或者是退格键
                //如果是小数点，循环判断每个字符是不是小数点，如果存在不能输入，如果不存在允许输入
                //如果是退格键，允许输入——if (e.KeyChar == '\b')
                //如果不是小数点也不是退格键，不允许输入
                if (e.KeyChar == Convert.ToChar("。") || e.KeyChar == Convert.ToChar("."))
                {
                    int i_d = 0;
                    for (int i = 0; i < textBox5.Text.Length; i++)
                    {
                        if (textBox5.Text.Substring(i, 1) == ".")
                        {
                            e.Handled = true;
                            i_d++;
                            return;
                        }
                    }
                    if (i_d == 0)
                    {
                        e.KeyChar = Convert.ToChar(".");//设置按键输入的值为"."
                        e.Handled = false;
                    }
                }
                else if (e.KeyChar == '\b')
                {
                    e.Handled = false;
                }
                else
                {
                    e.Handled = true;
                }
            }
        }
    }
}
