using PapersStatisticTools;
using StatisticalTools.po;
using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace StatisticalTools
{
    public partial class NewGuest : Form
    {
        public NewGuest()
        {
            InitializeComponent();
        }


        // 仅输入 数字
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9') || (e.KeyChar == 8))
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        //// 仅输入汉字
        //private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    Regex rg = new Regex("^[\u4e00-\u9fa5]$");
        //    if (!rg.IsMatch(e.KeyChar.ToString()) && e.KeyChar != '\b') //'\b'是退格键
        //    {
        //        e.Handled = true;
        //    }
        //}

        // 确定按钮
        private void button1_Click(object sender, EventArgs e)
        {
            Guest guest = new Guest();
            guest.Name = textBox1.Text;
            guest.Tel = textBox2.Text;
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("检查客户名 电话 是否输入为空");
            }
            else {
                MysqlManager mysqlManager = new MysqlManager();
                Boolean f = mysqlManager.checkGuest(guest.Name);
                if (!f)
                {
                    MessageBox.Show("客户名已存在，请更换其他标识");
                }
                else
                {
                    int result = mysqlManager.insertGuest(guest);
                    if (result != 1)
                    {
                        MessageBox.Show("插入失败");
                    }
                    else
                    {
                        MessageBox.Show("添加用户成功");
                    }
                    
                }
                this.Close();
            }
        }

        // 取消按钮
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
