using PapersStatisticTools;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
// 暂时不用 
namespace StatisticalTools.po
{
    public partial class Newzhonglei : Form
    {
        public Newzhonglei()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Zhonglei zhonglei = new Zhonglei();
            zhonglei.Color = textBox1.Text;
            if (textBox1.Text == "")
            {
                MessageBox.Show("检查种类是否输入正确");
            }
            else
            {
                MysqlManager mysqlManager = new MysqlManager();
                int result = mysqlManager.insertZhonglei(zhonglei);

                MessageBox.Show("添加种类： "+ textBox1.Text + " 成功");
                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
