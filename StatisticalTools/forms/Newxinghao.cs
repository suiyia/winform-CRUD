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
    public partial class Newxinghao : Form
    {
        public Newxinghao()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Xinghao xinghao = new Xinghao();
            xinghao.Kind = textBox1.Text;
            if (textBox1.Text == "")
            {
                MessageBox.Show("检查型号是否输入正确");
            }
            else
            {
                MysqlManager mysqlManager = new MysqlManager();
                int result = mysqlManager.insertXinghao(xinghao);

                MessageBox.Show("插入型号 "+ textBox1.Text + " 成功");
                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
