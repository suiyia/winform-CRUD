using PapersStatisticTools;
using System;

using System.Windows.Forms;
// 暂时不用 
namespace StatisticalTools.po
{
    public partial class Newcolor : Form
    {
        public Newcolor()
        {
            InitializeComponent();
        }

        
        private void button1_Click(object sender, EventArgs e)
        {
            Color color = new Color();
            color.Cid = textBox1.Text;
            color.Mcolor = textBox2.Text;
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("检查颜色型号 颜色中文 是否输入正确");
            }
            else {
                MysqlManager mysqlManager = new MysqlManager();
                int result = mysqlManager.insertColor(color);
                MessageBox.Show("添加 "+ textBox1.Text+"#"+ textBox2.Text + " 成功");
                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
