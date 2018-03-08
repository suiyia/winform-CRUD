using PapersStatisticTools;
using StatisticalTools.po;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StatisticalTools.forms
{
    public partial class DeleteGuest : Form
    {
        public DeleteGuest()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            Guest guest = new Guest();
            guest.Name = textBox1.Text;
            if (textBox1.Text == "")
            {
                MessageBox.Show("用户名不为空");
            }
            else
            {
                MysqlManager mysqlManager = new MysqlManager();
                int result = mysqlManager.deleteGuest(guest);

                if (result != 1)
                {
                    MessageBox.Show("删除失败");
                }
                else
                {
                    MessageBox.Show("删除用户成功");
                }
                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
