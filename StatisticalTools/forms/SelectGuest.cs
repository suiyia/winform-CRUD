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

namespace StatisticalTools.forms
{
    public partial class SelectGuest : Form
    {
        private Form1 form1;
        public SelectGuest(Form1 form)
        {
            InitializeComponent();
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            // 从数据库读取用户
            MysqlManager manager = new MysqlManager();
            DataTable table = manager.getGuestNames();
            String[] guests = new string[table.Rows.Count];
            for (int i = 0; i < table.Rows.Count; i++)
            {
                comboBox1.Items.Add(table.Rows[i]["name"]);
            }
            form1 = form;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("您还未选择用户");
            }
            else {
                form1.selectindex = comboBox1.SelectedIndex;
                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
