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
    public partial class ChooseGuest : Form
    {
        private Form1 form1;

        public ChooseGuest(Form1 f)
        {
            InitializeComponent();
            form1 = f;
        }

        private void ChooseGuest_Load(object sender, EventArgs e)
        {
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            MysqlManager manager = new MysqlManager();
            DataTable table = manager.getGuestNames();
            String[] guests = new string[table.Rows.Count];
            for (int i = 0; i < table.Rows.Count; i++)
            {
                comboBox1.Items.Add(table.Rows[i]["name"]);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("请先选择用户");
            }
            else {
                this.Close();
                NewBill bill = new NewBill(form1,comboBox1.SelectedIndex);
                bill.StartPosition = FormStartPosition.CenterParent;
                bill.ShowDialog();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        
    }
}
