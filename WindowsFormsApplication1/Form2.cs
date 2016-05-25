using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form2 : Form
    {
        Form1 form1 ;
        public Form2(Form1 m_parent)
        {
            InitializeComponent();
            TB1 = textBox1;
            form1 = m_parent;
         
        }
        public TextBox TB1 = new TextBox();
        private void Form2_Load(object sender, EventArgs e)
        {

        }

        public void button1_Click(object sender, EventArgs e)
        {

            form1.dataGridView1.Rows[form1.dataGridView1.CurrentCell.RowIndex].Cells[form1.dataGridView1.CurrentCell.ColumnIndex].Value = textBox2.Text.ToString();
            this.Close();              
        }
    }
}
