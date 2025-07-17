using System;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ANData
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            variables.source = new CancellationTokenSource();
            variables.token = variables.source.Token;
            try
            {
                Analyze.AN(variables.token);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            label1.Text = "Готово!";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fileschoosed = new OpenFileDialog())
            {
                fileschoosed.Filter = "Excel files (*.xls, *.xlsx, *.xlsm)|*.xls*";
                fileschoosed.Multiselect = true;
                if (fileschoosed.ShowDialog() == DialogResult.OK)
                {
                    variables.directory = fileschoosed.FileNames;
                    label1.Text = "Файлы загружены";
                }
            }
        }
    }
}
