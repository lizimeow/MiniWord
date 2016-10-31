using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;

namespace MiniWord
{
    public partial class SearchForm : Form
    {
        private TextForm activeForm;
        private int res = 0;
        private ArrayList rArr = null;
        public SearchForm()
        {
            InitializeComponent();
        }
        public SearchForm(TextForm tf)
        {
            InitializeComponent();
            activeForm = tf;
            this.res = 0;
            this.rArr = new ArrayList();
        }
        public void button1_Click(object sender, EventArgs e)
        {
            //count++;
            String str = searchText.Text;
            string file = activeForm.getRichTextBox().Text;
            
            if(str.Length == 0)
            {
                MessageBox.Show("查找项为空，请重新输入");
            }

            if (this.checkBox1.Checked)
            {
                str = str.ToLower();
                file = file.ToLower();
            }
            res = file.IndexOf(str,res);
            if (res == -1)
            {
                MessageBox.Show("查找结束");
                this.Close();
            }
            else
            {
                activeForm.getRichTextBox().Select(res, str.Length);
                activeForm.getRichTextBox().Focus();
                res++;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
