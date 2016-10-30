using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MiniWord
{
    public partial class TextForm : Form
    {
        public TextForm()
        {
            InitializeComponent();
        }

        private void TextForm_Load(object sender, EventArgs e)
        {

        }
        public RichTextBox getRichTextBox()
        {
            return this.richTextBox1;
        }
    }
}
