using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Collections;
using System.Drawing.Printing;

namespace MiniWord
{
    public partial class MDIParent1 : Form
    {
        private int childFormNumber = 0;
        //private ArrayList tfArr;
        private PageSettings storedPageSettings = null;
        private PrintDocument printDoc = new PrintDocument();
        private StringReader lineReader;
        private Font printFont = new Font("Arial",16);
       
        public MDIParent1()
        {
            InitializeComponent();
            //tfArr = new ArrayList();
        }

        private void ShowNewForm(object sender, EventArgs e)
        {
            TextForm childForm = new TextForm();
            childForm.MdiParent = this;
            childForm.Text = "新建文档" + childFormNumber++;
            childForm.Show();
            //tfArr.Add(childForm);
        }

        private void OpenFile(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            openFileDialog.Filter = "文本文件(*.txt)|*.txt|Rtf文档(*.rtf)|*.rtf";
            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                string str;
                string FileName = openFileDialog.FileName;
                StreamReader sr = new StreamReader(FileName,Encoding.Default);
                TextForm tf = new TextForm();
                tf.MdiParent = this;
                tf.Text = FileName;
                tf.Show();
                while ((str= sr.ReadLine())!= null)
                {
                    tf.getRichTextBox().Text += str + '\n';
                }
            }
        }

        private void ExitToolsStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void CutToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void CopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void PasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

    

        private void CascadeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.Cascade);
        }

        private void TileVerticalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileVertical);
        }

        private void TileHorizontalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileHorizontal);
        }

        private void ArrangeIconsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.ArrangeIcons);
        }

        private void CloseAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (Form childForm in MdiChildren)
            {
                childForm.Close();
            }
        }

        private void toolStrip_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void menuStrip_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
        private void saveFile(TextForm form, object sender, EventArgs e) 
        {
            string fileName = form.Text;
            if (File.Exists(fileName))
            {
                FileStream fs = new FileStream(fileName, FileMode.Open);
                StreamWriter sw = new StreamWriter(fs);
                try
                {
                    foreach (string line in form.getRichTextBox().Lines)
                    {
                        sw.WriteLine(line);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
                finally
                {
                    sw.Close();
                    fs.Close();
                }
            }
            else
            {
                SaveAs(form,sender, e);
            }
        }
        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            TextForm form = (TextForm)this.ActiveMdiChild;
            saveFile(form, sender, e);
        }

        private void fileMenu_Click(object sender, EventArgs e)
        {

        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveToolStripButton_Click(sender, e);
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {

        }

        private void closeAll(object sender, EventArgs e)
        {
            foreach (Form childForm in MdiChildren)
            {
                childForm.Close();
            }
        }

        private void closeCurFile(object sender, EventArgs e)
        {
            this.ActiveMdiChild.Close();
        }

        private void saveAll(object sender, EventArgs e)
        {
            foreach (Form childForm in MdiChildren)
            {
               saveFile((TextForm)childForm, sender, e);
            }
        }

        private void SaveAs(TextForm form,object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            saveFileDialog.Filter = "文本文件(*.txt)|*.txt|Rtf文档(*.rtf)|*.rtf";
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                string FileName = saveFileDialog.FileName;
                FileStream fs = new FileStream(FileName, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs);

                //TextForm form = (TextForm)this.ActiveMdiChild;

                try
                {
                    foreach (string line in form.getRichTextBox().Lines)
                    {
                        sw.WriteLine(line);
                    }
                    form.Text = FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
                finally
                {
                    sw.Close();
                    fs.Close();
                }

            }

        }
        private void SaveAs(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            saveFileDialog.Filter = "文本文件(*.txt)|*.txt|Rtf文档(*.rtf)|*.rtf";
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                string FileName = saveFileDialog.FileName;
                FileStream fs = new FileStream(FileName, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs);

                TextForm form = (TextForm)this.ActiveMdiChild;

                try
                {
                    foreach (string line in form.getRichTextBox().Lines)
                    {
                        sw.WriteLine(line);
                    }
                    form.Text = FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
                finally
                {
                    sw.Close();
                    fs.Close();
                }

            }

        }

        private void setColor(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                Color color = colorDialog.Color;
                TextForm tf = (TextForm)this.ActiveMdiChild;
                tf.getRichTextBox().ForeColor = color;
            }
            
        }

        private void setFont(object sender, EventArgs e)
        {
            FontDialog fontDialog = new FontDialog();
            if (fontDialog.ShowDialog() == DialogResult.OK)
            {
                Font font = fontDialog.Font;
                TextForm tf = (TextForm)this.ActiveMdiChild;
                tf.getRichTextBox().Font = font;
                printFont = font;
            }

            
        }

        
        private void printFile(object sender, EventArgs e)
        {

            TextForm tf = (TextForm)this.ActiveMdiChild;
            if (!File.Exists(tf.Text))
            {
                SaveAs(tf, sender, e);
            }
            StreamReader streamToPrint = new StreamReader(tf.Text);
            try
            {
                TextPrintDoc tpd = new TextPrintDoc(streamToPrint); 
                
                
                if (storedPageSettings != null) 
                {
                    tpd.DefaultPageSettings = storedPageSettings;
                }
                PrintDialog pd = new PrintDialog();
                pd.Document = tpd;
                if (pd.ShowDialog() == DialogResult.OK)
                {
                    tpd.Print(); //打印
                }
            }
            catch(Exception ee)
            {
                MessageBox.Show(ee.Message);
            }
            finally
            {
                streamToPrint.Close();
            }
        }

  
        private void printSetup(object sender, EventArgs e)
        {
            try
            {
                PageSetupDialog psd = new PageSetupDialog();
                if (storedPageSettings == null)
                {
                    storedPageSettings = new PageSettings();
                }
                psd.PageSettings = storedPageSettings;
                psd.ShowDialog();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
        }
        
        private void printPreview(object sender, EventArgs e)
        {
            PrintPreviewDialog ppd = new PrintPreviewDialog();
            TextForm tf = (TextForm)this.ActiveMdiChild;
            String s = tf.getRichTextBox().Text.ToString();
            lineReader = new StringReader(s);
            PrintDocument document = new PrintDocument();
            document.PrintPage += new PrintPageEventHandler(PrintDocument_PrintPage);
            ppd.Document = document;
            ppd.ShowDialog();

        }
        private void PrintDocument_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            Graphics g = e.Graphics;
            float linesPerPage = 0;
            float yPosition = 0;
            int count = 0;

            float leftMargin = e.MarginBounds.Left;
            float topMargin = e.MarginBounds.Top;
            string line = null;

            SolidBrush myBrush = new SolidBrush(Color.Black);

            linesPerPage = e.MarginBounds.Height / printFont.GetHeight(g);

            while (count < linesPerPage && ((line = lineReader.ReadLine()) != null))
            {
                yPosition = topMargin + (count * printFont.GetHeight(g));
                g.DrawString(line, printFont, myBrush, leftMargin, yPosition, new StringFormat());
                count++;
            }
            if (line != null)
                e.HasMorePages = true;
            else
                e.HasMorePages = false;
        }
        //将流输出到打印机
        public class TextPrintDoc : PrintDocument
        {
            private Font printFont = null;
            private StreamReader streamToPrint = null;

            public TextPrintDoc(StreamReader streamToPrint)
                : base()
            {
                this.streamToPrint = streamToPrint;
            }
            //重写 OnPrintPage 以为文档提供打印逻辑
            protected override void OnPrintPage(PrintPageEventArgs ev)
            {

                base.OnPrintPage(ev);

                float lpp = 0;
                float yPos = 0;
                int count = 0;
                float leftMargin = ev.MarginBounds.Left;
                float topMargin = ev.MarginBounds.Top;
                String line = null;

                //算出每页的行数
                //在事件上使用 MarginBounds 以达到此目的
                lpp = ev.MarginBounds.Height / printFont.GetHeight(ev.Graphics);

                //现在，在文件上重复此操作以输出每行
                //注意：假设单行比页宽窄
                //首先检查行数，以便看不到不打印的行
                while (count < lpp && ((line = streamToPrint.ReadLine()) != null))
                {
                    yPos = topMargin + (count * printFont.GetHeight(ev.Graphics));

                    ev.Graphics.DrawString(line, printFont, Brushes.Black, leftMargin, yPos, new StringFormat());

                    count++;
                }

                //如果有多行，则另外打印一页
                if (line != null)
                    ev.HasMorePages = true;
                else
                    ev.HasMorePages = false;
            }

        }
        private SearchForm searchForm = null;
        private void SearchInFile(object sender, EventArgs e)
        {
            searchForm = new SearchForm((TextForm)this.ActiveMdiChild);
            searchForm.Show();
        }

        private void 查找下一个ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(searchForm ==null)
                MessageBox.Show("请先进行查找");
            else
                searchForm.button1_Click(sender,e);
        }

        private void 插入日期时间ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            String dt = System.DateTime.Now.ToString();
            TextForm tf = (TextForm)this.ActiveMdiChild;
            tf.getRichTextBox().Text += dt;
        }

        private void 自动换行ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.wrapItem.Checked = !this.wrapItem.Checked;
            RichTextBox rich = ((TextForm)this.ActiveMdiChild).getRichTextBox();
            rich.WordWrap = this.wrapItem.Checked ? true : false ;
           
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TextForm tf = (TextForm)this.ActiveMdiChild;
            tf.getRichTextBox().SelectAll();
            tf.getRichTextBox().Focus();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About about = new About();
            about.Show();
        }
    }
}
