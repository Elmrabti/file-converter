using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

/* 
-----------------------------------------------
MY C# Project
Name:        File Converter
Author :     EL MRABTI KHALID
License:     1.0
----------------------------------------------- */
namespace converter
{
   
    public partial class Form1 : Form
    {
        public string mSelectedFile;
        public string mSelectedFolder;
        Microsoft.Office.Interop.Word.Document wordDoc { get; set; }
        public Form1()
        {
            InitializeComponent();
            
        }


        // this code for get the file path
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "All Files (*.*)|*.*";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
           mSelectedFile = choofdlog.FileName;
            textBox1.Text = mSelectedFile;
            }
             
            else
                mSelectedFile = string.Empty;

            
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
         
        }

        // this code is responsible to convert button Action 
        private void btn_convert_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Length<=0&&comboBox1.Text.Length<=0)
            {
                MessageBox.Show("Plz check if there is a Empty Text Field !");

            }
            else
            {
               
                if (comboBox1.SelectedItem.Equals("word to .pdf"))
                {
                    if (mSelectedFile.Contains(".docx"))
                    {
                        using (Form2 frm = new Form2(toPdf))
                        {
                            frm.ShowDialog(this);
                        }
                    }else
                    {
                        MessageBox.Show("Plz Choose  a Word File !");
                    }
                    
                }
                else if (comboBox1.SelectedItem.Equals("any image format to .png"))
                {
                    using (Form2 frm = new Form2(toPng))
                    {
                        frm.ShowDialog(this);
                    }
                }
                else if (comboBox1.SelectedItem.Equals("any image format to .jpeg"))
                {
                    using (Form2 frm = new Form2(toJpeg))
                    {
                        frm.ShowDialog(this);
                    }
                }
                else if (comboBox1.SelectedItem.Equals("any image format to .gif"))
                {
                    using (Form2 frm = new Form2(toGif))
                    {
                        frm.ShowDialog(this);
                    }
                }
                else if (comboBox1.SelectedItem.Equals("any image format to .icon"))
                {
                    using (Form2 frm = new Form2(toIcon))
                    {
                        frm.ShowDialog(this);
                    }
                }
            }

        }

        // this is the converter Methods
        public void toPng()
        {
               System.Drawing.Image image = System.Drawing.Image.FromFile(mSelectedFile);
                image.Save(mSelectedFolder + textBox3.Text + ".png", System.Drawing.Imaging.ImageFormat.Png);
        }
        public void toJpeg()
        {
            System.Drawing.Image image = System.Drawing.Image.FromFile(mSelectedFile);
            image.Save(mSelectedFolder + textBox3.Text + ".jpeg", System.Drawing.Imaging.ImageFormat.Jpeg);
        }
        public void toGif()
        {
            System.Drawing.Image image = System.Drawing.Image.FromFile(mSelectedFile);
            image.Save(mSelectedFolder + textBox3.Text + ".gif", System.Drawing.Imaging.ImageFormat.Gif);
        }
        public void toIcon()
        {
            System.Drawing.Image image = System.Drawing.Image.FromFile(mSelectedFile);
            image.Save(mSelectedFolder + textBox3.Text + ".icon", System.Drawing.Imaging.ImageFormat.Icon);
        } 
        public void toPdf()
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            wordDoc = app.Documents.Open(mSelectedFile);
            mSelectedFolder = mSelectedFolder + "/";
            wordDoc.ExportAsFixedFormat(mSelectedFolder + textBox3.Text + ".pdf", Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
            wordDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
            app.Quit();
            Marshal.ReleaseComObject(wordDoc);
            Marshal.ReleaseComObject(app);
        }


        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }


        // this code for get folder Path
        private void button1_Click_1(object sender, EventArgs e)
        {
            
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    mSelectedFolder = fbd.SelectedPath + "/";
                    textBox2.Text = mSelectedFolder;
                }

                else
                    mSelectedFolder = string.Empty;
            
                
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            comboBox1.Text = "";
        }
    }
   
}
