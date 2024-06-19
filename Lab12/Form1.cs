using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Lab12
{
    public partial class Form1 : Form
    {
        private string fn = @"D:\Project_C#\Lab12\dovidka.dotx";
        private Word.Application word = new Word.Application();
        private Word.Document doc;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Object missingObj = System.Reflection.Missing.Value;
            Object templatePathObj = fn;

            try
            {
                doc = word.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
                doc.Activate();

                foreach (Word.FormField f in doc.FormFields)
                {
                    switch (f.Name)
                    {
                        case "name":
                            f.Range.Text = textBox1.Text;
                            break;
                        case "course":
                            f.Range.Text = textBox2.Text;
                            break;
                        case "group":
                            f.Range.Text = textBox3.Text;
                            break;
                        case "format":
                            f.Range.Text = textBox4.Text;
                            break;
                        case "zaklad":
                            f.Range.Text = textBox5.Text;
                            break;
                        default:
                            break;
                    }
                }

                word.Visible = true;
            }
            catch (Exception error)
            {
                MessageBox.Show("An error occurred: " + error.Message);
                
                if (doc != null)
                {
                    doc.Close(false);
                    doc = null;
                }
                
                if (word != null)
                {
                    word.Quit();
                    word = null;
                }
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (doc != null)
            {
                doc.Close(false);
                doc = null;
            }
            
            if (word != null)
            {
                word.Quit();
                word = null;
            }
        }
    }
}
