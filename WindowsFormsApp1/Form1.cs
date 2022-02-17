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
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
                openFileDialog1.InitialDirectory = "c:\\";
                openFileDialog1.Filter = "word files (*.docx)|*.docx";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;
                openFileDialog1.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            string directory = Path.GetDirectoryName(openFileDialog1.FileName);
            DocToPDFConverter converter = new DocToPDFConverter();
            WordDocument wordDocument = new WordDocument(openFileDialog1.FileName, FormatType.Docx);
            PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument);
            saveFileDialog1.ShowDialog();
            pdfDocument.Save(saveFileDialog1.FileName);
            pdfDocument.Close(true);
            wordDocument.Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            textBox1.Text = "rsdgdfsh";
        }
    }
}
