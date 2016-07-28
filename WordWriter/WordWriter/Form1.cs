using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using System.IO;

namespace WordWriter
{
    public partial class Form1 : Form
    {
        String sourceFile = "Z:\\BGM Engineering Forms\\Quote Forms_Includes BOM Form\\Quote_Template_07282016.docx";
        String destinationFile = "C:\\Temp1\\Quote.docx";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_click(object sender, EventArgs e)
        {
            if (!Directory.Exists("c:\\Temp1"))
            {
                Directory.CreateDirectory("C:\\Temp1");
            }
            if (Directory.Exists("Z:\\BGM Engineering Forms\\Quote Forms_Includes BOM Form"))
            {
                System.IO.File.Copy(sourceFile, destinationFile, true);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(destinationFile, true))
            {
                String docText;

                /* Use StreamReader to read the file */
                using (StreamReader sr = new StreamReader(document.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                /* Replace Text */
                Regex regexText = new Regex("##Quote_Number##");
                docText = regexText.Replace(docText, "BGMQ2016-9999-A");

                regexText = new Regex("##Customer_Name##");
                docText = regexText.Replace(docText, "AAA Corp.");

                /* Write the file */
                using (StreamWriter sw = new StreamWriter(document.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }
    }
}
