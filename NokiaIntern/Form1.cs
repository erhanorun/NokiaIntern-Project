using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using Xceed.Words.NET;
using Word = Microsoft.Office.Interop.Word;


namespace WindowsFormsApp6


{

    public partial class Form1 : Form
    {





        public static string getBetween(string strSource, string strStart, string strEnd)
        {
            int Start, End;
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }
            else
            {
                return "";
            }
        }

        public static string CreateMD5(string input)
        {

            using (System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create())
            {
                byte[] inputBytes = System.Text.Encoding.ASCII.GetBytes(input);
                byte[] hashBytes = md5.ComputeHash(inputBytes);

                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < hashBytes.Length; i++)
                {
                    sb.Append(hashBytes[i].ToString("X2"));
                }
                return sb.ToString();
            }
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            word.Visible = true; 
            Microsoft.Office.Interop.Word.Document wordDocument; 
            object wordObj = System.Reflection.Missing.Value;
            wordDocument = word.Documents.Add(ref wordObj);
            word.Selection.TypeText(label13.Text + " " + label7.Text);
            word.Selection.Font.Size = 24;
            word.Selection.Font.Name = "Arial";
            word = null;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void folderBrowserDialog2_HelpRequest(object sender, EventArgs e)
        {

        }


        public void button1_Click_1(object sender, EventArgs e)
        {   
            OpenFileDialog f1 = new OpenFileDialog();

            string directory;
            string folder;
            
             if (f1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                folder = Path.GetFileName(f1.FileName);
                directory = Path.GetDirectoryName(f1.FileName);
                
                
                System.IO.FileInfo ff = new System.IO.FileInfo(f1.FileName);
                string DosyaUzantisi = ff.Extension;
                if (DosyaUzantisi == ".dll" ) {

                    string[] filePaths = Directory.GetFiles(@directory, "*.docx");
                    string wordfile = string.Join("", filePaths);
                    string x = System.IO.Path.GetDirectoryName(folder);
                    File.Exists(x);
                    string ne_model = getBetween(folder, "Nokia_", "_R4");
                    label8.Text = ne_model;
                    string ne_release = getBetween(folder, "", "");
                    label10.Text = ne_release;
                    string adapter_version_communicator = getBetween(folder, "MGW_", ".dll");
                    label7.Text = adapter_version_communicator;
                    string Document_Name_Version = getBetween(folder, "", ".dll");
                    label13.Text = Document_Name_Version;
                    string md5 = CreateMD5(folder);
                    label9.Text = md5;

                }

                if (DosyaUzantisi == ".nak")
                {

                    string[] filePaths = Directory.GetFiles(@directory, "*.docx");
                    string wordfile = string.Join("", filePaths);
                    string x = System.IO.Path.GetDirectoryName(folder);
                    File.Exists(x);
                    string ne_model = getBetween(folder, "nokia-", "-usp");
                    label8.Text = ne_model;
                    string ne_release = getBetween(folder, "", "");
                    label10.Text = ne_release;
                    string adapter_version_communicator = getBetween(folder, "usp1707_", ".nak");
                    label7.Text = adapter_version_communicator;
                    string Document_Name_Version = getBetween(folder, "", ".nak");
                    label13.Text = Document_Name_Version;
                    string md5 = CreateMD5(folder);
                    label9.Text = md5;

                }
                if (DosyaUzantisi == ".jar")
                {
                   

                    
                    string[] filePaths = Directory.GetFiles(@directory, "*.docx");
                    string wordfile = string.Join("", filePaths);
                    string x = System.IO.Path.GetDirectoryName(folder);
                    File.Exists(x);
                    string ne_model = getBetween(folder, "NOKIA_", "_R31");
                    label8.Text = ne_model;
                    string ne_release = getBetween(folder, "", "");
                    label10.Text = ne_release;
                    string adapter_version_communicator = getBetween(folder, "DROUTER_", ".jar");
                    label7.Text = adapter_version_communicator;
                    string Document_Name_Version = getBetween(folder, "", ".jar");
                    label13.Text = Document_Name_Version;
                    string md5 = CreateMD5(folder);
                    label9.Text = md5;
                }


            }




        }
                













                


    






        public void label8_Click(object sender, EventArgs e)
        {


        }

        private void label7_Click(object sender, EventArgs e)
        {

        }



        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }


    }
}
