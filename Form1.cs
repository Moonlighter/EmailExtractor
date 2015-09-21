using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;


namespace EmailExtractor
{
    public partial class l4 : Form
    {
        public l4()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        
                        using (StreamReader sr = File.OpenText(openFileDialog1.FileName))
                        {
                            String line;
                            while ((line = sr.ReadLine()) != null)
                            {
                                string pattern;
                                
                                pattern = @"([\w-]+\.)*?[\w]+@[\w-]+\.([\w-]+\.)*?[\w]+$";                               
                                String[] substrings = Regex.Split(line, ";");
                                foreach (string mot in substrings)
                                {                                                                                                                
                                    Match email = Regex.Match(mot, pattern);
                                    if(email.Value!=""){
                                        textBox1.AppendText(email.Value + "\n");
                                    }
                                    
                                    //*******************************************
                                    if (email.Value.Length!=0)
                                    {                                    
                                        String[] recupe = Regex.Split(email.Value, "@");                                                                                   
                                        textBox2.AppendText(recupe[0] + "\n");
                                        textBox3.AppendText(recupe[1] + "\n");                                                                              
                                    }
                                }
                            }
                            
                            this.label5.Text = textBox1.Lines.Count().ToString();
                            sr.Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var excelApp = new Excel.Application();
            
            excelApp.Visible = true;
            
            excelApp.Workbooks.Add();

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
            workSheet.Name = "EXPORT";            

            workSheet.Cells[1, "A"] = "Courriel";
            workSheet.Cells[1, "B"] = "Utilisateur";
            workSheet.Cells[1, "C"] = "Domaine";            
            var row = 1;            
            for (int i=0;i<textBox1.Lines.Length;i++)
            {
                row++;
                workSheet.Cells[row, "A"] = ConvertirChaineSansAccent(textBox1.Lines[i]);
                workSheet.Cells[row, "B"] = textBox2.Lines[i];
                workSheet.Cells[row, "C"] = textBox3.Lines[i];
            }
            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();
            workSheet.Columns[3].AutoFit();
        }
        private string ConvertirChaineSansAccent(string texte)
        {
            if ((texte != null) && (texte != string.Empty))
            {
                char[] oldChar = { 'À', 'Á', 'Â', 'Ã', 'Ä', 'Å', 'à', 'á', 'â', 'ã', 'ä', 'å', 'Ò', 'Ó', 'Ô', 'Õ', 'Ö', 'Ø', 'ò', 'ó', 'ô', 'õ', 'ö', 'ø', 'È', 'É', 'Ê', 'Ë', 'è', 'é', 'ê', 'ë', 'Ì', 'Í', 'Î', 'Ï', 'ì', 'í', 'î', 'ï', 'Ù', 'Ú', 'Û', 'Ü', 'ù', 'ú', 'û', 'ü', 'ÿ', 'Ñ', 'ñ', 'Ç', 'ç', '°' };
                char[] newChar = { 'A', 'A', 'A', 'A', 'A', 'A', 'a', 'a', 'a', 'a', 'a', 'a', 'O', 'O', 'O', 'O', 'O', 'O', 'o', 'o', 'o', 'o', 'o', 'o', 'E', 'E', 'E', 'E', 'e', 'e', 'e', 'e', 'I', 'I', 'I', 'I', 'i', 'i', 'i', 'i', 'U', 'U', 'U', 'U', 'u', 'u', 'u', 'u', 'y', 'N', 'n', 'C', 'c', ' ' };
                int i = 0;

                foreach (char monc in oldChar)
                {
                    texte = texte.Replace(monc, newChar[i]);
                    i++;
                }
            }
            return texte;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label5.Text = "";
        }
    }
}