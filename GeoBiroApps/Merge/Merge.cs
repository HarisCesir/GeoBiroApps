using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GeoBiroApps.Merge
{
    public partial class Merge : Form
    {
        List<string> path = new List<string>();
        string SacuvanTxt;
        string Obavijest = " ";
        List<string> Lista = new List<string>();
        public Merge()
        {
            InitializeComponent();
            button2.Enabled = false;
            button3.Enabled = false;
        }

        private void Merge_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Multiselect = true;
            open.Filter = "Text|*.txt";
            open.FilterIndex = 1;

            if (open.ShowDialog() == DialogResult.OK)
            {
                path = open.FileNames.ToList();
                button2.Enabled = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();

            saveDialog.FileName = "obrazac.txt";
            if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var file = File.Create(saveDialog.FileName);
                file.Close();


            }
          
                foreach (string a in path)
                {
                    int BrojLinije = 1;
                    string line;
                    System.IO.StreamReader file =
                     new System.IO.StreamReader(a);
                    while ((line = file.ReadLine()) != null)
                    {
                        if (!string.IsNullOrEmpty(line))
                        {

                                try
                                {
                                     Lista.Add(line);
                                     BrojLinije++;
                                  
                                }
                                catch
                                {

                                 Obavijest = Obavijest + "Greška u txt fajlu: " + a + " Na liniji:  " + BrojLinije;
                                 break;

                                }

                            }

                        }
                file.Close();
            }

            if (Obavijest == " ")
            {
                try
                {
                    using (var tw = new StreamWriter(saveDialog.FileName, true))
                    {
                        int broj = 1;
                        foreach (string t in Lista)
                        {
                            if (t.Substring(t.Length - 3) == "POL")
                            {
                                tw.WriteLine(t);
                            }
                            else
                            {
                                string k = broj + "," + t.Substring(t.IndexOf(',') + 1);

                                tw.WriteLine(k);
                                broj++;
                            }
                        }


                    }

                    MessageBox.Show("Završeno");
                }
                catch
                {
                    MessageBox.Show("Dogodila se greska prilikom pisanja obrasca molimo pokušajte ponovo");
                   
                }
               
            }
            else
            {
                MessageBox.Show(Obavijest);
            }

            
                

                

            
           
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            String resourceName = "templateMerge.xlsx";
            String path = System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);


            Assembly asm = Assembly.GetExecutingAssembly();
            string res = string.Format("{0}.Resources." + resourceName, asm.GetName().Name);
            Stream stream = asm.GetManifestResourceStream(res);
            try
            {
                using (Stream filea = File.Create(path + @"\" + resourceName))
                {
                    CopyStream(stream, filea);
                }

            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message);
            }


            Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlsWorkBook = xls.Workbooks.Open(path + @"\" + resourceName);





            try
            {


                int zadnjiROwO25 = 6;
                int zadnjiROwTacke = 5;
                int broj = 1;
                string line;
                System.IO.StreamReader file =
                 new System.IO.StreamReader(SacuvanTxt);
                while ((line = file.ReadLine()) != null)
                {

                    if (line.Substring(line.Length - 3) == "POL")
                    {

                        Microsoft.Office.Interop.Excel.Worksheet O25 = xlsWorkBook.Worksheets[1];
                        O25 = xlsWorkBook.Sheets[1];
                        O25.Select(true);

                        O25.Cells[zadnjiROwO25, 1] = line.Split(',')[0];
                        O25.Cells[zadnjiROwO25, 3] = line.Split(',')[1];
                        O25.Cells[zadnjiROwO25, 4] = line.Split(',')[2];
                        O25.Cells[zadnjiROwO25, 5] = line.Split(',')[3];
                        zadnjiROwO25++;

                        Microsoft.Office.Interop.Excel.Worksheet Tacke = xlsWorkBook.Worksheets[2];
                        Tacke = xlsWorkBook.Sheets[2];
                        Tacke.Select(true);

                        Tacke.Cells[zadnjiROwTacke, 1] = broj;
                        Tacke.Cells[zadnjiROwTacke, 2] = line.Split(',')[0];
                        Tacke.Cells[zadnjiROwTacke, 3] = "POL";
                        Tacke.Cells[zadnjiROwTacke, 4] = line.Split(',')[1];
                        Tacke.Cells[zadnjiROwTacke, 5] = line.Split(',')[2];
                        Tacke.Cells[zadnjiROwTacke, 6] = line.Split(',')[3];
                        broj++;
                        zadnjiROwTacke++;
                    }
                    else
                    {
                        Microsoft.Office.Interop.Excel.Worksheet Tacke = xlsWorkBook.Worksheets[2];
                        Tacke = xlsWorkBook.Sheets[2];
                        Tacke.Select(true);
                        Tacke.Cells[zadnjiROwTacke, 1] = broj;
                        Tacke.Cells[zadnjiROwTacke, 2] = line.Split(',')[0];

                        Tacke.Cells[zadnjiROwTacke, 4] = line.Split(',')[1];
                        Tacke.Cells[zadnjiROwTacke, 5] = line.Split(',')[2];
                        Tacke.Cells[zadnjiROwTacke, 6] = line.Split(',')[3];
                        broj++;
                        zadnjiROwTacke++;

                    }

                }

                file.Close();

                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx, *.xls)|*.xlsx; *.xls";
                saveDialog.FilterIndex = 2;
                saveDialog.FileName = "Excel.xlsx";

                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {

                    Microsoft.Office.Interop.Excel.Worksheet Tacke = xlsWorkBook.Worksheets[2];

                    Tacke.SaveAs(saveDialog.FileName);

                    //  Microsoft.Office.Interop.Excel.Worksheet O25 = xlsWorkBook.Worksheets[1];
                    // O25.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Zavrseno", "Informacija");
                }
                xlsWorkBook.Close(true);

                xls.Quit();

                Marshal.ReleaseComObject(xlsWorkBook);
                //Marshal.ReleaseComObject(Tacke);
                Marshal.ReleaseComObject(xls);
            }
            catch
            {
                xlsWorkBook.Close(true);

                xls.Quit();

                Marshal.ReleaseComObject(xlsWorkBook);
                //Marshal.ReleaseComObject(Tacke);
                Marshal.ReleaseComObject(xls);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Multiselect = false;
            open.Filter = "Text|*.txt";
            open.FilterIndex = 1;

            if (open.ShowDialog() == DialogResult.OK)
            {
                SacuvanTxt = open.FileName;
                button3.Enabled = true;
            }
        }

        public void CopyStream(Stream input, Stream output)
        {
            byte[] buffer = new byte[8 * 1024];
            int len;
            while ((len = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, len);
            }
        } //kraj copy stream
    }
}
