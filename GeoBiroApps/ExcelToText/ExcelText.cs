using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GeoBiroApps.ExcelToText
{
    public partial class ExcelText : Form
    {
        public string url = " ";
        public ExcelText()
        {
            InitializeComponent();
            button2.Enabled = false;
        }

        private void ExcelText_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.ShowDialog();
            url = openFileDialog1.FileName;
            if (url != " ")
            {
                button2.Enabled = true;
            }
            else
            {
                button2.Enabled = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(url);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            xlRange = (Range)xlWorksheet.Cells[xlWorksheet.Rows.Count, 1];
            long lastRow = (long)xlRange.get_End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row;

            SaveFileDialog saveDialog = new SaveFileDialog();

            saveDialog.FileName = "obrazac.txt";
            if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var file = File.Create(saveDialog.FileName);
                file.Close();


            }




            TimeSpan vrijeme = new TimeSpan(08, 00, 00);
            DateTime datum = dateTimePicker1.Value;
            for (int i = 1; i <= lastRow; i++)
            {


                using (var tw = new StreamWriter(saveDialog.FileName, true))
                {
                    tw.WriteLine("GPS,PN" + (xlWorksheet.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value.ToString() + ",LA" + (xlWorksheet.Cells[i, 5] as Microsoft.Office.Interop.Excel.Range).Value.ToString() + ",LN" + (xlWorksheet.Cells[i, 6] as Microsoft.Office.Interop.Excel.Range).Value.ToString() + ",EL" + (xlWorksheet.Cells[i, 7] as Microsoft.Office.Interop.Excel.Range).Value.ToString() + ",--" + (xlWorksheet.Cells[i, 8] as Microsoft.Office.Interop.Excel.Range).Value.ToString());
                    tw.WriteLine("--GS,PN" + (xlWorksheet.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value.ToString() + ",N " + (xlWorksheet.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value.ToString() + ",E " + (xlWorksheet.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Value.ToString() + ",EL" + (xlWorksheet.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Value.ToString() + ",--" + (xlWorksheet.Cells[i, 8] as Microsoft.Office.Interop.Excel.Range).Value.ToString());
                    tw.WriteLine("--GT,PN1,SW2013,ST457715000,EW2013,ET457719000");
                    Random r = new Random();

                    tw.WriteLine("--Valid Readings: " + 8 + " of " + 8);
                    tw.WriteLine("--Fixed Readings: " + 8 + " of " + 8);
                    decimal a = Convert.ToDecimal(r.NextDouble() * (0.1 - 0.0) + 0.0);
                    decimal min = Convert.ToDecimal((xlWorksheet.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value.ToString()) - a;
                    decimal max = Convert.ToDecimal((xlWorksheet.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value.ToString()) + a;

                    tw.WriteLine("--Nor Min: " + decimal.Round(min, 4) + " Max: " + decimal.Round(max, 4));
                    a = Convert.ToDecimal(r.NextDouble() * (0.1 - 0.0) + 0.0);
                    min = Convert.ToDecimal((xlWorksheet.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Value.ToString()) - a;
                    max = Convert.ToDecimal((xlWorksheet.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Value.ToString()) + a;

                    tw.WriteLine("--Eas Min: " + decimal.Round(min, 4) + " Max: " + decimal.Round(max, 4));

                    a = Convert.ToDecimal(r.NextDouble() * (0.1 - 0.0) + 0.0);
                    min = Convert.ToDecimal((xlWorksheet.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Value.ToString()) - a;
                    max = Convert.ToDecimal((xlWorksheet.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Value.ToString()) + a;

                    tw.WriteLine("--Elv Min: " + decimal.Round(min, 4) + " Max: " + decimal.Round(max, 4));

                    a = Convert.ToDecimal(r.NextDouble() * (0.07 - 0.0) + 0.0);


                    tw.WriteLine("--Nor Avg: " + (xlWorksheet.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value.ToString() + " SD: " + decimal.Round(a, 4));

                    a = Convert.ToDecimal(r.NextDouble() * (0.07 - 0.0) + 0.0);
                    tw.WriteLine("--Eas Avg: " + (xlWorksheet.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Value.ToString() + " SD: " + decimal.Round(a, 4));

                    a = Convert.ToDecimal(r.NextDouble() * (0.07 - 0.0) + 0.0);
                    tw.WriteLine("--Elv Avg: " + (xlWorksheet.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Value.ToString() + " SD: " + decimal.Round(a, 4));

                    a = decimal.Round(Convert.ToDecimal(r.NextDouble() * (0.3 - 0.0) + 0.0), 4);
                    tw.WriteLine("--NRMS Avg: " + a + " SD: 0.0000" + " Min: " + a + " Max: " + a);
                    decimal NRMS = decimal.Round(a, 3);
                    a = decimal.Round(Convert.ToDecimal(r.NextDouble() * (0.3 - 0.0) + 0.0), 4);
                    tw.WriteLine("--ERMS Avg: " + a + " SD: 0.0000" + " Min: " + a + " Max: " + a);
                    decimal ERMS = decimal.Round(a, 3);
                    a = decimal.Round(Convert.ToDecimal(r.NextDouble() * (0.3 - 0.0) + 0.0), 4);
                    tw.WriteLine("--HSDV Avg: " + a + " SD: 0.0000" + " Min: " + a + " Max: " + a);

                    a = decimal.Round(Convert.ToDecimal(r.NextDouble() * (0.3 - 0.0) + 0.0), 4);
                    tw.WriteLine("--VSDV Avg: " + a + " SD: 0.0000" + " Min: " + a + " Max: " + a);

                    a = decimal.Round(Convert.ToDecimal(r.NextDouble() * (4 - 0.5) + 0.5), 4);
                    tw.WriteLine("--HDOP Avg: " + a + " Min: " + (a - Convert.ToDecimal(0.0001)) + " Max: " + (a + Convert.ToDecimal(0.0001)));
                    decimal HDOP = decimal.Round(a, 3);
                    a = decimal.Round(Convert.ToDecimal(r.NextDouble() * (4 - 0.5) + 0.5), 4);
                    tw.WriteLine("--VDOP Avg: " + a + " Min: " + (a - Convert.ToDecimal(0.0001)) + " Max: " + (a + Convert.ToDecimal(0.0001)));
                    decimal VDOP = decimal.Round(a, 3);
                    a = decimal.Round(Convert.ToDecimal(r.NextDouble() * (4 - 0.5) + 0.5), 4);
                    tw.WriteLine("--PDOP Avg: " + a + " Min: " + (a - Convert.ToDecimal(0.0001)) + " Max: " + (a + Convert.ToDecimal(0.0001)));
                    decimal PDOP = decimal.Round(a, 3);
                    int BrojSatelitaAverage = r.Next(9, 16);

                    tw.WriteLine("--Number of Satellites Avg: " + BrojSatelitaAverage + " Min: " + (BrojSatelitaAverage - 1) + " Max: " + (BrojSatelitaAverage + 1));
                    a = Convert.ToDecimal(r.NextDouble() * (0.05 - 0.002) + 0.002);
                    decimal b = Convert.ToDecimal(r.NextDouble() * (0.05 - 0.002) + 0.002);
                    tw.WriteLine("--HSDV:" + decimal.Round(a, 3) + ", VSDV:" + decimal.Round(b, 3));
                    tw.WriteLine("SATS:" + BrojSatelitaAverage + ", PDOP:" + PDOP + ", HDOP:" + HDOP);
                    a = Convert.ToDecimal(r.NextDouble() * (0.4 - 0.2) + 0.2);
                    decimal g = decimal.Round(Convert.ToDecimal(r.NextDouble() * (0.4 - 0.2) + 0.2), 3);

                    tw.WriteLine("VDOP:" + VDOP + ", TDOP:" + (decimal.Round((HDOP + a), 3)) + ", GDOP:" + (decimal.Round((PDOP + g), 3)) + ", NSDV:" + NRMS + ", ESDV:" + ERMS);

                    TimeSpan provjera = new TimeSpan(08, 00, 00);

                    if (vrijeme == provjera)
                    {
                        TimeSpan rangeEnd = new TimeSpan(09, 00, 00);
                        TimeSpan rangeStart = provjera;
                        TimeSpan span = rangeEnd - rangeStart;

                        int randomMinutes = r.Next(0, (int)span.TotalMinutes);
                        vrijeme = rangeStart + TimeSpan.FromMinutes(randomMinutes);
                        int s = r.Next(0, 60);
                        TimeSpan se = new TimeSpan(00, 00, s);
                        vrijeme = vrijeme.Add(se);

                    }
                    else
                    {



                        int m = r.Next(1, 4);
                        int s = r.Next(0, 60);
                        TimeSpan ts = new TimeSpan(00, m, 00);
                        TimeSpan se = new TimeSpan(00, 00, s);
                        vrijeme = vrijeme.Add(ts);
                        vrijeme = vrijeme.Add(se);



                    }

                    TimeSpan krajenjVrijeme = new TimeSpan(16, 00, 00);
                    if (vrijeme > krajenjVrijeme)
                    {
                        TimeSpan rangeEnd = new TimeSpan(09, 00, 00);
                        TimeSpan rangeStart = provjera;
                        TimeSpan span = rangeEnd - rangeStart;

                        int randomMinutes = r.Next(0, (int)span.TotalMinutes);
                        vrijeme = rangeStart + TimeSpan.FromMinutes(randomMinutes);
                        int s = r.Next(0, 60);
                        TimeSpan se = new TimeSpan(00, 00, s);
                        vrijeme = vrijeme.Add(se);
                        datum = datum.AddDays(1);
                    }


                    tw.WriteLine("--DT" + datum.ToString("MM-dd-yyyy"));
                    tw.WriteLine("--TM" + vrijeme);


                }


            }

            xlWorkbook.Close(true);

            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlApp);
            MessageBox.Show("Završeno");
        }
    }
}
    

