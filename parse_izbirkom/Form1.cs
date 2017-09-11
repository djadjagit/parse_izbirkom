using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;


namespace parse_izbirkom
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;
            openFileDialog1.ShowDialog();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            WebClient wc = new WebClient();
            string ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=rezult_vybor.accdb";
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = ConnectionString;
            conn.Open();
            OleDbDataAdapter rez = new OleDbDataAdapter();
            DateTime t1 = DateTime.Now;
            int[] a = new int[18];
            string[] adr = File.ReadAllLines(openFileDialog1.FileName);
            Excel.Application excelApp = new Excel.Application();
            for (int k = 0; k < adr.Length; k++)
            {
                //                wc.DownloadFile("http://www.novgorod.vybory.izbirkom.ru/servlet/ExcelReportVersion?region=53&sub_region=53&root=534018021&global=null&vrn=4534018180415&tvd=4534018180426&type=427&vibid=4534018180426&condition=&action=show&version=null&prver=0&sortorder=1", "report.xls");
                if (adr[k]!="")
                {
                    wc.DownloadFile(adr[k], "report.xls");
                    while (!File.Exists("report.xls"))
                    {
                        Text = "Файл report.xls еще не получен. Прошло времени:" + (DateTime.Now - t1).TotalSeconds.ToString();
                        Thread.Sleep(1);
                    }
                    Excel.Workbook workBook = excelApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "report.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
                    Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Sheets[1];
                    //            excelApp.Visible = true;
                    string rayon = workSheet.Cells[4, 1].Text.ToString();
                    rayon = rayon.Remove(0, rayon.IndexOf("    ") + 4);
                    label1.Text = rayon;
                    int j = 4;
                    string uik = workSheet.Cells[8, j].Text.ToString();
                    while (uik != "")
                    {
                        uik = uik.Remove(0, uik.IndexOf("№") + 1);
                        a[0] = Convert.ToInt16(uik);
                        for (int i = 9; i < 21; i++)
                        {
                            if (workSheet.Cells[i, j].Text.ToString() != "")
                                a[i - 8] = Convert.ToInt16(workSheet.Cells[i, j].Text.ToString());
                            else
                                a[i - 8] = 0;
                        }
                        int ind = 13;
                        for (int i = 22; i < 32; i = i + 2)
                        {
                            if (workSheet.Cells[i, j].Text.ToString() != "")
                                a[ind] = Convert.ToInt16(workSheet.Cells[i, j].Text.ToString());
                            else
                                a[ind] = 0;
                            ind++;
                        }
                        try
                        {
                            string sql = "insert into rezult ([rayon], [uik], [p1], [p2], [p3], [p4], [p5], [p6], [p7], [p8], [p9], " +
                            "[p10], [p11], [p12], [k1_c], [k1_p], [k2_c], [k2_p], [k3_c], [k3_p], [k4_c], [k4_p], [k5_c], [k5_p]) " +
                            "values ('" + rayon + "', " + a[0] + ", " + a[1] + ", " + a[2] + ", " + a[3] + ", " + a[4] + ", " + a[5] +
                            ", " + a[6] + ", " + a[7] + ", " + a[8] + ", " + a[9] + ", " + a[10] + ", " + a[11] + ", " + a[12] +
                            ", " + a[13] + ", " + ((float)a[13] / (float)(a[9] + a[10])) * 100 + ", " + a[14] + ", " +
                            ((float)a[14] / (float)(a[9] + a[10])) * 100 + ", " + a[15] + ", " + ((float)a[15] / (float)(a[9] + a[10])) * 100 +
                            ", " + a[16] + ", " + ((float)a[16] / (float)(a[9] + a[10])) * 100 + ", " + a[17] + ", " +
                            ((float)a[17] / (float)(a[9] + a[10])) * 100 + ")";
                            rez.InsertCommand = new OleDbCommand(sql, conn);
                            rez.InsertCommand.ExecuteScalar();
                            Text = (j - 3).ToString() + "; Прошло времени:" + (DateTime.Now - t1).TotalSeconds.ToString();
                        }
                        catch (Exception err)
                        {
                            Text = err.Message;
                        }
                        Thread.Sleep(10);
                        j++;
                        uik = workSheet.Cells[8, j].Text.ToString();
                    }
                    workBook.Close(false, Type.Missing, Type.Missing);
                }
            }
            excelApp.Quit();
            conn.Close();
            label1.Text = "Завершено. Прошло времени:" + (DateTime.Now - t1).TotalSeconds.ToString();
            Text = "Данные избиркома";
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            timer1.Enabled = true;
        }
    }
}
