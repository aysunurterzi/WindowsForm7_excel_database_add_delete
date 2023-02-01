using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
using OfficeOpenXml;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {

        SqlConnection baglanti;
        SqlCommand komut;
        SqlDataAdapter dataa;
        public Form1()
        {
            InitializeComponent();

        }
        
        void kullanıcı_getir()
        {
            baglanti = new SqlConnection("Data Source=IPKMRKNB082\\SQLEXPRESS;Initial Catalog=kullanici;Integrated Security=True");
            baglanti.Open();
            dataa = new SqlDataAdapter("Select * From tablo1 Order by tcno", baglanti);
            System.Data.DataTable tablo1 = new System.Data.DataTable();
            dataa.Fill(tablo1);
            dataGridView1.DataSource = tablo1;


        }

        private void Form1_Load(object sender, EventArgs e)
        {
          
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string kayıtetsorgu = "insert into tablo1 (tcno,kullaniciadi,sifre) values(@tcno,@kullaniciadi,@sifre)";
                komut = new SqlCommand(kayıtetsorgu, baglanti);
                komut.Parameters.AddWithValue("@tcno", textBox1.Text);
                komut.Parameters.AddWithValue("@kullaniciadi", textBox2.Text);
                komut.Parameters.AddWithValue("@sifre", textBox3.Text);
                komut.ExecuteNonQuery();
                baglanti.Close();
                kullanıcı_getir();
            }

            catch (Exception)
            {
                MessageBox.Show("Tc numaranız kayıtlı kullanıcılarla aynı olamaz");
              
            }
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            dataGridView1.ClearSelection();


        }

        private void button4_Click(object sender, EventArgs e)
        {
            string silsorgu = "delete from tablo1 where tcno=@tcno";
            komut = new SqlCommand(silsorgu, baglanti);
            komut.Parameters.AddWithValue("@tcno", Convert.ToDecimal(textBox1.Text));
            komut.ExecuteNonQuery();
            baglanti.Close();
            kullanıcı_getir();

            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            dataGridView1.ClearSelection();
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string güncellesorgu = "Update tablo1 set tcno=@tcno,kullaniciadi=@kullaniciadi,sifre=@sifre where tcno=@tcno";
                komut = new SqlCommand(güncellesorgu, baglanti);
                komut.Parameters.AddWithValue("@tcno", Convert.ToDecimal(textBox1.Text));
                komut.Parameters.AddWithValue("@kullaniciadi", textBox2.Text);
                komut.Parameters.AddWithValue("@sifre", textBox3.Text);
                komut.ExecuteNonQuery();
                baglanti.Close();
                kullanıcı_getir();

            }
            catch (Exception)
            {
                MessageBox.Show("Tc numaranız kayıtlı kullanıcılarla aynı olamaz");

            }
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            dataGridView1.ClearSelection();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            kullanıcı_getir();

            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            dataGridView1.ClearSelection();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            dataGridView1.ClearSelection();

        }
      
        private void button6_Click(object sender, EventArgs e)
        {

            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Workbook workbook = excel.Workbooks.Add(Missing);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];

            int StartCol = 1;
            int StartRow = 1;

            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                Range myRange = (Range)sheet1.Cells[StartRow, StartCol + j];
                myRange.Value2 = dataGridView1.Columns[j].HeaderText;
            }
            StartRow++;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {

                    Range myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                    myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                    myRange.Select();
                }
            }
            workbook.SaveAs(@"C:\Users\aysu.terzi\Desktop\datalar.xls", XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excel.Quit();

            MessageBox.Show("Excel dosyası Masaüstüne kaydedildi");
          

            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "Metin Dosyası|*.txt";
            save.OverwritePrompt = true;
            save.CreatePrompt = true;

            if (save.ShowDialog() == DialogResult.OK)
            {
                StreamWriter Kayit = new StreamWriter(save.FileName);
                Kayit.WriteLine("Excel dosyasını indiren Kullanıcının adı: "+Environment.UserName + Environment.NewLine + "Excel dosyasının kaydedildiği zaman: " +DateTime.Now);
                Kayit.Close();
            }
            MessageBox.Show("Text dosyası kaydedildi");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
