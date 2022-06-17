using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Newtonsoft.Json;
using System.IO;

namespace KKC1_Razliv
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            dataGridView2.DataSource = ReadJSON();
        }

        //Метод для считывания схемы файла xlsx
        public DataTable loadExcelScheme()
        {
            string dataSource = textBox1.Text;
            string conneсtionString = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + dataSource + ";" + "Extended Properties='Excel 8.0;HDR=Yes'";
            OleDbConnection cnn = new OleDbConnection(conneсtionString);
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = cnn;

            cnn.Open();

            DataTable dtExcelSchema;
            dtExcelSchema = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            comboBox1.DataSource = dtExcelSchema;
            comboBox1.DisplayMember = "TABLE_NAME";
            string sheetName = comboBox1.GetItemText(this.comboBox1.SelectedItem);
            string sql = $"SELECT * FROM [{sheetName}]";
            DataTable dt = new DataTable();
            OleDbDataAdapter da;
            da = new OleDbDataAdapter(sql, cnn);
            da.Fill(dt);

            cnn.Close();

            return dt;

        }

        //Считываем файл через метод и диалоговое окно
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                CheckFileExists = true,
                CheckPathExists = true,
            };
            
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }

           dataGridView1.DataSource = loadExcelScheme();
        }

        //Метод для считывания листа из файла
        public DataTable loadSheetFromOutput()
        {
            string sheetName = comboBox1.GetItemText(this.comboBox1.SelectedItem);
            string sql = $"SELECT * FROM [{sheetName}]";

            string conneсtionString = null;
            string dataSource = textBox1.Text;
            conneсtionString = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + dataSource + ";" + "Extended Properties='Excel 8.0;HDR=Yes'"; ;
            OleDbConnection cnn = new OleDbConnection(conneсtionString);
            cnn.Open();

            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(sql, cnn);
            da.Fill(dt);
            cnn.Close();

            return dt;
        }
        // Обновление dataGridView при смене item в combobox
        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            dataGridView1.DataSource = loadSheetFromOutput();
        }
        //Метод для конвертации json в datatable
        public DataTable ReadJSON()
        {
            DataTable dt = new DataTable();
            string json = File.ReadAllText("C:\\Users\\kuptsov_ae\\source\\repos\\KKC1_Razliv\\KKC1_Razliv\\Filec_0.json");
            dt = JsonConvert.DeserializeObject<DataTable>(json);
            return dt;
        }
        
        //Метод для добавления данных в json
        public void WriteJSON(DataTable dt)
        {
            File.WriteAllText("C:\\Users\\kuptsov_ae\\source\\repos\\KKC1_Razliv\\KKC1_Razliv\\Filec_0.json", JsonConvert.SerializeObject(dt));
        }
        //Расчет
        private void button3_Click(object sender, EventArgs e)
        {
            DataTable dt2 = new DataTable();
            dt2 = ReadJSON();

                        DataRow insert = dt2.NewRow();
            insert["NPLV"] = textBox2.Text;
            insert["VES_BEF"] = textBox3.Text;
            insert["VIDERZ_H"] = textBox4.Text;
            insert["VIDERZ_M"] = textBox5.Text;
            insert["PREDICT"] = Convert.ToDouble(textBox2.Text);
            dt2.Rows.Add(insert);
            WriteJSON(dt2);
            dataGridView2.DataSource = ReadJSON();

           
        }
    }
}
