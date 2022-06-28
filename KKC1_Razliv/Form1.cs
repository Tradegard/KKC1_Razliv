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

            comboBox2.DataSource = dt;
            comboBox2.DisplayMember = "NPLV";

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
            conneсtionString = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + dataSource + ";" + "Extended Properties='Excel 8.0;HDR=Yes'";
            OleDbConnection cnn = new OleDbConnection(conneсtionString);
            cnn.Open();

            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(sql, cnn);
            da.Fill(dt);
            cnn.Close();

            comboBox2.DataSource = dt;
            comboBox2.DisplayMember = "NPLV";

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
            if (textBox1.Text == "")
            {
                MessageBox.Show("Выберите файл!");
            }
            else
            {

                DataTable dt2 = new DataTable();
                DataTable dt3 = new DataTable();
                dt2 = loadSheetFromOutput();
                dt3 = ReadJSON();
                var nplv = 0;
                var vesbef = 0;
                var viderz_h = 0;
                var viderz_m = 0;
                nplv = Convert.ToInt32(comboBox2.GetItemText(this.comboBox2.SelectedItem));
                if (textBox3.Text == "") { MessageBox.Show("Введите Вес До!"); } else if (int.TryParse(textBox3.Text, out vesbef)) { vesbef = Convert.ToInt32(textBox3.Text); };
                if (textBox4.Text == "") { MessageBox.Show("Введите Часы Выдержки!"); } else if (int.TryParse(textBox4.Text, out viderz_h)) { viderz_h = Convert.ToInt32(textBox4.Text); };
                if (textBox5.Text == "") { MessageBox.Show("Введите Минуты Выдержки!"); } else if (int.TryParse(textBox5.Text, out viderz_m)) { viderz_m = Convert.ToInt32(textBox5.Text); };        
                var viderz = viderz_h + viderz_m / 60;
                string itemOUT = null;

                               
                DataRow[] result = dt2.Select($"NPLV = {nplv}");
                foreach (DataRow row in result)
                {
                    var kolplav = row["KOL_PLAV"].Equals(DBNull.Value) ? 0 : Convert.ToInt32(row["KOL_PLAV"]);                    
                    var sostnum = row["SOST_NUM"].Equals(DBNull.Value) ? 0 : Convert.ToInt32(row["SOST_NUM"]);
                    var veskant = row["VES_KANT"].Equals(DBNull.Value) ? 0 : Convert.ToDouble(row["VES_KANT"]);
                    var vipusk = row["VIPUSK"].Equals(DBNull.Value) ? 0 : Convert.ToDouble(row["VIPUSK"]);
                    var doduv = row["DODUV"].Equals(DBNull.Value) ? 0 : Convert.ToInt32(row["DODUV"]);
                    var izvest = row["IZVEST"].Equals(DBNull.Value) ? 0 : Convert.ToDouble(row["IZVEST"]);
                    var firstzamer = row["FIRST_ZAMER"].Equals(DBNull.Value) ? 0 : Convert.ToDouble(row["FIRST_ZAMER"]);
                    var temprazliv = row["TEMP_RAZLIV"].Equals(DBNull.Value) ? 0 : Convert.ToDouble(row["TEMP_RAZLIV"]);
                    double predict;
                    if (sostnum == 1)
                    {
                        predict = -108.745 + vipusk * 0.2852 + doduv * 0.4698 + kolplav * 0.0512 + izvest * (-0.0013) + firstzamer * 0.0638 + viderz * (-3.3892) + vesbef * 0.9229 + veskant * (-0.5556);
                    }
                    else if (sostnum == 2 || sostnum == 3)
                    {
                        predict = 56.9146 + vesbef * 0.9216 + veskant * (-0.6847) + izvest * (-0.0007) + vipusk * 0.2548 + viderz * (-1.7444) + temprazliv * (-0.0389) + sostnum * 0.0999 + doduv * 0.6079;
                    }
                    else
                    {
                        predict = 57.1838 + vesbef * 0.9186 + veskant * (-0.6814) + izvest * (-0.0008) + vipusk * 0.2301 + viderz * (-1.6068) + temprazliv * (-0.0383) + sostnum * 0.1223;
                    }

                    itemOUT = predict.ToString();                    
                }
                
                label6.Text = itemOUT;

                
                DataRow insert = dt3.NewRow();
                insert["NPLV"] = comboBox2.GetItemText(this.comboBox2.SelectedItem);
                insert["VES_BEF"] = textBox3.Text;
                insert["VIDERZ_H"] = textBox4.Text;
                insert["VIDERZ_M"] = textBox5.Text;
                insert["PREDICT"] = Convert.ToDouble(itemOUT);
                dt3.Rows.Add(insert);
                WriteJSON(dt3);
                dataGridView2.DataSource = ReadJSON();
                
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                dataGridView1.DataSource = loadSheetFromOutput();
            }                
        }
    }
}
