using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace csv_to_form
{
        public partial class Form1 : Form
    {
        // txt посредник
        string addr2 = @"C:\temp_csv\test.txt";
        string addr3 = @"C:\temp_csv\Book1.csv";

        public Form1()
        {
            InitializeComponent();                        
        }

        private SqlConnection conn = null;
        private string source_csv;
        public string open_file(string dialogue) // диалог выбора файла
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = dialogue;
            ofd.InitialDirectory = @"c:\";
            ofd.Filter = "All files (*.*)|*.csv|All files (*.*)|*.*";
            ofd.FilterIndex = 1;
            ofd.RestoreDirectory = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                source_csv = ofd.FileName;
            }
            return source_csv;
        }

        // кнопка для открытия сохраненной таблицы
        private void button1_Click(object sender, EventArgs e)
        {   
            string new_addr1 = open_file("Выберите, из какого файла загрузить таблицу (формат CSV)"); // адрес открытого
            if (new_addr1 != null)          // если адрес не отменен
            {
                // перекодировка в кириллицу
                using (StreamReader sr = new StreamReader(new_addr1, Encoding.Default))
                {
                    // сброс в нужной кодировке посреднику
                    using (FileStream fs = File.Create(addr2))
                    {
                        while (sr.EndOfStream == false)
                        {
                            string data = sr.ReadLine();
                            byte[] info = new UTF8Encoding(true).GetBytes(data + Environment.NewLine);
                            fs.Write(info, 0, info.Length);
                        }
                    }
                }
                cll();  // разбивка на кадры из txt и склад в sql

                // из sql в таблицу Form1
                string query = "SELECT * FROM [Table] ORDER BY Id";
                SqlCommand cmd_grid = new SqlCommand(query, conn);
                SqlDataReader dr = cmd_grid.ExecuteReader();
                List<string[]> list = new List<string[]>();
                while (dr.Read())
                {
                    list.Add(new string[4]);

                    list[list.Count - 1][0] = dr[0].ToString();
                    list[list.Count - 1][1] = dr[1].ToString();
                    list[list.Count - 1][2] = dr[2].ToString();
                    list[list.Count - 1][3] = dr[3].ToString();
                }
                dr.Close();

                dataGridView1.Rows.Clear();
                foreach (string[] list2 in list)
                    dataGridView1.Rows.Add(list2);
            }
        }

        // разбивка на кадры из txt и склад в sql
        void cll()
        {
            using (TextFieldParser res = new TextFieldParser(addr2))
            {
                SqlCommand cmd_clr = new SqlCommand($"TRUNCATE TABLE [Table]", conn);
                cmd_clr.ExecuteNonQuery();

                res.TextFieldType = FieldType.Delimited;
                res.SetDelimiters(";");

                while (!res.EndOfData)
                {
                    string[] datas = res.ReadFields();

                    SqlCommand cmd_add = new SqlCommand($"INSERT INTO [Table] (Code, Name, Area) VALUES (N'{datas[0]}',N'{datas[1]}',N'{datas[2]}')", conn);
                    cmd_add.ExecuteNonQuery();
                }
            }
        }

         private void Form1_Load(object sender, EventArgs e)
        {
            // коннект к sql
            string way = System.Windows.Forms.Application.StartupPath; // текущая директория
            string addr_conn = @"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = "+way+@"Database1.mdf;Integrated Security=True";
            conn = new SqlConnection(addr_conn);

            conn.Open();

            // папка для посредника
            DirectoryInfo di = new DirectoryInfo(@"c:\temp_csv");
            if (!di.Exists)
            {
                di.Create();
            }

            if (!File.Exists(addr3))
            using (FileStream fs = File.Create(addr3))
            {

            }
        }

        // кнопка сохранения
        private void button2_Click(object sender, EventArgs e)
        {
            Saver();
        }

        // сохраняем из dataGrid в Excel
        public void Saver()
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;

            //Книга
            string excl_file = open_file("Выберите файл, в который хотите сохранить таблицу (формат CSV)");

            if (excl_file != null)
            {
                ExcelWorkBook = ExcelApp.Workbooks.Open(excl_file);

                //Таблица
                ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
                ExcelWorkSheet.Cells.ClearContents();

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        ExcelApp.Cells[i + 1, j + 1].EntireColumn.NumberFormat = "@";  // все как текст
                        ExcelApp.Cells[i + 1, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                    }
                }
                //Вызываем excel файл
                ExcelApp.Visible = true;
                ExcelApp.UserControl = true;
            }
        }
    }
}
