using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using Microsoft.Office.Interop.Word;

namespace MRT_log_client
{
    public partial class Form2 : Form
    {
        private string login;
        private string pass;
        
        
        
        public Form2()
        {
            InitializeComponent();
            
            (new Login(this)).Show();
        }
        public void set_con(string new_login, string new_pass)
        {
            login = new_login;
            pass = new_pass;
            
            
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if ((textBox1.Text != "")&&(textBox2.Text!="")&&(textBox3.Text != ""))
            {
                string constring = "datasource=localhost;port=3306;username=" + login + ";password=" + pass;
                MySqlConnection connection = new MySqlConnection(constring);
                string query = "insert into test_database.test_data (name,doctor,price,admin) values('"+textBox1.Text+"',"+textBox2.Text+","+textBox3.Text+",'"+login+"') ;";
                MySqlCommand cmdDB = new MySqlCommand(query, connection);
                MySqlDataReader myReader;
                try
                {
                    connection.Open();
                    myReader = cmdDB.ExecuteReader();
                    MessageBox.Show("Saved");
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                    connection.Close();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string constring = "datasource=localhost;port=3306;username=" + login + ";password=" + pass;
            MySqlConnection connection = new MySqlConnection(constring);
            string query = "SELECT * FROM test_database.test_data;";
            MySqlCommand cmdDB = new MySqlCommand(query, connection);
            MySqlDataReader myReader;
            try
            {
                connection.Open();
                myReader = cmdDB.ExecuteReader();
                Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();
                
                winword.Visible = false;
                object missing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing,ref missing);
                document.Content.SetRange(0, 0);
                document.Content.Text = "Статистика по врачу №" + doc_req.Text + Environment.NewLine;
                Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
                Table stats = document.Tables.Add(para1.Range,1,3);
                stats.Borders.Enable = 1;
                document.Tables[1].Rows.Add(document.Tables[1].Rows[1]);
                document.Tables[1].Cell(1, 1).Range.Text = "ФИО пациента";
                document.Tables[1].Cell(1, 2).Range.Text = "Дата посещения";
                document.Tables[1].Cell(1, 3).Range.Text = "Стоимость услуги";
                int k = 1;
                while (myReader.Read())
                {
                    if (Convert.ToString(myReader["doctor"])==doc_req.Text) {
                        k++;
                        document.Tables[1].Rows.Add(document.Tables[1].Rows[k]);
                        document.Tables[1].Cell(k, 1).Range.Text = Convert.ToString(myReader["name"]);
                        document.Tables[1].Cell(k, 2).Range.Text = Convert.ToString(myReader["date"]);
                        document.Tables[1].Cell(k, 3).Range.Text = Convert.ToString(myReader["price"]);

                    }
                }
                object filename = @"d:\report_"+doc_req.Text+".docx";
                document.SaveAs2(ref filename);
                document.Close(ref missing,ref missing,ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
                MessageBox.Show("Report created");
                connection.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
