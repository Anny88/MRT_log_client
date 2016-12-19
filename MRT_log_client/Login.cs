using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace MRT_log_client
{
    public partial class Login : Form
    {
        private Form parent_form;
        public Login(Form parent)
        {
            InitializeComponent();
            parent_form = parent;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string constring = "datasource=localhost;port=3306;username="+login_textbox.Text+";password="+pass_textbox.Text;
                MySqlConnection conDataBase = new MySqlConnection(constring);
                conDataBase.Open();
                DataSet ds = new DataSet();
                MessageBox.Show("Connected");
                conDataBase.Close();
                ((Form2)parent_form).set_con(login_textbox.Text,pass_textbox.Text);
                this.Close();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
       
    }
}
