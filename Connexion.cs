using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using System.IO;
using Cake.Core.IO;

namespace Calage_Inserts
{
    public partial class Connexion : Form
    {
        
       
        public Connexion()
        {
            InitializeComponent();
           
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        

       

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox1.BackColor = Color.White;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            //conditions de remplissage du champ login
            if (textBox2.Text.Trim() == "" && textBox1.Text.Trim() == "")
            {
                MessageBox.Show("Veuillez renseigner le champ login et password");
                textBox2.BackColor = Color.Red;
                textBox1.BackColor = Color.Red;
                return;
            }
            else
            {
                string query = "SELECT * FROM Connexion WHERE Login = @user AND Password = @pass";
                SQLiteConnection conn = new SQLiteConnection("Data Source=inserts.db;");
                conn.Open();
                SQLiteCommand cmd = new SQLiteCommand(query, conn);
                cmd.Parameters.AddWithValue("@user", textBox1.Text);
                cmd.Parameters.AddWithValue("@pass", textBox2.Text);
                SQLiteDataAdapter da = new SQLiteDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);


                if (dt.Rows.Count > 0)
                {
                    this.Hide();
                    if(dt.Rows[0][0].ToString() == "1")
                    {
                        Accueil a = new Accueil(dt.Rows[0][1].ToString());
                        a.Show();
                    }
                    else if(dt.Rows[0][0].ToString() == "2")
                    {
                        Accueil1 a = new Accueil1(dt.Rows[0][1].ToString());
                        a.Show();
                    }
                    else if (dt.Rows[0][0].ToString() == "3")
                    {
                        Accueil2 a = new Accueil2(dt.Rows[0][1].ToString());
                        a.Show();
                    }
                }
                else
                {
                    MessageBox.Show("Login ou Password incorrect");
                    textBox1.BackColor = Color.Red;
                    textBox2.BackColor = Color.Red;
                    return;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {
            textBox2.BackColor = Color.White;
        }
    }
}
