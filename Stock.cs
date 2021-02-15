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

namespace Calage_Inserts
{
    public partial class Stock : Form
    {
        SqlConnection myConnection;
        SqlCommand myCommand;
        string myConnectionString;
        string strRequete;
        public Stock()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Accueil r = new Accueil("");
            r.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Connexion à la base de données CARDEX
            myConnectionString = "user id=sa; password=K@rdexlsadm21!; data source=FRESD32615'\'SQLEXPRESS";
            strRequete = "select * from [PPG].[dbo].[Materialproperty]";

            try
            {
                myConnection = new SqlConnection(myConnectionString);
                myConnection.Open();
                myCommand = new SqlCommand(strRequete, myConnection);
                SqlDataReader mySqDataReader = myCommand.ExecuteReader();

                dataGridView1.Rows.Clear();
                DataSet dataset = new DataSet();

                while (mySqDataReader.Read())
                {
                    dataGridView1.Rows.Add(mySqDataReader[0], mySqDataReader[1]);
                }
            }
            catch (Exception eMsg1)
            {
                Console.WriteLine("Erreur durant l’execution de la requete : " + eMsg1.Message);
            }
            finally
            {
                myConnection.Close();
            }


        }
    }
}
