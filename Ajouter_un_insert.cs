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
    public partial class Ajouter_un_insert : Form
    {
        public Ajouter_un_insert()
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
            //Ajouter un insert
            string myConnectionString = @"user id=sa; password=K@rdexlsadm21!; data source=FRESD32615\SQLEXPRESS";
            SqlConnection myConnection = new SqlConnection(myConnectionString);

            try
            {
                myConnection.Open();
                string strRequete = @"INSERT INTO [ACI].[dbo].[Inserts] (Numero,Produit,Base1,Base2,Addition,Oeil,CCCX,Glass,Hauteur_Centre,Hauteur_Bord)VALUES(@Numero,@Produit,@Base1,@Base2,@Addition,@Oeil,@CCCX,@Glass,@Hauteur_Centre,@Hauteur_Bord); ";
                SqlCommand myCommand = new SqlCommand(strRequete, myConnection);
                myCommand.Parameters.AddWithValue("@Numero", Numero.Text);
                myCommand.Parameters.AddWithValue("@Produit", Produit.Text);
                myCommand.Parameters.AddWithValue("@Base1", Base1.Text);
                myCommand.Parameters.AddWithValue("@Base2", Base2.Text);
                myCommand.Parameters.AddWithValue("@Addition", Addition.Text);
                myCommand.Parameters.AddWithValue("@Oeil", Oeil.Text);
                myCommand.Parameters.AddWithValue("@CCCX", CCCX.Text);
                myCommand.Parameters.AddWithValue("@Glass", Glass.Text);
                myCommand.Parameters.AddWithValue("@Hauteur_Centre", Hauteur_Centre.Text);
                myCommand.Parameters.AddWithValue("@Hauteur_Bord", Hauteur_Bord.Text);
                SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                    
                    if (Numero.Text == "")
                    {
                        MessageBox.Show("Veuillez entrez le Numéro de l'insert");
                    }
                    else if (Produit.Text == "")
                    {
                        MessageBox.Show("Veuillez entrez le Produit de l'insert");
                    }
                    else if (Base1.Text == "")
                    {
                        MessageBox.Show("Veuillez entrez la Base N°1 de l'insert");
                    }
                    else if (Base2.Text == "")
                    {
                        MessageBox.Show("Veuillez entrez la Base N°2 de l'insert");
                    }
                    else if (CCCX.Text == "")
                    {
                        MessageBox.Show("Veuillez entrez le CCCX de l'insert");
                    }
                    else if (Glass.Text == "")
                    {
                        MessageBox.Show("Veuillez entrez la Base N°1 de l'insert");
                    }
                    else if (Hauteur_Centre.Text == "")
                    {
                        MessageBox.Show("Veuillez entrez la Hauteur Centre de l'insert");
                    }
                    else if (Hauteur_Bord.Text == "")
                    {
                        MessageBox.Show("Veuillez entrez la Hauteur Bord de l'insert");
                    }
                    else
                    {
                        MessageBox.Show("Vous avez inséré l'insert " + Numero.Text);
                        Numero.Text = ""; Produit.Text = ""; Base1.Text = ""; Base2.Text = ""; Addition.Text = ""; Oeil.Text = "";
                        CCCX.Text = ""; Glass.Text = ""; Hauteur_Centre.Text = ""; Hauteur_Bord.Text = "";

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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
 
        }
    }
}
