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
using System.IO;

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
                //Mise à jour de la base de données
                var numero = 0;
                myConnection.Open();
                string strRequete3 = @"SELECT  max ([num_max]) + 1 FROM [ACI].[dbo].[numero_max]";
                //MessageBox.Show(strRequete3);
                SqlCommand myCommand3 = new SqlCommand(strRequete3, myConnection);
                SqlDataReader mySqDataReader3 = myCommand3.ExecuteReader();

                while (mySqDataReader3.Read())
                { 
                    numero = mySqDataReader3.GetInt32(0);
                }

                    myConnection.Close();

                myConnection.Open();
                string strRequete2 = @"insert into[ACI].[dbo].[numero_max] values("+numero+") ";
                SqlCommand myCommand2 = new SqlCommand(strRequete2, myConnection);
                SqlDataReader mySqDataReader2 = myCommand2.ExecuteReader();
                myConnection.Close();

                myConnection.Open();
                string strRequete = @"INSERT INTO [ACI].[dbo].[Inserts] (Numero,Produit,Base1,Base2,Addition,Oeil,CCCX,Glass,Hauteur_Centre,Hauteur_Bord)VALUES(@Numero,@Produit,@Base1,@Base2,@Addition,@Oeil,@CCCX,@Glass,@Hauteur_Centre,@Hauteur_Bord); ";
                SqlCommand myCommand = new SqlCommand(strRequete, myConnection);
                myCommand.Parameters.AddWithValue("@Numero", label13.Text);
                myCommand.Parameters.AddWithValue("@Produit", Produit.Text);
                myCommand.Parameters.AddWithValue("@Base1", Base1.Text);
                myCommand.Parameters.AddWithValue("@Base2", Base2.Text);
                myCommand.Parameters.AddWithValue("@Addition", Addition.Text);
                myCommand.Parameters.AddWithValue("@Oeil", Oeil.Text);
                myCommand.Parameters.AddWithValue("@CCCX", CCCX.Text);
                myCommand.Parameters.AddWithValue("@Glass", Glass.Text);
                myCommand.Parameters.AddWithValue("@Hauteur_Centre", Hauteur_Centre.Text);
                myCommand.Parameters.AddWithValue("@Hauteur_Bord", Hauteur_Bord.Text);
                SqlDataReader mySqDataReader1 = myCommand.ExecuteReader();
                    
                    if (label13.Text == "")
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
                        MessageBox.Show("Vous avez inséré l'insert " + label13.Text);
                    label13.Text = ""; Produit.Text = ""; Base1.Text = ""; Base2.Text = ""; Addition.Text = ""; Oeil.Text = "";
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


            //création d'un fichier .dat

           


            string fileName = @"R:\Commun\ACI\Autre\" + Numero_ins.Text + ".dat";
            try
            {

                // Créer un nouveau fichier   
                using (FileStream fileStr = File.Create(fileName))
                {
                    // Ajouter du texte au fichier  
                    Byte[] text = new UTF8Encoding(true).GetBytes("INSSERIAL;" + Numero_ins.Text +";" + label13.Text +";1");
                    fileStr.Write(text, 0, text.Length);
                }

               
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            Numero_ins.Text = "";
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
 
        }

        private void Produit_SelectedIndexChanged(object sender, EventArgs e)
        {
            string myConnectionString = @"user id=sa; password=K@rdexlsadm21!; data source=FRESD32615\SQLEXPRESS";
            SqlConnection myConnection = new SqlConnection(myConnectionString);

            var numero = 0;

            try
            {
                myConnection.Open();
                string strRequete = @"SELECT  max ([num_max]) + 1 FROM [ACI].[dbo].[numero_max]";
                //MessageBox.Show(strRequete);
                SqlCommand myCommand = new SqlCommand(strRequete, myConnection);
                SqlDataReader mySqDataReader = myCommand.ExecuteReader();
                while (mySqDataReader.Read())
                {
                   

                    if (Produit.Text == "ROCKY")
                    {
                        numero = mySqDataReader.GetInt32(0);
                        label13.Text = "R" + numero.ToString();
                    }
                    
                    else if (Produit.Text == "SPHERIQUE")
                    {
                        numero = mySqDataReader.GetInt32(0);

                        label13.Text = "S" + numero.ToString();
                    }
                    else if (Produit.Text == "ASPHERIQUE")
                    {
                        numero = mySqDataReader.GetInt32(0);

                        label13.Text = "S" + numero.ToString();
                    }
                    else if (Produit.Text == "CONDE")
                    {
                        numero = mySqDataReader.GetInt32(0);

                        label13.Text = "K" + numero.ToString();
                    }
                    else if (Produit.Text == "SNL")
                    {
                        numero = mySqDataReader.GetInt32(0);

                        label13.Text = "SN" + numero.ToString();
                    }
                    else if (Produit.Text == "GX")
                    {
                        numero = mySqDataReader.GetInt32(0);

                        label13.Text = "GX" + numero.ToString();
                    }
                    else if (Produit.Text == "OVATION")
                    {
                        numero = mySqDataReader.GetInt32(0);
                        
                        label13.Text = "E" + numero.ToString();
                    }
                    
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

        private void Ajouter_un_insert_Load(object sender, EventArgs e)
        {
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }
    }
}
