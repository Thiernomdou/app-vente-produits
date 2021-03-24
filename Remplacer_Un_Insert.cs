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
using Oracle.DataAccess.Client;
using System.Configuration;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Data.SQLite;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using Microsoft.VisualBasic;

namespace Calage_Inserts
{
    public partial class Remplacer_Un_Insert : Form
    {
        string connString = @"Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=fra2exa01-sxdir1-vip.europe.essilor.group)(PORT=1561)))
                             (CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=PUE1)));User Id=combo;Password=combo;";
        public Remplacer_Un_Insert()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Accueil r = new Accueil("");
            r.Show();
        }


        public class donnees
        {
            public donnees() { }

            public string numero { get; set; }
           

        }

        private void button3_Click(object sender, EventArgs e)
        {
            var product = "";
            var ins_cv = "";
            var eye = "";
            var LineIndex = "";
            var ColumnIndex = "";

            //CHERCHER LE JOB POUR PRE-REMPLIR LES AUTRES 

            OracleConnection conn = new OracleConnection(connString);
            try
            {
                conn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = conn;
                cmd.CommandText = "SELECT COMBO_JOB_LINES.CAVITY_JOB_NB, COMBO_JOB_LINES.LB_LOGI_SKU, COMBO_ITEMS.COLUMN_INDEX, COMBO_ITEMS.LINE_INDEX, COMBO_ITEMS.EYE, COMBO_BOM.TYPE_INS_CV, COMBO_ITEMS.PRODUCT, COMBO_ITEMS.DIAMETER FROM COMBO_JOB_LINES, COMBO_ITEMS, COMBO_BOM where COMBO_ITEMS.LB_LOGI_SKU = COMBO_JOB_LINES.LB_LOGI_SKU and COMBO_BOM.CD_CCE_SKU = COMBO_ITEMS.CD_CCE_SKU and JOB_NB = '"+textBox1.Text+"' and COMBO_JOB_LINES.CAVITY_JOB_NB = '"+comboBox1.Text+"' ORDER BY COMBO_JOB_LINES.CAVITY_JOB_NB";
                MessageBox.Show(cmd.CommandText);
                OracleDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    ColumnIndex = reader.GetString(2);
                    LineIndex = reader.GetString(3);
                    eye = reader.GetString(4);
                    ins_cv = reader.GetString(5);
                    product = reader.GetString(6);
                   

                }
                // MessageBox.Show(product);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            conn.Close();

            string myConnectionString = @"user id=sa; password=K@rdexlsadm21!; data source=FRESD32615\SQLEXPRESS";
            SqlConnection myConnection = new SqlConnection(myConnectionString);

            
            ///////////////////////////////////////////
            //TROUVER LE PRODUIT
            /////////////////////////////////////////
            var Nom_Produit = "";
            

            string strRequette1 = "SELECT [Description_produit] FROM[ACI].[dbo].[Code_produit] where Product = '" + product + "' ";
            MessageBox.Show(strRequette1);
            myConnection.Open();
            SqlCommand myCommandd = new SqlCommand(strRequette1, myConnection);
            SqlDataReader mySqlDataReader = myCommandd.ExecuteReader();
            List<donnees> list = new List<donnees>();
            while (mySqlDataReader.Read())
            {
                Nom_Produit = mySqlDataReader.GetString(0);
                MessageBox.Show(Nom_Produit);
            }
                myConnection.Close();

            if (Nom_Produit == "SPHERIQUE" || Nom_Produit == "ASPHERIQUE")
            {
                MessageBox.Show("c'est du progressif");
            }
            else
            {
                ///////////////////////////////////////////
                //TOUT LES INSERTS POUR LE NON PROGESSIF
                /////////////////////////////////////////
                string strRequete1 = "SELECT  [Numero],[Produit],[Base1],[Base2],[Addition],[Oeil],[CCCX],[Glass],[Hauteur_Centre],[Hauteur_Bord] FROM [ACI].[dbo].[Inserts] where Glass = '" + ins_cv + "' and Oeil = '" + eye + "' and Base1 = " + LineIndex + " and Addition = " + ColumnIndex + " and CCCX = '"+comboBox2.Text+"' and Produit = '" + Nom_Produit + "' ";
                MessageBox.Show(strRequete1);
                myConnection.Open();
                SqlCommand myCommand1 = new SqlCommand(strRequete1, myConnection);
                SqlDataReader mySqllDataReader = myCommand1.ExecuteReader();
               
                while (mySqllDataReader.Read())
                {
                    list.Add(new donnees
                    {
                        numero = mySqllDataReader.GetString(0)
                    });
                }
                foreach (donnees mesdonnees in list)
                {
                    comboBox3.Text = mesdonnees.numero;
                }
                
                    myConnection.Close();
            }

            
            

        }


        private void button2_Click(object sender, EventArgs e)
        {
            //Remplacer un insert

            


            /*
            


            */
        }

        
    }
}
