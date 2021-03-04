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
    public partial class Extraction_Jobs : Form
    {

        string connString = @"Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=fra2exa01-sxdir1-vip.europe.essilor.group)(PORT=1561)))
                             (CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=PUE1)));User Id=combo;Password=combo;";
        public Extraction_Jobs()
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
            // création d'entrées TNS  



            /*
                        PUE1.WORLD =
             (DESCRIPTION =
               (ADDRESS_LIST =
                 (ADDRESS = (COMMUNITY = tcpip.world)(PROTOCOL = TCP)(Host = fra2exa01 - sxdir1 - vip.europe.essilor.group)(Port = 1561))

                        */

            OracleConnection conn = new OracleConnection(connString);
            try
            {
                conn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = conn;
                cmd.CommandText = "SELECT DISTINCT JOB_NB, ROUTING_NAME, CD_PRESS, CD_MOLD, PRODUCT FROM COMBO_JOB_HEADER_TRACKING WHERE JOB_NB NOT IN(SELECT DISTINCT JOB_NB FROM COMBO_JOB_HEADER_TRACKING WHERE IS_INSERT_READY= 'T') and CD_PRESS <> 'P24' and CD_PRESS <> 'P23' order by JOB_NB desc";
                OracleDataReader reader = cmd.ExecuteReader();

                dataGridView1.Rows.Clear();
                while (reader.Read())
                {
                    dataGridView1.Rows.Add(reader[0], reader[1], reader[2], reader[3], reader[4]);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
        }

        private void Extraction_Jobs_Load(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        public class donnees2
        {
            public donnees2() { }

            public string LineIndex { get; set; }
            public string ColumnIndex { get; set; }
            public string Eye { get; set; }

            

        }
        public class donnees3
        {
            public donnees3() { }

            public string numero { get; set; }
            public float HauteurCentre { get; set; }
            public float HauteurBord { get; set; }

        }

        public class donnees4
        {
            public donnees4() { }

            public string numero { get; set; }
            public float HauteurCentre { get; set; }
            public float HauteurBord { get; set; }

        }


        private void button3_Click(object sender, EventArgs e)
        {
            

            var shots_nb = 0;
            var cd_press = "";
            var cd_mold = "";
            var routing_name = "";

            OracleConnection conn = new OracleConnection(connString); ;
            conn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            cmd.CommandText = "SELECT distinct (SHOTS_NB), CD_PRESS, CD_MOLD, ROUTING_NAME FROM COMBO_JOB_HEADER_TRACKING where JOB_NB = '" + dataGridView1.CurrentRow.Cells["JOB_NB"].Value + "' ";
            OracleDataReader reader = cmd.ExecuteReader();
            while (reader.Read() == true)
            {
                shots_nb = reader.GetInt32(0);
                cd_press = reader.GetString(1);
                cd_mold = reader.GetString(2);
                routing_name = reader.GetString(3);
            }

            conn.Close();





            object misValue = System.Reflection.Missing.Value;

            Excel._Application xlsp = new Microsoft.Office.Interop.Excel.Application();

            if (xlsp == null)
            {
                MessageBox.Show("Excel n'est pas corectement installé");
            }

            Excel.Workbook xlworkbook = xlsp.Workbooks.Add(misValue);
            //Enregistrer les données dans la feuille 1
            Excel.Worksheet xlworkSheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item(1);
            xlworkSheet.get_Range("A1", "I30").Borders.Weight = Excel.XlBorderWeight.xlThin;

            xlworkSheet.Cells[1, 1] = "Date";
            xlworkSheet.Cells[1, 2] = DateTime.Now.ToString("yyyy/MM/dd");
            xlworkSheet.Cells[1, 5] = "Semaine";
            xlworkSheet.Cells[1, 6] = CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            xlworkSheet.Cells[1, 7] = "Validation préparation";
            xlworkSheet.get_Range("A1", "I1").Font.Size = 10;
            xlworkSheet.get_Range("B1", "D1").Merge(false);
            xlworkSheet.get_Range("G1", "H1").Merge(false);

            xlworkSheet.Cells[2, 7] = "VALIDATION RECEPTION REGLEUR";
            xlworkSheet.get_Range("A2", "I2").Font.Bold = 1;
            xlworkSheet.get_Range("A2:A3", "F2:F3").Merge(false);
            xlworkSheet.get_Range("G3", "I3").Merge(false);

            xlworkSheet.get_Range("A4:A30", "I4:I30").Font.Bold = 1;
            xlworkSheet.Cells[4, 1] = "Job";
            xlworkSheet.Cells[4, 5] = cd_mold;
            xlworkSheet.Cells[4, 2] = dataGridView1.CurrentRow.Cells["JOB_NB"].Value;
            xlworkSheet.get_Range("B4", "C4").Merge(false);
            xlworkSheet.Cells[4, 4] = "Moule";
            xlworkSheet.get_Range("E4", "F4").Merge(false);
            xlworkSheet.Cells[4, 7] = "Presse";
            xlworkSheet.Cells[4, 8] = cd_press;
            xlworkSheet.get_Range("H4", "I4").Merge(false);

            xlworkSheet.get_Range("A5", "I5").Merge(false);

            xlworkSheet.Cells[6, 1] = "Produit";
            xlworkSheet.Cells[6, 2] = routing_name;
            xlworkSheet.get_Range("B6", "E6").Merge(false);
            xlworkSheet.Cells[6, 6] = "Num Shot";
            xlworkSheet.Cells[6, 7] = shots_nb;
            xlworkSheet.get_Range("G6", "I6").Merge(false);

            xlworkSheet.get_Range("A7", "I7").Merge(false);

            xlworkSheet.Cells[8, 1] = "Inserts convexes";
            xlworkSheet.get_Range("A8", "I8").Merge(false);

            xlworkSheet.get_Range("A9:A10", "I9:I10").Merge(false);

            xlworkSheet.Cells[11, 1] = "Cavités";

            xlworkSheet.get_Range("B11", "I11").NumberFormat = "@";
            xlworkSheet.Cells[11, 2] = "01";
            xlworkSheet.Cells[11, 3] = "02";
            xlworkSheet.Cells[11, 4] = "03";
            xlworkSheet.Cells[11, 5] = "04";
            xlworkSheet.Cells[11, 6] = "05";
            xlworkSheet.Cells[11, 7] = "06";
            xlworkSheet.Cells[11, 8] = "07";
            xlworkSheet.Cells[11, 9] = "08";
            xlworkSheet.Cells[12, 1] = "Base/Sphère";

            xlworkSheet.Cells[13, 1] = "Ins CC";

            xlworkSheet.Cells[14, 1] = "Epais.Bord";
            //xlworkSheet.get_Range("B14", "I14").NumberFormat.NumberDecimalDigits = "2";
            xlworkSheet.Cells[15, 1] = "Epais.Centre";
            xlworkSheet.Cells[16, 1] = "Cales CC";


            xlworkSheet.get_Range("A17", "I17").Merge(false);

            xlworkSheet.Cells[18, 1] = "CYL";

            xlworkSheet.Cells[19, 1] = "ins CX";

            xlworkSheet.Cells[20, 1] = "Epais.Ctr.CX";

            xlworkSheet.Cells[21, 1] = "Epais verre";

            xlworkSheet.Cells[22, 1] = "Cales CX";


            xlworkSheet.get_Range("A23:A24", "I23:I24").Merge(false);

            xlworkSheet.Cells[25, 1] = "Changement d'inserts";
            xlworkSheet.get_Range("A25", "I25").Merge(false);
            xlworkSheet.Cells[26, 1] = "Cavités";

            xlworkSheet.get_Range("B26", "I26").NumberFormat = "@";
            xlworkSheet.Cells[26, 2] = "01";
            xlworkSheet.Cells[26, 3] = "02";
            xlworkSheet.Cells[26, 4] = "03";

            xlworkSheet.Cells[26, 5] = "04";
            xlworkSheet.Cells[26, 6] = "05";
            xlworkSheet.Cells[26, 7] = "06";
            xlworkSheet.Cells[26, 8] = "07";
            xlworkSheet.Cells[26, 9] = "08";

            string destination = @"R:\COMMUN\ACI\Data\Jobs_Combo\Job_" + dataGridView1.CurrentRow.Cells["JOB_NB"].Value;



            var cavite = "";
            var ins_cv = "";

            OracleConnection con = new OracleConnection(connString); ;
            con.Open();
            OracleCommand cm = new OracleCommand();
            cm.Connection = con;
            cm.CommandText = "SELECT COMBO_JOB_LINES.CAVITY_JOB_NB,COMBO_JOB_LINES.LB_LOGI_SKU,COMBO_ITEMS.COLUMN_INDEX,COMBO_ITEMS.LINE_INDEX,COMBO_ITEMS.EYE,COMBO_BOM.TYPE_INS_CV,COMBO_ITEMS.PRODUCT FROM COMBO_JOB_LINES, COMBO_ITEMS, COMBO_BOM where COMBO_ITEMS.LB_LOGI_SKU = COMBO_JOB_LINES.LB_LOGI_SKU and COMBO_BOM.CD_CCE_SKU = COMBO_ITEMS.CD_CCE_SKU and JOB_NB = '" + dataGridView1.CurrentRow.Cells["JOB_NB"].Value + "' ORDER BY COMBO_JOB_LINES.CAVITY_JOB_NB ";
            OracleDataReader reade = cm.ExecuteReader();
            List<donnees2> list2 = new List<donnees2>();
            List<donnees3> list3 = new List<donnees3>();
            List<donnees4> list4 = new List<donnees4>();
            while (reade.Read())
            {
                cavite = reade.GetString(0);
                ins_cv = reade.GetString(5);

                list2.Add(new donnees2
                {
                    LineIndex = reade.GetString(3),
                    ColumnIndex = reade.GetString(2),
                    Eye = reade.GetString(4),
                });

            }

            string myConnectionString = @"user id=sa; password=K@rdexlsadm21!; data source=FRESD32615\SQLEXPRESS";
            SqlConnection myConnection = new SqlConnection(myConnectionString);


            var Nom_Produit = "";


            string strRequette1 = "SELECT [Description_produit] FROM[ACI].[dbo].[Code_produit] where Product = '" + dataGridView1.CurrentRow.Cells["PRODUCT"].Value + "' ";
            MessageBox.Show(strRequette1);
            myConnection.Open();
            SqlCommand myCommandd = new SqlCommand(strRequette1, myConnection);
            SqlDataReader mySqlDataReader = myCommandd.ExecuteReader();
            while (mySqlDataReader.Read())
            {
                Nom_Produit = mySqlDataReader.GetString(0);
            }
            myConnection.Close();


            int i = 2;

            foreach (donnees2 mesdonnees2 in list2)
            {

                //MessageBox.Show(mesdonnees2.LineIndex + " " + mesdonnees2.ColumnIndex + " " + mesdonnees2.Eye);

                if (cavite == "04")
                {
                    if (i <= 5)
                    {
                        xlworkSheet.Cells[12, i] = mesdonnees2.LineIndex + "/" + mesdonnees2.ColumnIndex + "/" + mesdonnees2.Eye;
                    }

                }
                else if (cavite == "08")
                {
                    if (i <= 9)
                    {
                        xlworkSheet.Cells[12, i] = mesdonnees2.LineIndex + "/" + mesdonnees2.ColumnIndex + "/" + mesdonnees2.Eye;
                    }

                }
                i += 1;


                string strRequete1 = "SELECT DISTINCT [Numero],[Hauteur_Centre],[Hauteur_Bord] FROM [ACI].[dbo].[Inserts] where Glass = '" + ins_cv + "' and Oeil = '" + mesdonnees2.Eye + "' and Base1 = " + mesdonnees2.LineIndex + " and Addition = " + mesdonnees2.ColumnIndex + " and CCCX = 'CC' and Produit = '" + Nom_Produit + "' ";
                MessageBox.Show(strRequete1);
                myConnection.Open();
                SqlCommand myCommand1 = new SqlCommand(strRequete1, myConnection);
                SqlDataReader mySqllDataReader = myCommand1.ExecuteReader();

                while (mySqllDataReader.Read())
                {

                    list3.Add(new donnees3
                    {
                        numero = mySqllDataReader.GetString(0),
                        HauteurCentre = mySqllDataReader.GetFloat(1),
                        HauteurBord = mySqllDataReader.GetFloat(2),
                    });
                }
                myConnection.Close();

                int j = 5;
                foreach(donnees3 mesdonnees3 in list3)
                {
                    string strRequette2 = "SELECT [MaterialName] FROM [PPG].[dbo].[Materialbase] where  [MaterialName] = '"+mesdonnees3.numero+"' ";
                    MessageBox.Show(strRequette2);
                    myConnection.Open();
                    SqlCommand myCommand2 = new SqlCommand(strRequette2, myConnection);
                    SqlDataReader mySqllDataReader2 = myCommand2.ExecuteReader();
                    if(mySqllDataReader2.Read() != true)
                    {
                        MessageBox.Show("Pas d'insert disponible dans le Kardex");

                    } else
                    {
                        MessageBox.Show("Linsert est disponible dans le Kardex");
                        //je crée une nouvelle liste des numéros disponibles dans le kardex
                       
                    }

                    myConnection.Close();
                }
                //Dans la liste que j'ai crée dans le else
                // je choisis un numéro aléatoire et le saisir dans le fichier excel avec son hauteur bord et hauteur centre 










                /*
                foreach (donnees3 mesdonnees3 in list3)
                {
                   //MessageBox.Show(mesdonnees3.numero + " " + mesdonnees3.HauteurCentre + " " + mesdonnees3.HauteurBord);

                    string strRequette2 = "SELECT [MaterialName] FROM [PPG].[dbo].[Materialbase] where  [MaterialName] = '" + mesdonnees3.numero + "'  ";
                    MessageBox.Show(strRequette2);

                    /*
                    if(strRequette2 == "")
                    {
                        //supprime deux éléments à partir de la 1ere position
                        MessageBox.Show("1 = "+list3);
                        list3.RemoveRange(1, 2);
                        MessageBox.Show("2 = " + list3);
                    }
                    

                    if (cavite == "04")
                    {
                        if (j <= 5)
                        {
                            xlworkSheet.Cells[13, j] = mesdonnees3.numero;
                            xlworkSheet.Cells[14, j] = mesdonnees3.HauteurCentre;
                            xlworkSheet.Cells[15, j] = mesdonnees3.HauteurBord;
                        }

                    }
                    else if (cavite == "08")
                    {
                        if (j <= 9)
                        {
                            xlworkSheet.Cells[13, j] = mesdonnees3.numero;
                            xlworkSheet.Cells[14, j] = mesdonnees3.HauteurCentre;
                            xlworkSheet.Cells[15, j] = mesdonnees3.HauteurBord;
                        }

                    }
                    j += 1;
                }
            */

                myConnection.Close();


            }

            conn.Close();
            xlworkbook.SaveAs(destination, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue);
            xlworkbook.Close(true, misValue, misValue);
            xlsp.Quit();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            string searchValue = textBox1.Text;

            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[0].Value.ToString().Equals(searchValue))
                    {
                        row.Cells[0].Selected = true;
                        break;
                    }

                }

            }
            catch (Exception exc)
            {
                MessageBox.Show("Ce Job n'existe pas dans la liste");
                Console.WriteLine(exc.Message);
            }
        }


    }
}