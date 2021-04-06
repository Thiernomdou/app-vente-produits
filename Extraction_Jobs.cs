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
            OracleConnection conn = new OracleConnection(connString);
            // AFFICHAGE DES JOBS
            try
            {

                conn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = conn;
                cmd.CommandText = "SELECT DISTINCT JOB_NB, ROUTING_NAME, CD_PRESS, CD_MOLD, PRODUCT FROM COMBO_JOB_HEADER_TRACKING WHERE JOB_NB NOT IN(SELECT DISTINCT JOB_NB FROM COMBO_JOB_HEADER_TRACKING WHERE IS_INSERT_READY= 'T') and JOB_NB > 10125000 order by JOB_NB desc";
                //cmd.CommandText = "SELECT DISTINCT JOB_NB, ROUTING_NAME, CD_PRESS, CD_MOLD, PRODUCT FROM COMBO_JOB_HEADER_TRACKING WHERE JOB_NB = 10126971";
                OracleDataReader reader = cmd.ExecuteReader();
                dataGridView1.Rows.Clear();
                while (reader.Read())
                {
                    dataGridView1.Rows.Add(reader[0], reader[1], reader[2], reader[3], reader[4]);
                }
            }
            catch (Exception ex)
            {
                ////MessageBox.Show(ex.Message);
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
            public double HauteurCentre { get; set; }
            public double HauteurBord { get; set; }

        }
        public class donnees4
        {
            public donnees4() { }

            public float Epaisseur { get; set; }
            public float CX { get; set; }
        }
        public class donnees5
        {
            public donnees5() { }

            public string numero { get; set; }
            public double HauteurCentre { get; set; }
            public double HauteurBord { get; set; }
        }
        public class donnees6
        {
            public donnees6() { }
            public string numero { get; set; }
            public double hauteur_centre { get; set; }
            public double Cales_cc { get; set; }

        }
        public class donnees7
        {
            public donnees7() { }

            public float HauteurCentre { get; set; }
            public float HauteurBord { get; set; }

        }
        public class donnees8
        {
            public donnees8() { }

            public string CC { get; set; }
            public double Epais { get; set; }
            public string LB { get; set; }
        }
        public class donnees9
        {
            public donnees9() { }

            public string numero { get; set; }
            public double Base1 { get; set; }
            public double base2 { get; set; }
            public double Hauteur_centre { get; set; }

        }
        public class donnees10
        {
            public donnees10() { }

            public string numero { get; set; }
            public double Base1 { get; set; }
            public double base2 { get; set; }
            public double Hauteur_centre { get; set; }

        }
        public class donnees11
        {
            public donnees11() { }

            public string numero { get; set; }
            public double Base1 { get; set; }
            public string base2 { get; set; }
            public double Hauteur_centre { get; set; }

        }
        public class donnees12
        {
            public donnees12() { }

            public string LineIndex { get; set; }
            public string ColumnIndex { get; set; }
            public string Eye { get; set; }
        }
        private DataGridViewCell GetCellWhereTextExistsInGridView(string searchText, DataGridView dataGridView, int columnIndex)
        {
            DataGridViewCell cellWhereTextIsMet = null;
            //Pour chaque ligne de la grille
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (row.Cells[columnIndex].Value != null && searchText == row.Cells[columnIndex].Value.ToString())
                {
                    //le searchText est égal au texte de cette cellule
                    cellWhereTextIsMet = row.Cells[columnIndex];
                    break;
                }
            }
            return cellWhereTextIsMet;
        }
        private void button4_Click(object sender, EventArgs e)
        {
            DataGridViewCell cell = GetCellWhereTextExistsInGridView(textBox1.Text, dataGridView1, 0);
            if (cell != null)
            {
                // La valeur existe dans la grille
                // vous pouvez faire des choses supplémentaires sur la cellule
                cell.Style = new DataGridViewCellStyle { BackColor = Color.Red };
            }
            else
            {
                // La valeur n'existe pas dans la grille
                MessageBox.Show("Cet job n'existe pas");
            }
            /*
            string searchValue = textBox1.Text;
            int rowIndex = -1;

            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[row.Index].Value.ToString().Equals(searchValue))
                    {
                        rowIndex = row.Index;
                        dataGridView1.Rows[row.Index].Selected = true;
                        break;
                    }
                }
                
            }
            catch (Exception exc)
            {
                //MessageBox.Show("Ce Job n'existe pas dans la liste");
                Console.WriteLine(exc.Message);
            }
            */
        }
        private void button5_Click_1(object sender, EventArgs e)
        {
            //variable dans l'entête
            var shots_nb = 0;
            var cd_press = "";
            var cd_mold = "";
            var routing_name = "";
            var FINI_PAS_FINI = "";
            var cavite = "";
            var ins_cv = "";
            var Diametre = 0;
            int i = 2;
            int cavité = 1;

            //connection a oracle 
            OracleConnection conn = new OracleConnection(connString); ;
            conn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            cmd.CommandText = "SELECT distinct (SHOTS_NB), CD_PRESS, CD_MOLD, ROUTING_NAME, MANAGER_CODE FROM COMBO_JOB_HEADER_TRACKING where JOB_NB = '" + dataGridView1.CurrentRow.Cells["JOB_NB"].Value + "' and CD_MOLD = '" + dataGridView1.CurrentRow.Cells["CD_MOLD"].Value + "'";
            OracleDataReader reader = cmd.ExecuteReader();
            while (reader.Read() == true)
            {
                shots_nb = reader.GetInt32(0);
                cd_press = reader.GetString(1);
                cd_mold = reader.GetString(2);
                routing_name = reader.GetString(3);
                FINI_PAS_FINI = reader.GetString(4);
            }

            conn.Close();

            object misValue = System.Reflection.Missing.Value;

            Excel._Application xlsp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook xlworkbook = xlsp.Workbooks.Add(misValue);
            //Enregistrer les données dans la feuille 1
            Excel.Worksheet xlworkSheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item(1);

            //MissingFieldException en forme du dossier excel
            if (xlsp == null)
            {
                MessageBox.Show("Excel n'est pas corectement installé");
            }
            else
            {
                xlworkSheet.get_Range("A1", "I30").Borders.Weight = Excel.XlBorderWeight.xlThin;
                xlworkSheet.get_Range("A1", "I30").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xlworkSheet.Cells[1, 1] = "Date";
                xlworkSheet.Cells[1, 2] = DateTime.Now.ToString("yyyy/MM/dd");
                xlworkSheet.Cells[1, 5] = "Semaine";
                xlworkSheet.Cells[1, 6] = CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                xlworkSheet.Cells[1, 7] = "Validation préparation";
                xlworkSheet.get_Range("A1:A3", "I1:I3").Font.Size = 10;
                xlworkSheet.get_Range("B1", "D1").Merge(false);
                xlworkSheet.get_Range("G1", "H1").Merge(false);
                xlworkSheet.get_Range("G2", "H2").Merge(false);
                xlworkSheet.get_Range("G2", "H2").Font.Size = 8;
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
                xlworkSheet.get_Range("B14", "I14").NumberFormat = "0.00";
                xlworkSheet.Cells[15, 1] = "Epais.Centre";
                xlworkSheet.get_Range("B15", "I15").NumberFormat = "0.00";
                xlworkSheet.Cells[16, 1] = "Cales CC";
                xlworkSheet.get_Range("B16", "I16").NumberFormat = "0.0";
                xlworkSheet.get_Range("A17", "I17").Merge(false);
                xlworkSheet.Cells[18, 1] = "CYL";
                xlworkSheet.get_Range("B18", "I18").NumberFormat = "0.00";
                xlworkSheet.Cells[19, 1] = "ins CX";
                xlworkSheet.Cells[20, 1] = "Epais.Ctr.CX";
                xlworkSheet.get_Range("B20", "I20").NumberFormat = "0.00";
                xlworkSheet.Cells[21, 1] = "Epais verre";
                xlworkSheet.get_Range("B21", "I21").NumberFormat = "0.00";
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
            }

            string destination = @"R:\Commun\ACI\Data\Jobs_Combo\Job_" + dataGridView1.CurrentRow.Cells["JOB_NB"].Value;

            OracleConnection con = new OracleConnection(connString); ;
            con.Open();
            OracleCommand cm = new OracleCommand();
            cm.Connection = con;
            cm.CommandText = "SELECT COMBO_JOB_LINES.CAVITY_JOB_NB, COMBO_JOB_LINES.LB_LOGI_SKU, COMBO_ITEMS.COLUMN_INDEX, COMBO_ITEMS.LINE_INDEX, COMBO_ITEMS.EYE, COMBO_BOM.TYPE_INS_CV, COMBO_ITEMS.PRODUCT, COMBO_ITEMS.DIAMETER  FROM COMBO_JOB_LINES, COMBO_ITEMS, COMBO_BOM where COMBO_ITEMS.LB_LOGI_SKU = COMBO_JOB_LINES.LB_LOGI_SKU and COMBO_BOM.CD_CCE_SKU = COMBO_ITEMS.CD_CCE_SKU and JOB_NB = '" + dataGridView1.CurrentRow.Cells["JOB_NB"].Value + "' ORDER BY COMBO_JOB_LINES.CAVITY_JOB_NB ";
            OracleDataReader reade = cm.ExecuteReader();

            List<donnees2> list2 = new List<donnees2>();
            List<donnees3> list3 = new List<donnees3>();
            List<donnees4> list4 = new List<donnees4>();
            List<donnees5> list5 = new List<donnees5>();
            List<donnees6> list6 = new List<donnees6>();
            List<donnees7> list7 = new List<donnees7>();
            List<donnees8> list8 = new List<donnees8>();
            List<donnees9> list9 = new List<donnees9>();
            List<donnees10> list10 = new List<donnees10>();
            List<donnees11> list11 = new List<donnees11>();
            List<donnees12> list12 = new List<donnees12>();

            string myConnectionString = @"user id=sa; password=K@rdexlsadm21!; data source=FRESD32615\SQLEXPRESS";
            SqlConnection myConnection = new SqlConnection(myConnectionString);

            //TROUVER LE PRODUIT

            var Nom_Produit = "";

            string strRequette1 = "SELECT [Description_produit] FROM[ACI].[dbo].[Code_produit] where Product = '" + dataGridView1.CurrentRow.Cells["PRODUCT"].Value + "' ";
            myConnection.Open();
            SqlCommand myCommandd = new SqlCommand(strRequette1, myConnection);
            SqlDataReader mySqlDataReader = myCommandd.ExecuteReader();
            while (mySqlDataReader.Read())
            {
                Nom_Produit = mySqlDataReader.GetString(0);
            }
            xlworkSheet.Cells[6, 2] = Nom_Produit;
            myConnection.Close();
            con.Close();






            con.Open();
            cm.Connection = con;
            cm.CommandText = "SELECT COMBO_JOB_LINES.CAVITY_JOB_NB, COMBO_JOB_LINES.LB_LOGI_SKU, COMBO_ITEMS.COLUMN_INDEX, COMBO_ITEMS.LINE_INDEX, COMBO_ITEMS.EYE, COMBO_BOM.TYPE_INS_CV, COMBO_ITEMS.PRODUCT, COMBO_ITEMS.DIAMETER  FROM COMBO_JOB_LINES, COMBO_ITEMS, COMBO_BOM where COMBO_ITEMS.LB_LOGI_SKU = COMBO_JOB_LINES.LB_LOGI_SKU and COMBO_BOM.CD_CCE_SKU = COMBO_ITEMS.CD_CCE_SKU and JOB_NB = '" + dataGridView1.CurrentRow.Cells["JOB_NB"].Value + "' ORDER BY COMBO_JOB_LINES.CAVITY_JOB_NB ";
            OracleDataReader reade1 = cm.ExecuteReader();


            //EN FONCTION DU PRODUIT

            if (FINI_PAS_FINI == "FINIS")
            {
                //recuperer cavité + ins_cv + Diamètre
                while (reade1.Read())
                {
                    cavite = reade1.GetString(0);
                    ins_cv = reade1.GetString(5);
                    Diametre = reade1.GetInt32(7);

                    // ajoute des cavilté dans la liste 2
                    list2.Add(new donnees2
                    {
                        LineIndex = reade1.GetString(3),
                        ColumnIndex = reade1.GetString(2),
                    });
                }
                myConnection.Close();

                //TROUVER LE MOULE
                if (Diametre == 55)
                {
                    cd_mold = "DF57";
                }
                else if (Diametre == 60)
                {
                    cd_mold = "DF62";
                }
                else if (Diametre == 65)
                {
                    cd_mold = "DF67";
                }
                else if (Diametre == 70)
                {
                    cd_mold = "DF72";
                }

                /*var result = MessageBox.Show("coucou thierno", "Attention",    MessageBoxButtons.YesNo,    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    MessageBox.Show("coucou Anthony");
                }
                if (result == DialogResult.No)
                {
                    MessageBox.Show("coucou Carine");
                }*/

                // pour toutes les cavités 
                foreach (donnees2 mesdonnees2 in list2)
                {
                    var ins = "";
                    var Cales_CC = 0.0;
                    var hauteur_centreCC = 0.0;

                    xlworkSheet.Cells[12, i] = mesdonnees2.LineIndex;
                    xlworkSheet.Cells[18, i] = mesdonnees2.ColumnIndex;

                    //TROUVER L'INS CC
                    if (Diametre == 55)
                    {
                        myConnection.Close();
                        string strRequete1 = "SELECT [INS] FROM[ACI].[dbo].[Table_PCF_Kids_ins_num_concave_55] where Cylindre ='" + mesdonnees2.ColumnIndex + "' and Sphere ='" + mesdonnees2.LineIndex + "' ";
                        myConnection.Open();
                        SqlCommand myCommand1 = new SqlCommand(strRequete1, myConnection);
                        SqlDataReader mySqllDataReader = myCommand1.ExecuteReader();

                        while (mySqllDataReader.Read())
                        {
                            ins = mySqllDataReader.GetString(0);
                        }
                        myConnection.Close();
                    }
                    else if (Diametre == 60)
                    {
                        myConnection.Close();
                        string strRequete1 = "SELECT [INS] FROM[ACI].[dbo].[Table_PCF_Kids_ins_num_concave_60] where Cylindre ='" + mesdonnees2.ColumnIndex + "' and Sphere ='" + mesdonnees2.LineIndex + "' ";
                        myConnection.Open();
                        SqlCommand myCommand1 = new SqlCommand(strRequete1, myConnection);
                        SqlDataReader mySqllDataReader = myCommand1.ExecuteReader();

                        while (mySqllDataReader.Read())
                        {
                            ins = mySqllDataReader.GetString(0);
                        }
                        myConnection.Close();
                    }
                    else
                    {
                        myConnection.Close();
                        string strRequete1 = "SELECT [INS] FROM[ACI].[dbo].[Table_PCF_ins_num_concave] where Cylindre ='" + mesdonnees2.ColumnIndex + "' and Sphere ='" + mesdonnees2.LineIndex + "' ";
                        myConnection.Open();
                        SqlCommand myCommand1 = new SqlCommand(strRequete1, myConnection);
                        SqlDataReader mySqllDataReader = myCommand1.ExecuteReader();

                        while (mySqllDataReader.Read())
                        {
                            ins = mySqllDataReader.GetString(0);
                        }
                        myConnection.Close();
                    }

                    list5.Clear();

                    // CHERCHER L'INS CC DANS DANS LE KARDEX
                    string strRequete2 = "  SELECT distinct C.MaterialName ,D.Hauteur_Bord, D.Hauteur_Centre FROM[PPG].[dbo].[LocContent]A, [PPG].[dbo].[LocContentbreakdown]B,[PPG].[dbo].[Materialbase] C, [ACI].[dbo].[Inserts] D where A.LocContentId = B.LocContentId and C.MaterialId = A.MaterialId and D.Numero = C.MaterialName COLLATE Latin1_General_CI_AS and C.Info1 = '" + ins + "'";
                    myConnection.Open();
                    SqlCommand myCommand2 = new SqlCommand(strRequete2, myConnection);
                    SqlDataReader mySqllDataReader2 = myCommand2.ExecuteReader();

                    while (mySqllDataReader2.Read())
                    {
                        list5.Add(new donnees5
                        {
                            numero = mySqllDataReader2.GetString(0),
                            HauteurBord = mySqllDataReader2.GetDouble(1),
                            HauteurCentre = mySqllDataReader2.GetDouble(2),
                        });
                    }
                    myConnection.Close();

                    Boolean trouve = true;

                    // DOUBLONS CC + CC_CALES
                    while (trouve == true)
                    {
                        trouve = false;
                        var j = 0;

                        if (list5.Count > 0)
                        {
                            Random rnd = new Random();

                            int nbr = rnd.Next(0, list5.Count);

                            foreach (donnees5 mesdonnees5 in list5)
                            {
                                if (j == nbr)
                                {
                                    foreach (donnees6 mesdonnees6 in list6)
                                    {
                                        if (mesdonnees5.numero == mesdonnees6.numero)
                                        {
                                            trouve = true;
                                        }
                                    }
                                    if (trouve == false)
                                    {
                                        xlworkSheet.Cells[13, i] = mesdonnees5.numero;
                                        xlworkSheet.Cells[14, i] = mesdonnees5.HauteurBord;
                                        xlworkSheet.Cells[15, i] = mesdonnees5.HauteurCentre;

                                        myConnection.Close();
                                        string strRequete4 = "SELECT distinct [Profondeur_Moule].[CC]-[Inserts].Hauteur_Bord FROM [ACI].[dbo].[Profondeur_Moule],[ACI].[dbo].[Inserts] WHERE [Profondeur_Moule].Moule = '" + cd_mold + "' AND [Inserts].Numero = '" + mesdonnees5.numero + "'";
                                        ////MessageBox.Show(strRequete1);
                                        myConnection.Open();
                                        SqlCommand myCommand4 = new SqlCommand(strRequete4, myConnection);
                                        SqlDataReader mySqllDataReader4 = myCommand4.ExecuteReader();
                                        while (mySqllDataReader4.Read())
                                        {
                                            xlworkSheet.Cells[16, i] = mySqllDataReader4.GetDouble(0);
                                            Cales_CC = mySqllDataReader4.GetDouble(0);
                                            hauteur_centreCC = mesdonnees5.HauteurCentre;
                                            list6.Add(new donnees6
                                            {
                                                numero = mesdonnees5.numero,
                                                hauteur_centre = mesdonnees5.HauteurCentre,
                                                Cales_cc = mySqllDataReader4.GetDouble(0),
                                            });
                                        }
                                        myConnection.Close();
                                    }
                                }
                                j += 1;
                            }
                            list5.RemoveAt(nbr);
                        }
                        else
                        {
                            //MessageBox.Show("pas d'insert disponible pour la cavité " + i);
                            trouve = false;
                        }
                    }

                    //TROUVER L'INS CX
                    if (Diametre == 55)
                    {
                        myConnection.Close();
                        string strRequete1 = "SELECT [INS] FROM[ACI].[dbo].[Table_PCF_Kids_ins_num_convex_55] where Cylindre ='" + mesdonnees2.ColumnIndex + "' and Sphere ='" + mesdonnees2.LineIndex + "' ";
                        myConnection.Open();
                        SqlCommand myCommand1 = new SqlCommand(strRequete1, myConnection);
                        SqlDataReader mySqllDataReader = myCommand1.ExecuteReader();

                        while (mySqllDataReader.Read())
                        {
                            ins = mySqllDataReader.GetString(0);
                        }
                        myConnection.Close();
                    }
                    else if (Diametre == 60)
                    {
                        myConnection.Close();
                        string strRequete1 = "SELECT [INS] FROM[ACI].[dbo].[Table_PCF_Kids_ins_num_convex_60] where Cylindre ='" + mesdonnees2.ColumnIndex + "' and Sphere ='" + mesdonnees2.LineIndex + "' ";
                        myConnection.Open();
                        SqlCommand myCommand1 = new SqlCommand(strRequete1, myConnection);
                        SqlDataReader mySqllDataReader = myCommand1.ExecuteReader();

                        while (mySqllDataReader.Read())
                        {
                            ins = mySqllDataReader.GetString(0);
                        }
                        myConnection.Close();
                    }
                    else
                    {
                        myConnection.Close();
                        string strRequete1 = "SELECT [INS] FROM[ACI].[dbo].[Table_PCF_ins_num_convex] where Cylindre ='" + mesdonnees2.ColumnIndex + "' and Sphere ='" + mesdonnees2.LineIndex + "' ";
                        myConnection.Open();
                        SqlCommand myCommand1 = new SqlCommand(strRequete1, myConnection);
                        SqlDataReader mySqllDataReader = myCommand1.ExecuteReader();

                        while (mySqllDataReader.Read())
                        {
                            ins = mySqllDataReader.GetString(0);
                        }
                        myConnection.Close();
                    }

                    //CHERCHER L'INS CX DANS DANS LE KARDEX
                    list10.Clear();
                    string strRequete5 = "  SELECT distinct C.MaterialName , D.Hauteur_Centre FROM[PPG].[dbo].[LocContent]A, [PPG].[dbo].[LocContentbreakdown]B,[PPG].[dbo].[Materialbase] C, [ACI].[dbo].[Inserts] D where A.LocContentId = B.LocContentId and C.MaterialId = A.MaterialId and D.Numero = C.MaterialName COLLATE Latin1_General_CI_AS and C.Info1 = '" + ins + "' ";
                    myConnection.Open();
                    SqlCommand myCommand5 = new SqlCommand(strRequete5, myConnection);
                    SqlDataReader mySqllDataReader5 = myCommand5.ExecuteReader();

                    while (mySqllDataReader5.Read())
                    {
                        list10.Add(new donnees10
                        {
                            numero = mySqllDataReader5.GetString(0),
                            Hauteur_centre = mySqllDataReader5.GetDouble(1),
                        });
                    }
                    myConnection.Close();

                    // DOUBLONS CX POUR LE FINI
                    trouve = true;
                    var hauteur_centreCX = 0.0;
                    while (trouve == true)
                    {
                        trouve = false;

                        var j = 0;

                        if (list10.Count > 0)
                        {
                            Random rnd = new Random();

                            int nbr = rnd.Next(0, list10.Count);
                            foreach (donnees10 mesdonnees10 in list10)
                            {
                                if (j == nbr)
                                {
                                    foreach (donnees11 mesdonnees11 in list11)
                                    {
                                        if (mesdonnees10.numero == mesdonnees11.numero)
                                        {
                                            trouve = true;
                                        }
                                    }
                                    if (trouve == false)
                                    {
                                        xlworkSheet.Cells[19, i] = mesdonnees10.numero;
                                        xlworkSheet.Cells[20, i] = mesdonnees10.Hauteur_centre;
                                        hauteur_centreCX = mesdonnees10.Hauteur_centre;
                                        list11.Add(new donnees11
                                        {
                                            numero = mesdonnees10.numero,
                                            Hauteur_centre = mesdonnees10.Hauteur_centre,
                                        });
                                    }
                                }
                                j += 1;
                            }
                        }
                        else
                        {
                            //MessageBox.Show("pas d'insert disponible pour la cavité " + i);
                            trouve = false;
                        }
                    }

                    // TROUVER EPAISSEUR VERRE
                    var epaisseur = 0.0;
                    if (Diametre == 55 || Diametre == 60)
                    {
                        string strRequete7 = "SELECT [Epaisseur] FROM [ACI].[dbo].[Table_PCF_Kids_epaisseur] where Cylindre ='" + mesdonnees2.ColumnIndex + "' and Sphere ='" + mesdonnees2.LineIndex + "' ";
                        myConnection.Open();
                        SqlCommand myCommand7 = new SqlCommand(strRequete7, myConnection);
                        SqlDataReader mySqllDataReader7 = myCommand7.ExecuteReader();

                        while (mySqllDataReader7.Read())
                        {
                            epaisseur = mySqllDataReader7.GetDouble(0);
                        }
                        myConnection.Close();
                    }
                    else
                    {
                        string strRequete7 = "SELECT [Epaisseur] FROM [ACI].[dbo].[Table_PCF_Epaisseur] where Cylindre ='" + mesdonnees2.ColumnIndex + "' and Sphere ='" + mesdonnees2.LineIndex + "' ";
                        //MessageBox.Show(strRequete7);
                        myConnection.Open();
                        SqlCommand myCommand7 = new SqlCommand(strRequete7, myConnection);
                        SqlDataReader mySqllDataReader7 = myCommand7.ExecuteReader();

                        while (mySqllDataReader7.Read())
                        {
                            epaisseur = mySqllDataReader7.GetDouble(0);
                        }
                        myConnection.Close();
                    }
                    xlworkSheet.Cells[21, i] = epaisseur;

                    // TROUVER PROFONDEUR MOULES 
                    var Profondeur_mold_CX = 0.0;
                    string strRequete9 = "SELECT distinct [Profondeur_Moule].[CX] FROM [ACI].[dbo].[Profondeur_Moule] WHERE [Profondeur_Moule].Moule = '" + cd_mold + "'";
                    ////MessageBox.Show(strRequete1);
                    myConnection.Open();
                    SqlCommand myCommand9 = new SqlCommand(strRequete9, myConnection);
                    SqlDataReader mySqllDataReader9 = myCommand9.ExecuteReader();

                    while (mySqllDataReader9.Read())
                    {
                        Profondeur_mold_CX = mySqllDataReader9.GetDouble(0);
                    }
                    myConnection.Close();

                    // CALCULES CALES CX
                    var somme = 0.0;

                    somme = epaisseur + hauteur_centreCX + hauteur_centreCC + Cales_CC;

                    xlworkSheet.Cells[22, i] = Profondeur_mold_CX - somme;

                    i += 1;
                    cavité += 1;
                }
                //MessageBox.Show("" + Diametre);

            }
            else // POUR LE SEMI FINI
            {
                if (Nom_Produit == "SPHERIQUE" || Nom_Produit == "ASPHERIQUE") //SI PROGRESSIF
                {
                    // RECUPE INSCV SPHERE CYLINDRE CAVITE
                    while (reade1.Read())
                    {
                        cavite = reade1.GetString(0);
                        ins_cv = reade1.GetString(5);

                        // ajoute des cacilté dans la liste 2
                        list2.Add(new donnees2
                        {
                            LineIndex = reade1.GetString(3),
                            ColumnIndex = reade1.GetString(2),
                        });
                    }

                    // POUR TOUTES LES CAVITES
                    foreach (donnees2 mesdonnees2 in list2)
                    {
                        xlworkSheet.Cells[12, i] = mesdonnees2.LineIndex + "/" + mesdonnees2.ColumnIndex;


                        myConnection.Close();
                        string strRequete1 = "SELECT D.Numero, D.Hauteur_Centre,D.Hauteur_Bord FROM[PPG].[dbo].[LocContent]A, [PPG].[dbo].[LocContentbreakdown]B,[PPG].[dbo].[Materialbase] C, [ACI].[dbo].[Inserts] D where A.LocContentId = B.LocContentId and C.MaterialId = A.MaterialId and D.Numero = C.MaterialName COLLATE Latin1_General_CI_AS and D.Base1 ='" + mesdonnees2.LineIndex + "' and D.Produit = '" + Nom_Produit + "' and D.Glass = '" + ins_cv + "' and D.CCCX = 'CC' group by D.Numero,D.Hauteur_Centre,D.Hauteur_Bord";
                        myConnection.Open();
                        SqlCommand myCommand1 = new SqlCommand(strRequete1, myConnection);
                        SqlDataReader mySqllDataReader = myCommand1.ExecuteReader();

                        list3.Clear();

                        while (mySqllDataReader.Read())
                        {
                            list3.Add(new donnees3
                            {
                                numero = mySqllDataReader.GetString(0),
                                HauteurCentre = mySqllDataReader.GetDouble(1),
                                HauteurBord = mySqllDataReader.GetDouble(2),
                            });
                        }

                        myConnection.Close();
                        Boolean trouve = true;
                        trouve = true;

                        //VERIFICATION DOUBLON CC SPHERIQUE 
                        while (trouve == true)
                        {
                            trouve = false;

                            var j = 0;
                            if (list3.Count > 0)
                            {
                                Random rnd = new Random();

                                int nbr = rnd.Next(0, list3.Count);

                                foreach (donnees3 mesdonnees3 in list3)
                                {
                                    if (j == nbr)
                                    {
                                        foreach (donnees6 mesdonnees6 in list6)
                                        {
                                            if (mesdonnees3.numero == mesdonnees6.numero)
                                            {
                                                trouve = true;
                                            }
                                        }

                                        if (trouve == false)
                                        {
                                            xlworkSheet.Cells[13, i] = mesdonnees3.numero;
                                            xlworkSheet.Cells[14, i] = mesdonnees3.HauteurBord;
                                            xlworkSheet.Cells[15, i] = mesdonnees3.HauteurCentre;

                                            string strRequete4 = "SELECT distinct [Profondeur_Moule].[CC]-[Inserts].Hauteur_Bord FROM [ACI].[dbo].[Profondeur_Moule],[ACI].[dbo].[Inserts] WHERE [Profondeur_Moule].Moule = '" + cd_mold + "' AND [Inserts].Numero = '" + mesdonnees3.numero + "'";
                                            myConnection.Open();
                                            SqlCommand myCommand4 = new SqlCommand(strRequete4, myConnection);
                                            SqlDataReader mySqllDataReader4 = myCommand4.ExecuteReader();
                                            while (mySqllDataReader4.Read())
                                            {
                                                xlworkSheet.Cells[16, i] = mySqllDataReader4.GetDouble(0);
                                                list6.Add(new donnees6
                                                {
                                                    numero = mesdonnees3.numero,
                                                    hauteur_centre = mesdonnees3.HauteurCentre,
                                                    Cales_cc = mySqllDataReader4.GetDouble(0),
                                                });
                                            }
                                            myConnection.Close();
                                        }
                                    }
                                    j += 1;
                                }
                                list3.RemoveAt(nbr);
                            }
                            else
                            {
                                MessageBox.Show("il n'y a pas assez d'insert CC pour la cavité " + cavité + "");
                                trouve = false;
                            }
                        }


                        string strRequete2 = "SELECT [CX],[Epais],[LB] FROM[ACI].[dbo].[Contrainte_all] WHERE Produit = '" + Nom_Produit + "' AND Base = '" + mesdonnees2.LineIndex + "'";
                        myConnection.Open();
                        SqlCommand myCommand2 = new SqlCommand(strRequete2, myConnection);
                        SqlDataReader mySqllDataReader2 = myCommand2.ExecuteReader();
                        var CC = "";

                        while (mySqllDataReader2.Read())
                        {
                            xlworkSheet.Cells[18, i] = mySqllDataReader2.GetString(0);
                            xlworkSheet.Cells[21, i] = mySqllDataReader2.GetDouble(1);

                            list8.Add(new donnees8
                            {
                                CC = mySqllDataReader2.GetString(0),
                                Epais = mySqllDataReader2.GetDouble(1),
                                LB = mySqllDataReader2.GetString(2),
                            });

                            CC = mySqllDataReader2.GetString(0);
                        }

                        myConnection.Close();
                        list10.Clear();

                        string strRequete3 = "SELECT D.Numero, D.Hauteur_Centre,D.Hauteur_Bord FROM[PPG].[dbo].[LocContent]A, [PPG].[dbo].[LocContentbreakdown]B,[PPG].[dbo].[Materialbase] C, [ACI].[dbo].[Inserts] D where A.LocContentId = B.LocContentId and C.MaterialId = A.MaterialId and D.Numero = C.MaterialName COLLATE Latin1_General_CI_AS and D.Base1 ='" + CC + "' and D.Produit = '" + Nom_Produit + "' and D.Glass = '" + ins_cv + "' and D.CCCX = 'CX'";
                        myConnection.Open();
                        SqlCommand myCommand3 = new SqlCommand(strRequete3, myConnection);
                        SqlDataReader mySqllDataReader3 = myCommand3.ExecuteReader();

                        while (mySqllDataReader3.Read())
                        {
                            list10.Add(new donnees10
                            {
                                numero = mySqllDataReader3.GetString(0),
                                Hauteur_centre = mySqllDataReader3.GetDouble(1),
                            });
                        }

                        myConnection.Close();
                        trouve = true;

                        //VERIFICATION DOUBLON CX SPHERIQUE 
                        while (trouve == true)
                        {
                            trouve = false;

                            var j = 0;

                            if (list10.Count > 0)
                            {
                                Random rnd = new Random();

                                int nbr = rnd.Next(0, list10.Count);

                                foreach (donnees10 mesdonnees10 in list10)
                                {
                                    if (j == nbr)
                                    {
                                        foreach (donnees11 mesdonnees11 in list11)
                                        {
                                            if (mesdonnees10.numero == mesdonnees11.numero)
                                            {
                                                trouve = true;

                                            }
                                        }

                                        if (trouve == false)
                                        {
                                            xlworkSheet.Cells[19, i] = mesdonnees10.numero;
                                            xlworkSheet.Cells[20, i] = mesdonnees10.Hauteur_centre;


                                            list11.Add(new donnees11
                                            {
                                                numero = mesdonnees10.numero,
                                                Hauteur_centre = mesdonnees10.Hauteur_centre,
                                            });

                                            myConnection.Close();
                                        }
                                    }
                                    j += 1;
                                }
                                list10.RemoveAt(nbr);
                            }
                            else
                            {
                                MessageBox.Show("il n'y a pas assez d'insert CX pour la cavité " + cavité + "");
                                trouve = false;
                            }
                        }

                        i += 1;
                        cavité += 1;
                    }

                }
                else
                {
                    ///////////////////////////////////////////
                    //SI NON PROGRESSIF
                    /////////////////////////////////////////
                    while (reade1.Read())
                    {
                        cavite = reade1.GetString(0);
                        ins_cv = reade1.GetString(5);

                        // ajoute des cacilté dans la liste 2
                        list2.Add(new donnees2
                        {
                            LineIndex = reade1.GetString(3),
                            ColumnIndex = reade1.GetString(2),
                            Eye = reade1.GetString(4),
                        });
                    }
                    ///////////////////////////////////////////
                    //POUR TOUTES LES CAV
                    /////////////////////////////////////////
                    foreach (donnees2 mesdonnees2 in list2)
                    {

                        xlworkSheet.Cells[12, i] = mesdonnees2.LineIndex + "/" + mesdonnees2.ColumnIndex + "/" + mesdonnees2.Eye;

                        ///////////////////////////////////////////////////////
                        //TOUT LES INSERTS POUR LE NON PROGESSIF DISPONIBLE
                        ///////////////////////////////////////////////////////
                        myConnection.Close();
                        string strRequete1 = "SELECT distinct C.MaterialName ,D.Hauteur_Bord, D.Hauteur_Centre FROM[PPG].[dbo].[LocContent]A, [PPG].[dbo].[LocContentbreakdown]B,[PPG].[dbo].[Materialbase] C, [ACI].[dbo].[Inserts] D where A.LocContentId = B.LocContentId and C.MaterialId = A.MaterialId and D.Numero = C.MaterialName COLLATE Latin1_General_CI_AS and Glass = '" + ins_cv + "' and Oeil = '" + mesdonnees2.Eye + "' and Base1 = " + mesdonnees2.LineIndex + " and Addition = " + mesdonnees2.ColumnIndex + " and CCCX = 'CC' and Produit = '" + Nom_Produit + "' ";
                        myConnection.Open();
                        SqlCommand myCommand1 = new SqlCommand(strRequete1, myConnection);
                        SqlDataReader mySqllDataReader = myCommand1.ExecuteReader();

                        list5.Clear();

                        while (mySqllDataReader.Read())
                        {
                            list5.Add(new donnees5
                            {
                                numero = mySqllDataReader.GetString(0),
                                HauteurCentre = mySqllDataReader.GetDouble(2),
                                HauteurBord = mySqllDataReader.GetDouble(1),
                            });
                        }
                        myConnection.Close();


                        ///////////////////////////////////////////////////////
                        //VERIFICATION DES DOUBLONS CC ROCKY
                        ///////////////////////////////////////////////////////
                        var Cales_CC = 0.0;
                        var hauteur_centreCC = 0.0;
                        Boolean trouve = true;
                        trouve = true;

                        while (trouve == true)
                        {
                            trouve = false;

                            var j = 0;

                            if (list5.Count > 0)
                            {
                                Random rnd = new Random();

                                int nbr = rnd.Next(0, list5.Count);

                                foreach (donnees5 mesdonnees5 in list5)
                                {
                                    if (j == nbr)
                                    {
                                        foreach (donnees6 mesdonnees6 in list6)
                                        {
                                            if (mesdonnees5.numero == mesdonnees6.numero)
                                            {
                                                trouve = true;
                                            }
                                        }

                                        if (trouve == false)
                                        {
                                            xlworkSheet.Cells[13, i] = mesdonnees5.numero;
                                            xlworkSheet.Cells[14, i] = mesdonnees5.HauteurBord;
                                            xlworkSheet.Cells[15, i] = mesdonnees5.HauteurCentre;

                                            string strRequete4 = "SELECT distinct [Profondeur_Moule].[CC]-[Inserts].Hauteur_Bord FROM [ACI].[dbo].[Profondeur_Moule],[ACI].[dbo].[Inserts] WHERE [Profondeur_Moule].Moule = '" + cd_mold + "' AND [Inserts].Numero = '" + mesdonnees5.numero + "'";
                                            ////MessageBox.Show(strRequete1);
                                            myConnection.Open();
                                            SqlCommand myCommand4 = new SqlCommand(strRequete4, myConnection);
                                            SqlDataReader mySqllDataReader4 = myCommand4.ExecuteReader();
                                            while (mySqllDataReader4.Read())
                                            {
                                                xlworkSheet.Cells[16, i] = mySqllDataReader4.GetDouble(0);
                                                Cales_CC = mySqllDataReader4.GetDouble(0);
                                                hauteur_centreCC = mesdonnees5.HauteurCentre;
                                                list6.Add(new donnees6
                                                {
                                                    numero = mesdonnees5.numero,
                                                    hauteur_centre = mesdonnees5.HauteurCentre,
                                                    Cales_cc = mySqllDataReader4.GetDouble(0),
                                                });
                                            }
                                            myConnection.Close();
                                        }
                                    }
                                    j += 1;
                                }
                                list5.RemoveAt(nbr);
                            }
                            else
                            {
                                MessageBox.Show("il n'y a pas assez d'insert CC pour la cavité " + cavité + "");
                                trouve = false;
                            }
                        }

                        if (Nom_Produit == "SNL" || Nom_Produit == "CONDE")
                        {
                            var CC = "";
                            var epais = 0.0;
                            string strRequete5 = "SELECT [CX],[Epais] FROM[ACI].[dbo].[Contrainte_all] WHERE Produit = '" + Nom_Produit + "' AND Base = '" + mesdonnees2.LineIndex + "' AND Addition ='" + mesdonnees2.ColumnIndex + "'";
                            myConnection.Open();
                            SqlCommand myCommand5 = new SqlCommand(strRequete5, myConnection);
                            SqlDataReader mySqllDataReader5 = myCommand5.ExecuteReader();
                            while (mySqllDataReader5.Read())
                            {
                                xlworkSheet.Cells[18, i] = mySqllDataReader5.GetString(0);
                                xlworkSheet.Cells[21, i] = mySqllDataReader5.GetDouble(1);
                                CC = mySqllDataReader5.GetString(0);
                                epais = mySqllDataReader5.GetDouble(1);

                                list8.Add(new donnees8
                                {
                                    CC = mySqllDataReader5.GetString(0),
                                    Epais = mySqllDataReader5.GetDouble(1),
                                });
                            }
                            myConnection.Close();
                            list10.Clear();

                            string strRequete6 = "SELECT distinct C.MaterialName ,D.Base1, D.Base2 ,D.Hauteur_Centre FROM[PPG].[dbo].[LocContent]A, [PPG].[dbo].[LocContentbreakdown]B,[PPG].[dbo].[Materialbase] C, [ACI].[dbo].[Inserts] D where A.LocContentId = B.LocContentId and C.MaterialId = A.MaterialId and D.Numero = C.MaterialName COLLATE Latin1_General_CI_AS and D.Base1 = '" + CC + "' and(D.[Numero] like 'C%' or D.[Numero] like 'P%')";
                            myConnection.Open();
                            SqlCommand myCommand6 = new SqlCommand(strRequete6, myConnection);
                            SqlDataReader mySqllDataReader6 = myCommand6.ExecuteReader();
                            while (mySqllDataReader6.Read())
                            {
                                list10.Add(new donnees10
                                {
                                    numero = mySqllDataReader6.GetString(0),
                                    Base1 = mySqllDataReader6.GetDouble(1),
                                    base2 = mySqllDataReader6.GetDouble(2),
                                    Hauteur_centre = mySqllDataReader6.GetDouble(3),
                                });
                            }
                            myConnection.Close();

                            //DOUBLONS CX SNL CONDE
                            trouve = true;
                            var hauteur_centreCX = 0.0;
                            while (trouve == true)
                            {
                                trouve = false;

                                Random rnd = new Random();

                                int nbr = rnd.Next(0, list10.Count);


                                var j = 0;
                                if (list10.Count > 0)
                                {
                                    foreach (donnees10 mesdonnees10 in list10)
                                    {
                                        if (j == nbr)
                                        {
                                            foreach (donnees11 mesdonnees11 in list11)
                                            {
                                                if (mesdonnees10.numero == mesdonnees11.numero)
                                                {
                                                    trouve = true;
                                                }
                                            }
                                            if (trouve == false)
                                            {
                                                xlworkSheet.Cells[19, i] = mesdonnees10.numero;
                                                xlworkSheet.Cells[20, i] = mesdonnees10.Hauteur_centre;
                                                hauteur_centreCX = mesdonnees10.Hauteur_centre;

                                                string strRequete4 = "SELECT distinct [Profondeur_Moule].[CC]-[Inserts].Hauteur_Bord FROM [ACI].[dbo].[Profondeur_Moule],[ACI].[dbo].[Inserts] WHERE [Profondeur_Moule].Moule = '" + cd_mold + "' AND [Inserts].Numero = '" + mesdonnees10.numero + "'";
                                                ////MessageBox.Show(strRequete1);
                                                myConnection.Open();
                                                SqlCommand myCommand4 = new SqlCommand(strRequete4, myConnection);
                                                SqlDataReader mySqllDataReader4 = myCommand4.ExecuteReader();
                                                while (mySqllDataReader4.Read())
                                                {
                                                    list11.Add(new donnees11
                                                    {
                                                        numero = mesdonnees10.numero,
                                                        Hauteur_centre = mesdonnees10.Hauteur_centre,
                                                    });
                                                }
                                                myConnection.Close();
                                            }
                                        }
                                        j += 1;
                                    }
                                    list10.RemoveAt(nbr);
                                }
                                else
                                {
                                    //MessageBox.Show("pas d'insert disponible pour la cavité " + i);
                                    trouve = false;
                                }
                            }

                            var somme = 0.0;
                            var moule_CX = 0.0;

                            string strRequete8 = "SELECT [CC],[CX] FROM[ACI].[dbo].[Profondeur_Moule] where Moule =  '" + cd_mold + "'";
                            myConnection.Open();
                            SqlCommand myCommand8 = new SqlCommand(strRequete8, myConnection);
                            SqlDataReader mySqllDataReader8 = myCommand8.ExecuteReader();
                            while (mySqllDataReader8.Read())
                            {
                                moule_CX = mySqllDataReader8.GetDouble(1);
                            }

                            somme = epais + hauteur_centreCC + hauteur_centreCX + Cales_CC;
                            xlworkSheet.Cells[22, i] = moule_CX - somme;
                        }
                        else
                        {
                            var CC = "";
                            var epais = 0.0;
                            string strRequete5 = "SELECT [CX],[Epais] FROM[ACI].[dbo].[Contrainte_all] WHERE Produit = '" + Nom_Produit + "' AND Base = '" + mesdonnees2.LineIndex + "'";
                            myConnection.Open();
                            SqlCommand myCommand5 = new SqlCommand(strRequete5, myConnection);
                            SqlDataReader mySqllDataReader5 = myCommand5.ExecuteReader();
                            while (mySqllDataReader5.Read())
                            {
                                xlworkSheet.Cells[18, i] = mySqllDataReader5.GetString(0);
                                xlworkSheet.Cells[21, i] = mySqllDataReader5.GetDouble(1);
                                CC = mySqllDataReader5.GetString(0);
                                epais = mySqllDataReader5.GetDouble(1);

                                list8.Add(new donnees8
                                {
                                    CC = mySqllDataReader5.GetString(0),
                                    Epais = mySqllDataReader5.GetDouble(1),
                                });
                            }
                            myConnection.Close();

                            list10.Clear();
                            string strRequete6 = "SELECT distinct C.MaterialName ,D.Base1, D.Base2 ,D.Hauteur_Centre FROM[PPG].[dbo].[LocContent]A, [PPG].[dbo].[LocContentbreakdown]B,[PPG].[dbo].[Materialbase] C, [ACI].[dbo].[Inserts] D where A.LocContentId = B.LocContentId and C.MaterialId = A.MaterialId and D.Numero = C.MaterialName COLLATE Latin1_General_CI_AS and D.Base1 = '" + CC + "' and(D.[Numero] like 'C%' or D.[Numero] like 'P%')";
                            myConnection.Open();
                            SqlCommand myCommand6 = new SqlCommand(strRequete6, myConnection);
                            SqlDataReader mySqllDataReader6 = myCommand6.ExecuteReader();
                            while (mySqllDataReader6.Read())
                            {
                                list10.Add(new donnees10
                                {
                                    numero = mySqllDataReader6.GetString(0),
                                    Base1 = mySqllDataReader6.GetDouble(1),
                                    base2 = mySqllDataReader6.GetDouble(2),
                                    Hauteur_centre = mySqllDataReader6.GetDouble(3),
                                });
                            }
                            myConnection.Close();

                            //DOUBLONS CX ROCKY
                            trouve = true;
                            var hauteur_centreCX = 0.0;
                            while (trouve == true)
                            {
                                trouve = false;
                                var j = 0;

                                if (list10.Count > 0)
                                {
                                    Random rnd = new Random();

                                    int nbr = rnd.Next(0, list10.Count);
                                    foreach (donnees10 mesdonnees10 in list10)
                                    {
                                        if (j == nbr)
                                        {
                                            foreach (donnees11 mesdonnees11 in list11)
                                            {
                                                if (mesdonnees10.numero == mesdonnees11.numero)
                                                {
                                                    trouve = true;
                                                }
                                            }
                                            if (trouve == false)
                                            {
                                                xlworkSheet.Cells[19, i] = mesdonnees10.numero;
                                                xlworkSheet.Cells[20, i] = mesdonnees10.Hauteur_centre;
                                                hauteur_centreCX = mesdonnees10.Hauteur_centre;
                                                list11.Add(new donnees11
                                                {
                                                    numero = mesdonnees10.numero,
                                                    Hauteur_centre = mesdonnees10.Hauteur_centre,
                                                });
                                            }
                                        }
                                        j += 1;
                                    }
                                }
                                else
                                {
                                    //MessageBox.Show("pas d'insert disponible pour la cavité " + i);
                                    trouve = false;
                                }
                            }

                            var somme = 0.0;
                            var moule_CX = 0.0;

                            string strRequete8 = "SELECT [CC],[CX] FROM[ACI].[dbo].[Profondeur_Moule] where Moule =  '" + cd_mold + "'";
                            myConnection.Open();
                            SqlCommand myCommand8 = new SqlCommand(strRequete8, myConnection);
                            SqlDataReader mySqllDataReader8 = myCommand8.ExecuteReader();
                            while (mySqllDataReader8.Read())
                            {
                                moule_CX = mySqllDataReader8.GetDouble(1);
                            }

                            somme = epais + hauteur_centreCC + hauteur_centreCX + Cales_CC;

                            xlworkSheet.Cells[22, i] = moule_CX - somme;
                        }
                        i += 1;
                        cavité += 1;
                    }
                }
            }

            var Job_precedent = 0;
            var job_actuel = dataGridView1.CurrentRow.Cells["JOB_NB"].Value;
            var job_mere = dataGridView1.CurrentRow.Cells["JOB_NB"].Value;
            var limite_recherche = 0;
            var ins_cc_actuel = "";
            var ins_cc_precedent = "";
            var ins_cx_actuel = "";
            var ins_cx_precedent = "";


            if (cd_press == "P23" || cd_press == "P24")
            {
                con.Close();
                string strRequette15 = "SELECT distinct COMBO_BOM.INS_NB_INS_CX, COMBO_BOM.INS_NB_INS_CV FROM COMBO_JOB_LINES, COMBO_ITEMS, COMBO_BOM where COMBO_ITEMS.LB_LOGI_SKU = COMBO_JOB_LINES.LB_LOGI_SKU and COMBO_BOM.CD_CCE_SKU = COMBO_ITEMS.CD_CCE_SKU and JOB_NB = '" + job_mere + "'";
                con.Open();
                OracleCommand myCommand15 = new OracleCommand(strRequette15, con);
                OracleDataReader mySqlDataReader15 = myCommand15.ExecuteReader();
                while (mySqlDataReader15.Read())
                {
                    ins_cx_actuel = mySqlDataReader15.GetString(0);
                    ins_cc_actuel = mySqlDataReader15.GetString(1);
                }
                con.Close();

                string strRequette17 = "SELECT distinct JOB_NB, CREATION_DATE FROM COMBO_JOB_HEADER_TRACKING where CD_PRESS = '" + cd_press + "' ORDER BY CREATION_DATE";
                con.Open();
                OracleCommand myCommand17 = new OracleCommand(strRequette17, con);
                OracleDataReader mySqlDataReader17 = myCommand17.ExecuteReader();
                while (mySqlDataReader17.Read())
                {
                    job_actuel = mySqlDataReader17.GetInt32(0);
                }
                con.Close();

                while (limite_recherche < 50)
                {
                    //TOUVER LE DERNIER JOBS + Verification
                    string strRequette13 = "SELECT distinct JOB_NB, CREATION_DATE FROM COMBO_JOB_HEADER_TRACKING where CD_PRESS = '" + cd_press + "' and JOB_NB < '" + job_actuel + "' ORDER BY CREATION_DATE";
                    con.Open();
                    OracleCommand myCommand13 = new OracleCommand(strRequette13, con);
                    OracleDataReader mySqlDataReader13 = myCommand13.ExecuteReader();
                    while (mySqlDataReader13.Read())
                    {
                        Job_precedent = mySqlDataReader13.GetInt32(0);
                    }
                    con.Close();

                    string strRequette16 = "SELECT distinct COMBO_BOM.INS_NB_INS_CX, COMBO_BOM.INS_NB_INS_CV FROM COMBO_JOB_LINES, COMBO_ITEMS, COMBO_BOM where COMBO_ITEMS.LB_LOGI_SKU = COMBO_JOB_LINES.LB_LOGI_SKU and COMBO_BOM.CD_CCE_SKU = COMBO_ITEMS.CD_CCE_SKU and JOB_NB = '" + Job_precedent + "'";
                    con.Open();
                    OracleCommand myCommand16 = new OracleCommand(strRequette16, con);
                    OracleDataReader mySqlDataReader16 = myCommand16.ExecuteReader();
                    while (mySqlDataReader16.Read())
                    {
                        ins_cx_precedent = mySqlDataReader16.GetString(0);
                        ins_cc_precedent = mySqlDataReader16.GetString(1);
                    }
                    con.Close();

                    if (Job_precedent != System.Convert.ToInt32(job_mere))
                    {
                        if (ins_cx_actuel == ins_cx_precedent && ins_cc_actuel == ins_cc_precedent)
                        {
                            var result = MessageBox.Show("des inserts CC et CX sont identiques, jobs " + job_mere + " = " + Job_precedent + " \n Utiliser les mêmes inserts ?", "Attention", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        }
                        else if (ins_cc_actuel == ins_cc_precedent)
                        {
                            var result1 = MessageBox.Show("des inserts CC sont identiques, jobs " + job_mere + " = " + Job_precedent + " \n Utiliser les mêmes inserts ?", "Attention", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (result1 == DialogResult.Yes)
                            {
                                break;
                            }
                        }
                    }

                    job_actuel = Job_precedent;

                    limite_recherche += 1;
                }

                MessageBox.Show("fin de boucle");




            }

            conn.Close();
            xlworkbook.SaveAs(destination, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue);
            xlworkbook.Close(true, misValue, misValue);
            xlsp.Quit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
        }

        private void button3_Click1(object sender, EventArgs e)
        {

        }

        private void button6_Click_1(object sender, EventArgs e)
        {
        }
    }

}