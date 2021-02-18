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
            
            string connString = @"Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=fra2exa01-sxdir1-vip.europe.essilor.group)(PORT=1561)))
                             (CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=PUE1)));User Id=combo;Password=combo;";

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
                Console.WriteLine("Connecté à Oracle");
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = conn;
                cmd.CommandText = "Select JOB_NB FROM COMBO.COMBO_JOB_LINES group by JOB_NB";
                OracleDataReader reader = cmd.ExecuteReader();

                dataGridView1.Rows.Clear();
                DataSet dataset = new DataSet();
                while (reader.Read())
                {
                    dataGridView1.Rows.Add(reader[0]);
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

        
        private void button3_Click(object sender, EventArgs e)
        {

            
            
            ;

            object missValue = System.Reflection.Missing.Value;

            Excel._Application xlsp = new Excel.Application();

            Excel.Workbook xlworkbook = xlsp.Workbooks.Add(missValue);
            Excel.Worksheet xlworkSheet = (Excel.Worksheet)xlworkbook.ActiveSheet;
            xlworkSheet.get_Range("A1", "I30").Borders.Weight = Excel.XlBorderWeight.xlThin;

            xlworkSheet.Cells[1, 1] = "Date";
            xlworkSheet.Cells[1, 5] = "Semaine";
            xlworkSheet.Cells[1, 7] = "Validation préparation";
            xlworkSheet.get_Range("A1", "I1").Font.Size = 10;
            xlworkSheet.get_Range("B1", "D1").Merge(false);
            xlworkSheet.get_Range("G1", "H1").Merge(false);

            xlworkSheet.Cells[2, 7] = "VALIDATION RECEPTION REGLEUR";
            xlworkSheet.get_Range("A2", "I2").Font.Bold = 5;
            xlworkSheet.get_Range("A2:A3", "F2:F3").Merge(false);
            xlworkSheet.get_Range("G3", "I3").Merge(false);

            xlworkSheet.get_Range("A4:A30", "I4:I30").Font.Bold = 5;
            xlworkSheet.Cells[4, 1] = "Job";
            xlworkSheet.get_Range("B4", "C4").Merge(false);
            xlworkSheet.Cells[4, 4] = "Moule";
            xlworkSheet.get_Range("E4", "F4").Merge(false);
            xlworkSheet.Cells[4, 7] = "Presse";
            xlworkSheet.get_Range("H4", "I4").Merge(false);

            xlworkSheet.get_Range("A5", "I5").Merge(false);

            xlworkSheet.Cells[6, 1] = "Produit";
            xlworkSheet.get_Range("B6", "E6").Merge(false);
            xlworkSheet.Cells[6, 6] = "Num Shot";
            xlworkSheet.get_Range("G6", "I6").Merge(false);

            xlworkSheet.get_Range("A7", "I7").Merge(false);

            xlworkSheet.Cells[8, 1] = "Inserts convexes";
            xlworkSheet.get_Range("A8", "I8").Merge(false);

            xlworkSheet.get_Range("A9:A10", "I9:I10").Merge(false);

            xlworkSheet.Cells[111, 1] = "Cavités";
            xlworkSheet.get_Range("A11", "B11").Merge(false);
            xlworkSheet.get_Range("C11", "I11").NumberFormat = "@";
            xlworkSheet.Cells[11, 3] = "01";
            xlworkSheet.Cells[11, 4] = "02";
            xlworkSheet.Cells[11, 5] = "03";
            xlworkSheet.Cells[11, 6] = "04";
            xlworkSheet.Cells[11, 7] = "06";
            xlworkSheet.Cells[11, 8] = "07";
            xlworkSheet.Cells[11, 9] = "08";
            xlworkSheet.Cells[12, 1] = "Bases/Sphère";
            xlworkSheet.get_Range("A12", "B12").Merge(false);
            xlworkSheet.Cells[13, 1] = "Ins CC";
            xlworkSheet.get_Range("A13", "B13").Merge(false);
            xlworkSheet.Cells[14, 1] = "Epais. Centre";
            xlworkSheet.get_Range("A14", "B14").Merge(false);
            xlworkSheet.Cells[15, 1] = "Cales CC";
            xlworkSheet.get_Range("A15", "B15").Merge(false);

            xlworkSheet.get_Range("A16:A17", "I16:I17").Merge(false);

            xlworkSheet.Cells[18, 1] = "CYL";
            xlworkSheet.get_Range("A18", "B18").Merge(false);
            xlworkSheet.Cells[19, 1] = "ins CX";
            xlworkSheet.get_Range("A19", "B19").Merge(false);
            xlworkSheet.Cells[20, 1] = "Epais. Ctr. CX";
            xlworkSheet.get_Range("A20", "B20").Merge(false);
            xlworkSheet.Cells[21, 1] = "Epais verre";
            xlworkSheet.get_Range("A21", "B21").Merge(false);
            xlworkSheet.Cells[22, 1] = "Cales CX";
            xlworkSheet.get_Range("A22", "B22").Merge(false);

            xlworkSheet.get_Range("A23:A24", "I23:I24").Merge(false);

            xlworkSheet.Cells[25, 1] = "Changement d'inserts";
            xlworkSheet.get_Range("A25", "I25").Merge(false);
            xlworkSheet.Cells[26, 1] = "Cavités";
            xlworkSheet.get_Range("A26", "B26").Merge(false);
            xlworkSheet.get_Range("C26", "I26").NumberFormat = "@";
            xlworkSheet.Cells[26, 3] = "01";
            xlworkSheet.Cells[26, 4] = "02";
            xlworkSheet.Cells[26, 5] = "03";
            xlworkSheet.Cells[26, 6] = "04";
            xlworkSheet.Cells[26, 7] = "06";
            xlworkSheet.Cells[26, 8] = "07";
            xlworkSheet.Cells[26, 9] = "08";


            //xlworkSheet.get_Range("A4:A6", "I4:I6").Borders.Weight = Excel.XlBorderWeight.xlThin;

            xlworkbook.SaveAs(@"R:\Commun\ACI\Data\Jobs_Combo\job_" + dataGridView1.CurrentRow.Cells["ORGANIZATION_ID"].Value, Excel.XlSaveAction.xlSaveChanges, missValue, missValue, missValue, missValue, Excel.XlSaveAsAccessMode.xlExclusive, missValue, missValue, missValue, missValue, missValue);
            xlworkbook.Close(true, missValue, missValue);
            xlsp.Quit();



            /*
            string fileName = "Job.xlsx";
            string sourcePath = @"R:\COMMUN\ACI\Jobs\temple";
            string targetPath = @"R:\COMMUN\ACI\Data\Jobs_Combo";

            string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
            string destFile = System.IO.Path.Combine(targetPath, fileName);

            //System.IO.Directory.CreateDirectory(targetPath);

            if(!File.Exists(destFile))
            {
                System.IO.File.Copy(sourceFile, destFile);
              
                MessageBox.Show("Vous avez extrait le fichier " + dataGridView1.CurrentRow.Cells["ORGANIZATION_ID"].Value + "_Job" + " dans " + "R:\\COMMUN\\ACI\\Data\\Jobs_Combo");

                var excelApp = new Excel.Application();

                excelApp.Visible = true;

                Excel._Worksheet workBooks = (Excel.Worksheet)excelApp.ActiveSheet;

                //Ouverture du fichier Excel, à vous de choisir l'emplacement ou est situé le fichier excel ainsi que son nom!!

                Microsoft.Office.Interop.Excel._Workbook workbook = excelApp.Workbooks.Open(@"R:\COMMUN\ACI\Data\Jobs_Combo\Job.xlsx");

                workBooks = workbook.Sheets["Feuil1"]; // On sélectionne la Feuil1

                workBooks = workbook.ActiveSheet;

                //workBooks.Name = "Electronique71.com"; // on renomme la Feuil1 

                //dataGridView1.RowHeadersVisible = false;

                
                    workBooks.Cells[2, 4] = dataGridView1.CurrentRow.Cells["ORGANIZATION_ID"].Value;
                
            }
            else
            {
                MessageBox.Show("Le fichier " +  dataGridView1.CurrentRow.Cells["ORGANIZATION_ID"].Value + "_Job " + " existe déjà");
            }


            */


            //manipulation file et directory
            /*
            string sourceFile = System.IO.Path.Combine(@"R:\COMMUN\ACI\Jobs\temple\Job.xlsx");
            string destFile =   System.IO.Path.Combine(@"R:\COMMUN\ACI\Data\Jobs_Combo\Job" + dataGridView1.CurrentRow.Cells["ORGANIZATION_ID"].Value);
            

            try
            {
                

                // Ensure that the target does not exist.
                if (!File.Exists(destFile))
                {
                    System.IO.File.Copy(sourceFile, destFile, true);
                    MessageBox.Show("Vous avez extrait le fichier " + " Job" + dataGridView1.CurrentRow.Cells["ORGANIZATION_ID"].Value  + " dans " + "R:\\COMMUN\\ACI\\Data\\Jobs_Combo");

                    //File.Delete(destFile);
                    // Move the file.
                                     
                    object missValue = System.Reflection.Missing.Value;

                    Excel.Application appExcel = new Excel.Application();
                    Excel.Workbook xlworkbook = appExcel.Workbooks.Add(missValue);
                    Excel.Worksheet xlworkSheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item(1);

                    xlworkSheet.Cells[4, 2] = dataGridView1.CurrentRow.Cells["ORGANIZATION_ID"].Value;
                    
                    //xlworkbook.SaveAs(destFile,Excel.XlSaveAction.xlSaveChanges, missValue, missValue, missValue, missValue, Excel.XlSaveAsAccessMode.xlNoChange, missValue, missValue, missValue, missValue, missValue);
                    

                    //Fermeture d'Excel
                    //xlworkbook.Close(true, missValue, missValue);
                    
                }
                else
                {
                    MessageBox.Show("Le fichier " + " Job" + dataGridView1.CurrentRow.Cells["ORGANIZATION_ID"].Value + " existe déjà");
                }



            }
            catch (Exception es)
            {
                Console.WriteLine("Erreur", es.ToString());
            }
            
            */

            //System.IO.File.Move(sourceFile, destFile);

            //Extraction des données de la DatagridView vers Excel
            /*
            Excel._Application xlsp;
            Excel.Workbook xlworkbook;
            Excel.Worksheet xlworkSheet;

            object missValue = System.Reflection.Missing.Value;

            xlsp = new Excel.Application();

            xlworkbook = xlsp.Workbooks.Add(missValue);
            xlworkSheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item(1);
            xlworkSheet.Cells[1, 1] = dataGridView1.CurrentRow.Cells["ORGANIZATION_ID"].Value;


            xlworkbook.SaveAs(@"R:\Commun\ACI\Data\Jobs_Combo\", Excel.XlSaveAction.xlSaveChanges, missValue, missValue, missValue, missValue, Excel.XlSaveAsAccessMode.xlNoChange, missValue, missValue, missValue, missValue, missValue);
            
           xlworkbook.Close(true, missValue, missValue);
            */
            // xlsp.Quit();

            /*
            releaseObject(xlworkSheet);
            releaseObject(xlworkbook);
            releaseObject(xlsp);
            */
        }

        /*
        private void releaseObject(object obj)
        {
            
            try
            {
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                //obj = null;
            }
            catch (Exception e)
            {
                obj = null;
                MessageBox.Show("Exception occured while releasing object " + e.ToString());
            }
            finally
            {
                GC.Collect();
            }
            
        }

        */
    }
}