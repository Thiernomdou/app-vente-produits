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
            string fileName = "Job.xlsx";
            string sourcePath = @"R:\COMMUN\ACI\Jobs\temple";
            string targetPath = @"R:\COMMUN\ACI\Data\Jobs_Combo";

            string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
            string destFile = System.IO.Path.Combine(targetPath, dataGridView1.CurrentRow.Cells["ORGANIZATION_ID"].Value+"_"+fileName);

            //System.IO.Directory.CreateDirectory(targetPath);

            if(!File.Exists(destFile))
            {
                System.IO.File.Copy(sourceFile, destFile);
              
                MessageBox.Show("Vous avez extrait le fichier " + dataGridView1.CurrentRow.Cells["ORGANIZATION_ID"].Value + "_Job" + " dans " + "R:\\COMMUN\\ACI\\Data\\Jobs_Combo");

                object missValue = System.Reflection.Missing.Value;

                Excel.Application appExcel = new Excel.Application();
                Excel.Workbook xlworkbook = appExcel.Workbooks.Add(missValue);
                Excel.Worksheet xlworkSheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item(1);

                xlworkSheet.Cells[4, 2] = dataGridView1.CurrentRow.Cells["ORGANIZATION_ID"].Value;
                xlworkSheet.Cells[1, 2] = dataGridView1.CurrentRow.Cells["ORGANIZATION_ID"].Value;
                xlworkSheet.Cells[30, 1] = dataGridView1.CurrentRow.Cells["ORGANIZATION_ID"].Value;


                //xlworkbook.SaveAs(@"R:\COMMUN\ACI\Data\Jobs_Combo" + dataGridView1.CurrentRow.Cells["ORGANIZATION_ID"].Value + "_" + fileName, Excel.XlSaveAction.xlSaveChanges, missValue, missValue, missValue, missValue, Excel.XlSaveAsAccessMode.xlNoChange, missValue, missValue, missValue, missValue, missValue);

            }
            else
            {
                MessageBox.Show("Le fichier " +  dataGridView1.CurrentRow.Cells["ORGANIZATION_ID"].Value + "_Job " + " existe déjà");
            }





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