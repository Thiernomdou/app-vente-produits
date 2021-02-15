using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using Microsoft.VisualBasic;
using System.Configuration;

namespace Calage_Inserts
{
    public partial class Remplir_Feuille_Calage : Form
    {



        public Remplir_Feuille_Calage()
        {
            InitializeComponent();

        }

        private void DefinirColonnes()
        {

        }


        private void button_start_Click(object sender, EventArgs e)
        {

        }

        private void button__Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Suivie_scrap_Click(object sender, EventArgs e)
        {

        }

        private void Remplir_une_feuille_de_calage_Click(object sender, EventArgs e)
        {

        }

        private void Rechercher_une_feuille_Click(object sender, EventArgs e)
        {

        }

        private void Stock_Click(object sender, EventArgs e)
        {

        }

        private void Remplacer_un_insert_Click(object sender, EventArgs e)
        {

        }

        private void Ajouter_un_insert_Click(object sender, EventArgs e)
        {

        }

        private void extraction_jobs_Click(object sender, EventArgs e)
        {

        }

        private void retirer_insert_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            Connexion c = new Connexion();
            c.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }



        private void button2_Click(object sender, EventArgs e)
        {

            Excel._Application xlsp;
            Excel.Workbook xlworkbook;
            Excel.Worksheet xlworkSheet;

            object missValue = System.Reflection.Missing.Value;

            xlsp = new Excel.Application();

            xlworkbook = xlsp.Workbooks.Add(missValue);
            xlworkSheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item(1);
            xlworkSheet.Cells[1, 1] = "Date";
            xlworkSheet.Cells[1, 2] = DateTime.Now.ToString("dd MMMM yyyy");
            xlworkSheet.Cells[1, 4] = "Semaine";
            xlworkSheet.Cells[1, 5] = DateTime.Now.DayOfWeek;
            xlworkSheet.Cells[3, 1] = "Job";
            xlworkSheet.Cells[3, 2] = textBox1.Text;

            xlworkbook.SaveAs(@"R:\Commun\ACI\Data\Feuilles_de_calage\job_"+textBox1.Text, Excel.XlSaveAction.xlSaveChanges, missValue, missValue, missValue, missValue, Excel.XlSaveAsAccessMode.xlExclusive, missValue, missValue, missValue, missValue, missValue);
            xlworkbook.Close(true, missValue, missValue);
            xlsp.Quit();

            releaseObject(xlworkSheet);
            releaseObject(xlworkbook);
            releaseObject(xlsp);
        }



        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
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

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Accueil r = new Accueil("");
            r.Show();
        }
    }
}