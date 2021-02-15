using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Calage_Inserts
{
    public partial class Accueil2 : Form
    {
        public Accueil2(string Login)
        {
            InitializeComponent();
            label1.Text = Login;
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
            extraction_jobs.Visible = false;
            retirer_insert.Visible = false;

        }

        private void Suivie_scrap_Click(object sender, EventArgs e)
        {

        }

        private void Remplir_une_feuille_de_calage_Click(object sender, EventArgs e)
        {
            this.Hide();
            Remplir_Feuille_Calage r = new Remplir_Feuille_Calage();
            r.Show();
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

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Connexion c = new Connexion();
            c.Show();
        }
    }
}
