using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;

namespace Spese
{
    public partial class frmRicarica : Form
    {
        OleDbConnection conn = new OleDbConnection();
        public frmRicarica(OleDbConnection pconn)
        {
            InitializeComponent();
            conn = pconn;
        }

        private void frmRicarica_Load(object sender, EventArgs e)
        {
            conn.Open();
            OleDbCommand caricaCarte = new OleDbCommand();
            caricaCarte.Connection = conn;
            caricaCarte.CommandText = "select Nome from Carta";
            OleDbDataReader readerCarte = caricaCarte.ExecuteReader();

            while(readerCarte.Read())
            {
                lstCarta.Items.Add(readerCarte["Nome"]);
            }
            conn.Close();

            lstCarta.SelectedIndex = 0;
        }

        private void btnInvia_Click(object sender, EventArgs e)
        {
            if (string.Compare(lstCarta.Text, "") == 0 || string.Compare(txtImporto.Text, "") == 0)
            {
                MessageBox.Show("Tutti i campi sono obbligatori");
            }
            else
            {
                try
                {
                    conn.Open();
                    OleDbCommand valore = new OleDbCommand();
                    valore.Connection = conn;
                    valore.CommandText = "select Importo from Carta where Nome = '" + lstCarta.Text + "'";
                    OleDbDataReader reader_valore = valore.ExecuteReader();

                    reader_valore.Read();

                    double valore_attuale = Convert.ToDouble(reader_valore["Importo"]);
                    double valore_nuovo = valore_attuale + Convert.ToDouble(txtImporto.Text);

                    OleDbCommand ricarica = new OleDbCommand();
                    ricarica.Connection = conn;
                    ricarica.CommandText = "update Carta set Importo = '" + valore_nuovo + "' where Nome = '" + lstCarta.Text + "'";
                    ricarica.ExecuteNonQuery();

                    conn.Close();

                    MessageBox.Show("Valore aggiornato");
                    this.Close();

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Errore nella ricarica: " + ex);
                    conn.Close();
                }
            }
        }
    }
}
