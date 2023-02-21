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
    public partial class frmMedie_Totali : Form
    {
        OleDbConnection conn = new OleDbConnection();
        string escludi;
        public frmMedie_Totali(OleDbConnection pconn)
        {
            InitializeComponent();
            conn = pconn;
        }

        private void frmMedia_Load(object sender, EventArgs e)
        {
            chkEscludi.Checked = false;
            escludi = "where Escludi = 0";

            carica(escludi);
        }

        private void carica(string escludi)
        {
            double tot = 0, media_settimana = 0, media_mese = 0, num_settimane = 0, num_mesi = 0;
            try
            {
                conn.Open();

                OleDbCommand Totale = new OleDbCommand();
                Totale.Connection = conn;
                Totale.CommandText = "select sum(Importo) as Totale from Spesa " + escludi;
                OleDbDataReader readerTotale = Totale.ExecuteReader();

                OleDbCommand numeroSettimane = new OleDbCommand();
                numeroSettimane.Connection = conn;
                numeroSettimane.CommandText = "select distinct DateDiff('ww', (select min(Data) from Spesa), (select max(Data) from Spesa)) as Numero from Spesa";
                OleDbDataReader readerSettimane = numeroSettimane.ExecuteReader();

                OleDbCommand numeroMesi = new OleDbCommand();
                numeroMesi.Connection = conn;
                numeroMesi.CommandText = "select distinct DateDiff('m', (select min(Data) from Spesa), (select max(Data) from Spesa)) as Numero from Spesa";
                OleDbDataReader readerMesi = numeroMesi.ExecuteReader();

                readerTotale.Read();
                readerSettimane.Read();
                readerMesi.Read();

                if (Convert.ToInt32(readerTotale["Totale"]) != 0) //VERIFICO TOT 
                {
                    tot = Convert.ToDouble(readerTotale["Totale"]);
                    tot = Math.Round(tot, 2);

                    if (Convert.ToInt32(readerSettimane["Numero"]) != 0) //VERIFICO SETTIMANA
                    {
                        num_settimane = Convert.ToDouble(readerSettimane["Numero"]);
                        media_settimana = tot / num_settimane;
                        media_settimana = Math.Round(media_settimana, 2);
                    }
                    else // 0 SETTIMANE
                    {
                        num_settimane = 0;
                        media_settimana = 0;
                    }

                    if (Convert.ToInt32(readerMesi["Numero"]) != 0) //VERIFICO MESE
                    {
                        num_mesi = Convert.ToDouble(readerMesi["Numero"]);
                        media_mese = tot / num_mesi;
                        media_mese = Math.Round(media_mese, 2);
                    }
                    else // 0 MESI
                    {
                        num_mesi = 0;
                        media_mese = 0;
                    }
                }
                else // 0 TOTALE
                {
                    tot = 0;
                }

                txtSett.Text = media_settimana + " €";
                lblSett.Text = "N° settimane: " + num_settimane;

                txtMes.Text = media_mese + " €";
                lblMesi.Text = "N° mesi: " + num_mesi;

                txtTot.Text = tot + " €";

                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore nel calcolo della media totale/numero mesi/numero settimane: " + ex);
                conn.Close();
            }
        }

        private void chkEscludi_CheckedChanged(object sender, EventArgs e)
        {
            if (chkEscludi.Checked == true)
            {
                escludi = "";
                carica(escludi);
            }
            else
            {
                escludi = "where Escludi = 0";
                carica(escludi);
            }
        }
    }
}
