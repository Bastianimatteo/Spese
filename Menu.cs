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
    public partial class frmMenu : Form
    {
        OleDbConnection conn = new OleDbConnection();
        public frmMenu()
        {
            InitializeComponent();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source= C:\Users\basti\OneDrive\Desktop\Progetti\Spese\Spese.accdb; Persist Security Info = False";
        }

        private void Menu_Load(object sender, EventArgs e)
        {
            caricaLista();
            residuoCarte();

            try
            {
                conn.Open();

                OleDbCommand Tipologia = new OleDbCommand();
                Tipologia.Connection = conn;
                Tipologia.CommandText = "select Nome from Tipologia order by ID";
                OleDbDataReader readerTipologia = Tipologia.ExecuteReader();

                while(readerTipologia.Read())
                {
                    lstTipologia.Items.Add(readerTipologia["Nome"]);
                }

                conn.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show("Errore nel caricamento della lista delle tipologie: " + ex);
                conn.Close();
            }

            txtData.Text = DateTime.Now.ToString("dd/MM/yyyy");
            lstTipologia.SelectedItem = "Spesa";
            chkEscludi.Checked = false;
        }

        private void btnInvia_Click(object sender, EventArgs e)
        {
            int codice_tipologia = 0;

            try
            {
                conn.Open();

                OleDbCommand Id_tipologia = new OleDbCommand();
                Id_tipologia.Connection = conn;
                Id_tipologia.CommandText = "select ID from Tipologia where Nome = '" + lstTipologia.Text + "'";
                OleDbDataReader readerId_tipologia = Id_tipologia.ExecuteReader();

                readerId_tipologia.Read();
                codice_tipologia = Convert.ToInt32(readerId_tipologia["ID"]);

                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore nel recupero del codice Tipologia: " + ex);
                conn.Close();
            }

            if (string.Compare(lstTipologia.Text, "") == 0 || string.Compare(txtImporto.Text, "") == 0 || string.Compare(txtData.Text, "") == 0)
            {
                MessageBox.Show("I campi \"Tipologia\", \"Importo\" e \"Data\" sono obbligatori");
            }
            else
            {
                int escludi = 0;
                if (chkEscludi.Checked == false)
                    escludi = 0;
                else
                    escludi = 1;

                try
                {
                    conn.Open();
                    OleDbCommand Inserisci = new OleDbCommand();
                    Inserisci.Connection = conn;
                    Inserisci.CommandText = "Insert into Spesa (Tipologia, Importo, Data, Descrizione, Escludi) values ('" + codice_tipologia + "', '" + txtImporto.Text + "', '" + txtData.Text + "', '" + txtDescrizione.Text + "', " + escludi + ")";
                    Inserisci.ExecuteNonQuery();

                    MessageBox.Show("Spesa inserita");
                    conn.Close();

                    conn.Open();
                    OleDbCommand aggiornaCarta = new OleDbCommand();
                    aggiornaCarta.Connection = conn;
                    aggiornaCarta.CommandText = "update Carta set Importo = Importo - '" + txtImporto.Text + "'";
                    aggiornaCarta.ExecuteNonQuery();

                    conn.Close();

                    caricaLista();
                    residuoCarte();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ERRORE INSERIMENTO SPESA O AGGIORNAMENTO SALDO: " + ex);
                    conn.Close();
                }
            }
        }

        public void caricaLista()
        {
            try
            {
                conn.Open();

                OleDbCommand SpeseSettimana = new OleDbCommand();
                SpeseSettimana.Connection = conn;
                SpeseSettimana.CommandText = "select Tipologia.Nome as TIPOLOGIA, Importo as IMPORTO, Data as DATA, Descrizione as DESCRIZIONE from Tipologia, Spesa where Tipologia.ID = Spesa.Tipologia and Data between DateAdd('d', -((Weekday(Date(), 2) - 1)), Date()) AND Date() order by Data desc";

                OleDbDataAdapter dataAdap = new OleDbDataAdapter(SpeseSettimana);
                DataTable dataTab = new DataTable();
                dataAdap.Fill(dataTab);
                dgvSpeseSettimana.DataSource = dataTab;

                conn.Close();

                dgvSpeseSettimana.Columns["IMPORTO"].DefaultCellStyle.Format = "c";
                dgvSpeseSettimana.Columns["IMPORTO"].DefaultCellStyle.Font = new Font(dgvSpeseSettimana.Font, FontStyle.Bold);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore nel caricamento della tabella \"Spese della settimana\" " + ex);
                conn.Close();
            }

            //TOTALE SETTIMANA
            try
            {
                double med = media();
                lblMedia.Text = "Media settimanale: " + med + " €";

                conn.Open();
                OleDbCommand Totale = new OleDbCommand();
                Totale.Connection = conn;
                Totale.CommandText = "select sum(Importo) as Totale from Spesa where Data between DateAdd('d', -((Weekday(Date(), 2) -1)), Date()) AND Date()";
                OleDbDataReader readerTotale = Totale.ExecuteReader();
                readerTotale.Read();

                if (readerTotale["Totale"] != DBNull.Value)
                {
                    double valore = Convert.ToDouble(readerTotale["Totale"]);

                    if (valore > med && med != 0)
                    {
                        txtTotale.BackColor = txtTotale.BackColor;
                        txtTotale.ForeColor = Color.Red;
                    }
                    else
                    {
                        txtTotale.BackColor = txtTotale.BackColor;
                        txtTotale.ForeColor = Color.Black;
                    }

                    txtTotale.Text = valore.ToString() + " €";

                    conn.Close();
                }
                else
                {
                    txtTotale.ForeColor = Color.Black;
                    txtTotale.Text = "0,00 €";
                    conn.Close();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Errore nel calcolo del totale: " + ex);
                conn.Close();
            }
        }

        public void residuoCarte()
        {
            conn.Open();
            OleDbCommand Tot_Statale = new OleDbCommand();
            Tot_Statale.Connection = conn;
            Tot_Statale.CommandText = "select Importo from Carta";
            OleDbDataReader reader_Statale = Tot_Statale.ExecuteReader();

            
            reader_Statale.Read();

            double Residuo_Statale = Convert.ToDouble(reader_Statale["Importo"]);

            conn.Close();

            lblStatale.Text = "STATALE: " + Math.Round(Residuo_Statale, 2) + " €";
        }

        //MEDIA
        public double media()
        {
            double media_settimana = 0;
            try
            {
                conn.Close();
                conn.Open();
                OleDbCommand Totale = new OleDbCommand();
                Totale.Connection = conn;
                Totale.CommandText = "select sum(Importo) as Media from Spesa where Escludi = 0";
                OleDbDataReader readerTotale = Totale.ExecuteReader();

                OleDbCommand numeroSettimane = new OleDbCommand();
                numeroSettimane.Connection = conn;
                numeroSettimane.CommandText = "select distinct DateDiff('ww', (select min(Data) from Spesa), (select max(Data) from Spesa)) as Numero from Spesa";
                OleDbDataReader readerSettimane = numeroSettimane.ExecuteReader();

                readerTotale.Read();
                readerSettimane.Read();

                if (Convert.ToDouble(readerTotale["Media"]) != 0 && Convert.ToInt32(readerSettimane["Numero"]) != 0)
                {
                    media_settimana = Convert.ToDouble(readerTotale["Media"]) / Convert.ToDouble(readerSettimane["Numero"]);
                    media_settimana = Math.Round(media_settimana, 2);
                }
                else
                {
                    media_settimana = 0;
                }
                conn.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show("Errore nel calcolo della media: " + ex);
                conn.Close();
            }

            return media_settimana;
        }

        private void ultimaInseritaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmModifica_ultima frmUltima = new frmModifica_ultima(conn);
            frmUltima.Show();
        }

        private void precedenteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmModifica_precedente frmPredecente = new frmModifica_precedente(conn);
            frmPredecente.Show();
        }

        private void scorsaSettimanaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmScorsa_settimana frmScorsa = new frmScorsa_settimana(conn);
            frmScorsa.Show();
        }
        private void spesaPerMeseToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            frmIstogrammaSpesa frmStatSpesa = new frmIstogrammaSpesa(conn);
            frmStatSpesa.Show();
        }

        private void ricaricaCartaToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            frmRicarica frmRica = new frmRicarica(conn);
            frmRica.Show();
        }

        private void perMeseToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            frmPer_mese frmMese1 = new frmPer_mese(conn);
            frmMese1.Show();
        }
        private void medieTotaleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmMedie_Totali frmMed = new frmMedie_Totali(conn);
            frmMed.Show();
        }

        private void questoMESEToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            frmQuesto_mese frmMese = new frmQuesto_mese(conn);
            frmMese.Show();
        }

        private void tOTALIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmTipologiaTotali frmTipTotali = new frmTipologiaTotali(conn);
            frmTipTotali.Show();
        }

        private void mensileistogrammaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmIstogrammaMese frmIstoMese = new frmIstogrammaMese(conn);
            frmIstoMese.Show();
        }
        private void statisticheMensiliToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Medie_mensili frmStat = new Medie_mensili(conn);
            frmStat.Show();
        }

        private void lstTipologia_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch(lstTipologia.SelectedIndex)
            {
                case 0: txtDescrizione.Text = "Esselunga";
                    break;
                case 2: txtDescrizione.Text = "Caffè";
                    break;
                case 3: txtDescrizione.Text = "Lavatrici";
                    break;
                case 4:
                    txtDescrizione.Text = "Macchinette";
                    break;
                case 5:
                    txtDescrizione.Text = "Ricarica cellulare";
                    break;
                case 6:
                    txtDescrizione.Text = "Capelli";
                    break;
                default: txtDescrizione.Text = "";
                    break;
            }
        }

        private void btnAggiorna_Click(object sender, EventArgs e)
        {
            caricaLista();
            residuoCarte();
        }

        private void buttonSottrai_Click(object sender, EventArgs e)
        {
            DateTime data = Convert.ToDateTime(txtData.Text).Subtract(TimeSpan.FromDays(1));
            txtData.Text = data.ToString("dd/MM/yyyy");
            
        }

        private void btnSomma_Click(object sender, EventArgs e)
        {
            DateTime data = Convert.ToDateTime(txtData.Text).Add(TimeSpan.FromDays(1));
            txtData.Text = data.ToString("dd/MM/yyyy");
        }

        private void btnCarrefour_Click(object sender, EventArgs e)
        {
            txtDescrizione.Text = "Carrefour";
        }        
    }
}
