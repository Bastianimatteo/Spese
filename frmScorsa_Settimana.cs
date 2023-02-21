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
    public partial class frmScorsa_settimana : Form
    {
        OleDbConnection conn = new OleDbConnection();
        public frmScorsa_settimana(OleDbConnection pconn)
        {
            InitializeComponent();
            conn = pconn;
        }

        private void frmScorsa_Settimana_Load(object sender, EventArgs e)
        {
            int continua = somma();
            if (continua == 0)
            {
                try
                {
                    conn.Open();

                    OleDbCommand SpeseScorsaSettimana = new OleDbCommand();
                    SpeseScorsaSettimana.Connection = conn;
                    SpeseScorsaSettimana.CommandText = "select Tipologia.Nome as TIPOLOGIA, Spesa.Importo as IMPORTO, Data as DATA, Descrizione as DESCRIZIONE, Escludi as ESCLUDI from Tipologia, Spesa where Tipologia.ID = Spesa.Tipologia and Data between DateAdd('d', -(WeekDay(Date(), 2) -1), Date())-7 and DateAdd('d', -(WeekDay(Date(), 2) -1), Date())-1 order by Data desc";

                    OleDbDataAdapter dataAdap = new OleDbDataAdapter(SpeseScorsaSettimana);
                    DataTable dataTab = new DataTable();
                    dataAdap.Fill(dataTab);
                    dgvScorsaSettimana.DataSource = dataTab;

                    conn.Close();

                    dgvScorsaSettimana.Columns["IMPORTO"].DefaultCellStyle.Format = "c";
                    dgvScorsaSettimana.Columns["IMPORTO"].DefaultCellStyle.Font = new Font(dgvScorsaSettimana.Font, FontStyle.Bold);

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Errore nel caricamento della tabella \"Spese della scorsa settimana\" " + ex);
                    conn.Close();
                }

                try
                {
                    conn.Open();

                    OleDbCommand Totale = new OleDbCommand();
                    Totale.Connection = conn;
                    Totale.CommandText = "select sum(Importo) as Totale from Spesa where Data between DateAdd('d', -(WeekDay(Date(), 2) -1), Date())-7 and DateAdd('d', -(WeekDay(Date(), 2) -1), Date())-1";
                    OleDbDataReader readerTotale = Totale.ExecuteReader();
                    readerTotale.Read();

                    lblValore.Text = readerTotale["Totale"].ToString() + " €";

                    conn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Errore nel calcolo del totale: " + ex);
                    conn.Close();
                }
            }
            else
            {
                this.Close();
            }
        }

        public void btnChiudi_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private int somma()
        {
            try
            {
                conn.Open();

                OleDbCommand Totale = new OleDbCommand();
                Totale.Connection = conn;
                Totale.CommandText = "select sum(Spesa.Importo) as Totale from Spesa where Data between DateAdd('d', -(WeekDay(Date(), 2) -1), Date())-7 and DateAdd('d', -(WeekDay(Date(), 2) -1), Date())-1";
                OleDbDataReader readerTotale = Totale.ExecuteReader();
                readerTotale.Read();

                if (readerTotale["Totale"] != DBNull.Value)
                {
                    double totale = Convert.ToDouble(readerTotale["Totale"]);
                    lblValore.Text = totale.ToString() + " €";
                    conn.Close();
                    return 0;
                }
                else
                {
                    MessageBox.Show("La scorsa settimana non sono state effettuate spese");
                    conn.Close();
                    return 1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore nel calcolo del totale: " + ex);
                conn.Close();
                return 1;
            }

        }
    }
}
