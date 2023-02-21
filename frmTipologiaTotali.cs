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
    public partial class frmTipologiaTotali : Form
    {
        OleDbConnection conn = new OleDbConnection();
        string esclusi = " and Spesa.Escludi = 0";

        public frmTipologiaTotali(OleDbConnection pconn)
        {
            InitializeComponent();
            conn = pconn;
        }

        private void frmStatistiche_Load(object sender, EventArgs e)
        {
            esclusi = " and Spesa.Escludi = 0";
            caricaLista();
            caricaGrafico(esclusi);
        }
        private void cmbTipologia_SelectedIndexChanged(object sender, EventArgs e)
        {
            caricaTabella(esclusi);
            somma(esclusi);
            conteggio(esclusi);
            caricaGrafico(esclusi);
        }
        private void checkBoxEsclusi_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxEsclusi.Checked == true)
            {
                esclusi = "";
                lblEsclusi.Visible = false;
            }
            else
            {
                esclusi = " and Spesa.Escludi = 0";
                lblEsclusi.Visible = true;
            }

            if (string.Compare(cmbTipologia.Text, "") != 0)
            {
                caricaTabella(esclusi);
                somma(esclusi);
                conteggio(esclusi);
                caricaGrafico(esclusi);
            }
            else
            {
                caricaGrafico(esclusi);
            }
        }
        private void caricaLista()
        {
            classQuery cl = new classQuery();
            OleDbDataReader readerTipologie = cl.Tipologie(conn);

            while (readerTipologie.Read())
            {
                if (readerTipologie["Nome"] != DBNull.Value)
                {
                    cmbTipologia.Items.Add(readerTipologie["Nome"]);
                }
            }
            conn.Close();
        }

        private void caricaTabella(string esclusi)
        {
            if (string.Compare(cmbTipologia.Text, "") != 0)
            {
                try
                {
                    conn.Open();

                    OleDbCommand SpeseTipologia = new OleDbCommand();
                    SpeseTipologia.Connection = conn;
                    SpeseTipologia.CommandText = "select Tipologia.Nome as TIPOLOGIA, Spesa.Importo as IMPORTO, Data as DATA, Descrizione as DESCRIZIONE from Tipologia, Spesa where Tipologia.ID = Spesa.Tipologia and Tipologia.Nome = '" + cmbTipologia.Text + "' " + esclusi + " order by Data desc";

                    OleDbDataAdapter dataAdap = new OleDbDataAdapter(SpeseTipologia);
                    DataTable dataTab = new DataTable();
                    dataAdap.Fill(dataTab);
                    dgvPerTipologia.DataSource = dataTab;

                    conn.Close();

                    dgvPerTipologia.Columns["IMPORTO"].DefaultCellStyle.Format = "c";
                    dgvPerTipologia.Columns["IMPORTO"].DefaultCellStyle.Font = new Font(dgvPerTipologia.Font, FontStyle.Bold);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Errore nel caricamento della tabella \"Spese per tipologia\" " + ex);
                    conn.Close();
                }
            }
        }

        private void somma(string esclusi)
        {
            try
            {
                conn.Open();

                OleDbCommand Totale = new OleDbCommand();
                Totale.Connection = conn;
                Totale.CommandText = "select sum(Spesa.Importo) as Totale from Spesa, Tipologia where Spesa.Tipologia = Tipologia.ID and Tipologia.Nome = '" + cmbTipologia.Text + "'" + esclusi;
                OleDbDataReader readerTotale = Totale.ExecuteReader();
                readerTotale.Read();

                if (readerTotale["Totale"] != DBNull.Value)
                {
                    double totale = Convert.ToDouble(readerTotale["Totale"]);
                    lblValoreTotale.Text = totale.ToString() + " €";
                }
                else
                {
                    lblValoreTotale.Text = "0 €";
                }

                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore nel calcolo del totale: " + ex);
                conn.Close();
            }
        }

        private void conteggio(string esclusi)
        {
            try
            {
                conn.Open();

                OleDbCommand Conteggio = new OleDbCommand();
                Conteggio.Connection = conn;
                Conteggio.CommandText = "select count(Spesa.ID) as Totale from Spesa, Tipologia where Spesa.Tipologia = Tipologia.ID and Tipologia.Nome = '" + cmbTipologia.Text + "'" + esclusi;
                OleDbDataReader readerConteggio = Conteggio.ExecuteReader();
                readerConteggio.Read();

                lblValoreConteggio.Text = readerConteggio["Totale"].ToString() + " spese";

                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore nel calcolo del conteggio: " + ex);
                conn.Close();
            }
        }
        private void caricaGrafico(string esclusi)
        {
            chartSpese.Series[0].Points.Clear();

            try
            {
                conn.Open();
                OleDbCommand Tipologie = new OleDbCommand();
                Tipologie.Connection = conn;
                Tipologie.CommandText = "select distinct Nome from Tipologia, Spesa where Spesa.Tipologia = Tipologia.ID " + esclusi;
                OleDbDataReader readerTipologie = Tipologie.ExecuteReader();

                while (readerTipologie.Read())
                {
                    OleDbCommand caricaSpese = new OleDbCommand();
                    caricaSpese.Connection = conn;
                    caricaSpese.CommandText = "select sum (Importo) as somma from Spesa, Tipologia where Spesa.Tipologia = Tipologia.ID and Tipologia.Nome = '" + readerTipologie["Nome"] + "'" + esclusi;
                    OleDbDataReader readerSpese = caricaSpese.ExecuteReader();
                    readerSpese.Read();

                    chartSpese.Series[0].Points.AddXY(readerTipologie["Nome"], readerSpese["somma"]);
                }
                conn.Close();

                double tot = 0;
                try
                {    
                    conn.Open();

                    OleDbCommand Totale = new OleDbCommand();
                    Totale.Connection = conn;
                    Totale.CommandText = "select sum(Importo) as Totale from Spesa where 1=1" + esclusi;
                    OleDbDataReader readerTotale = Totale.ExecuteReader();
                    readerTotale.Read();

                    if (readerTotale["Totale"] != DBNull.Value)
                    {
                        tot = Convert.ToDouble(readerTotale["Totale"]);
                    }
                    else
                    {
                        tot = 0;
                    }
                    conn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ERRORE CALCOLO TOTALE: " + ex);
                    conn.Close();
                }
                    
                double tre_per_cento = tot * 3 / 100;

                int n = chartSpese.Series[0].Points.Count;
                for(int i=0; i<n; i++)
                {
                    if(Convert.ToDouble(chartSpese.Series[0].Points[i].YValues[0]) < tre_per_cento)
                    {
                        chartSpese.Series[0].Points[i].LabelForeColor = Color.Transparent;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore nel recupero dei nomi delle Tipologie o nella somma degli importi: " + ex);
                conn.Close();
            }
        }
    }
}