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
    public partial class frmPer_mese : Form
    {
        OleDbConnection conn = new OleDbConnection();
        string esclusi = "";
        public frmPer_mese(OleDbConnection pconn)
        {
            InitializeComponent();
            conn = pconn;
        }

        private void frmPer_mese_Load(object sender, EventArgs e)
        {
            try
            {
                conn.Open();

                OleDbCommand ListaMesi = new OleDbCommand();
                ListaMesi.Connection = conn;
                ListaMesi.CommandText = "select distinct Format(Data, \"MMMM\") as Mese from Spesa";
                OleDbDataReader readerMesi = ListaMesi.ExecuteReader();

                while(readerMesi.Read())
                {
                    cmbMese.Items.Add(readerMesi["Mese"]);
                }

                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore nel caricamento della lista dei mesi: " + ex);
                conn.Close();
            }
        }
        private void checkBoxEsclusi_CheckedChanged(object sender, EventArgs e)
        {
            if (string.Compare(cmbMese.Text, "") != 0)
            {
                if (checkBoxEsclusi.Checked == true)
                {
                    esclusi = "";
                    lblEsclusi.Visible = false;
                    caricaTabella(esclusi);
                    somma(esclusi);
                    caricaGrafico(esclusi);
                }
                else
                {
                    esclusi = " and Spesa.Escludi = 0";
                    lblEsclusi.Visible = true;
                    caricaTabella(esclusi);
                    somma(esclusi);
                    caricaGrafico(esclusi);
                }
            }
        }
        private void cmbMese_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.Compare(cmbMese.Text, "") != 0)
            {
                checkBoxEsclusi.Checked = false;

                esclusi = " and Spesa.Escludi = 0";
                caricaTabella(esclusi);
                somma(esclusi);
                caricaGrafico(esclusi);
            }
        }
        private void caricaTabella(string esclusi)
        {
            try
            {
                conn.Open();

                OleDbCommand SpeseMese = new OleDbCommand();
                SpeseMese.Connection = conn;
                SpeseMese.CommandText = "select Tipologia.Nome as TIPOLOGIA, Spesa.Importo as IMPORTO, Data as DATA, Descrizione as DESCRIZIONE from Tipologia, Spesa, Carta where Tipologia.ID = Spesa.Tipologia and Format(Data, \"mmmm\") = '" + cmbMese.Text + "'" + esclusi + " order by Data desc";

                OleDbDataAdapter dataAdap = new OleDbDataAdapter(SpeseMese);
                DataTable dataTab = new DataTable();
                dataAdap.Fill(dataTab);
                dgvPerMese.DataSource = dataTab;

                conn.Close();

                dgvPerMese.Columns["IMPORTO"].DefaultCellStyle.Format = "c";
                dgvPerMese.Columns["IMPORTO"].DefaultCellStyle.Font = new Font(dgvPerMese.Font, FontStyle.Bold);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore nel caricamento della tabella \"Spese del mese\" " + ex);
                conn.Close();
            }
        }
        private void somma(string esclusi)
        {
            string anno;

            classQuery cl = new classQuery();
            string a = cl.Anno(conn);
            string b = (Convert.ToDouble(a) + 1).ToString();

            if (cmbMese.Text == "settembre" || cmbMese.Text == "ottobre" || cmbMese.Text == "novembre" || cmbMese.Text == "dicembre")
            {
                anno = a;
            }
            else
            {
                anno = b;
            }

            try
            {
                conn.Open();

                OleDbCommand Totale = new OleDbCommand();
                Totale.Connection = conn;
                Totale.CommandText = "select sum(Spesa.Importo) as Totale from Spesa where Data between Dateserial(" + anno + ", month('01/" + cmbMese.Text + "/1970'), 1) and Dateserial(" + anno + ", month('01/" + cmbMese.Text + "/1970'), 31)" + esclusi;
                OleDbDataReader readerTotale = Totale.ExecuteReader();
                readerTotale.Read();

                double totale = Convert.ToDouble(readerTotale["Totale"]);
                lblValore.Text = totale.ToString() + " €";

                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore nel calcolo del totale: " + ex);
                conn.Close();
            }
        }

        private void caricaGrafico(string esclusi)
        {
            chartSpese.Series["0"].Points.Clear();

            try
            {
                conn.Open();

                OleDbCommand Tipologie = new OleDbCommand();
                Tipologie.Connection = conn;
                Tipologie.CommandText = "select distinct Nome from Tipologia, Spesa where Spesa.Tipologia = Tipologia.ID";
                OleDbDataReader readerTipologie = Tipologie.ExecuteReader();

                while (readerTipologie.Read())
                {
                    OleDbCommand caricaSpese = new OleDbCommand();
                    caricaSpese.Connection = conn;
                    caricaSpese.CommandText = "select sum (Importo) as somma from Spesa, Tipologia where Spesa.Tipologia = Tipologia.ID and Tipologia.Nome = '" + readerTipologie["Nome"] + "' and Format(Data, \"mmmm\") = '" + cmbMese.Text + "'" + esclusi + "";

                    OleDbDataReader readerSpese = caricaSpese.ExecuteReader();
                    readerSpese.Read();

                    chartSpese.Series["0"].Points.AddXY(readerTipologie["Nome"], readerSpese["somma"]);
                }

                conn.Close();

                double tot = 0;
                string anno;

                classQuery cl = new classQuery();
                string a = cl.Anno(conn);
                string b = (Convert.ToDouble(a) + 1).ToString();

                if (cmbMese.Text == "settembre" || cmbMese.Text == "ottobre" || cmbMese.Text == "novembre" || cmbMese.Text == "dicembre")
                {
                    anno = a;
                }
                else
                {
                    anno = b;
                }

                try
                {
                    conn.Open();

                    OleDbCommand Totale = new OleDbCommand();
                    Totale.Connection = conn;
                    Totale.CommandText = "select sum(Spesa.Importo) as Totale from Spesa where Data between Dateserial(" + anno + ", month('01/" + cmbMese.Text + "/1970'), 1) and Dateserial(" + anno + ", month('01/" + cmbMese.Text + "/1970'), 31)" + esclusi + "" + esclusi;
                    OleDbDataReader readerTotale = Totale.ExecuteReader();
                    readerTotale.Read();

                    tot = Convert.ToDouble(readerTotale["Totale"]);

                    conn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ERRORE CALCOLO TOTALE " + ex);
                }

                double tre_per_cento = tot * 3 / 100;

                int n = chartSpese.Series[0].Points.Count;
                for (int i = 0; i < n; i++)
                {
                    if (Convert.ToDouble(chartSpese.Series[0].Points[i].YValues[0]) < tre_per_cento)
                    {
                        chartSpese.Series[0].Points[i].LabelForeColor = Color.Transparent;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore nel caricamento del grafico a torta: " + ex);
                conn.Close();
            }
        }
    }
}
