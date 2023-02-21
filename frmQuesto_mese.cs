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
    public partial class frmQuesto_mese : Form
    {
        OleDbConnection conn = new OleDbConnection();
        string esclusi = " and Spesa.Escludi = 0";
        public frmQuesto_mese(OleDbConnection pconn)
        {
            InitializeComponent();
            conn = pconn;
        }

        private void frmQuesto_mese_Load(object sender, EventArgs e)
        {
            int continua = somma(esclusi);

            if (continua == 0)
            {
                caricaLista(esclusi);
                caricaGrafico(esclusi);
            }
            else
            {
                this.Close();
            }
        }

        private void btnChiudi_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void checkBoxEsclusi_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBoxEsclusi.Checked == true)
            {
                esclusi = "";
                lblEsclusi.Visible = false;
                somma(esclusi);
                caricaLista(esclusi);
                caricaGrafico(esclusi);
            }
            else
            {
                esclusi = " and Spesa.Escludi = 0";
                lblEsclusi.Visible = true;
                caricaLista(esclusi);
                somma(esclusi);
                caricaGrafico(esclusi);
            }
        }

        private void caricaLista(string esclusi)
        {
            try
            {
                conn.Open();

                OleDbCommand SpeseMese = new OleDbCommand();
                SpeseMese.Connection = conn;
                SpeseMese.CommandText = "select Tipologia.Nome as TIPOLOGIA, Spesa.Importo as IMPORTO, Data as DATA, Descrizione as DESCRIZIONE from Tipologia, Spesa where Tipologia.ID = Spesa.Tipologia and Data between Dateserial(year(date()), month(date()), 1) and Dateserial(year(Date()), month(Date()), day(Date())) " + esclusi + " order by Data desc";

                OleDbDataAdapter dataAdap = new OleDbDataAdapter(SpeseMese);
                DataTable dataTab = new DataTable();
                dataAdap.Fill(dataTab);
                dgvQuestoMese.DataSource = dataTab;

                conn.Close();

                dgvQuestoMese.Columns["IMPORTO"].DefaultCellStyle.Format = "c";
                dgvQuestoMese.Columns["IMPORTO"].DefaultCellStyle.Font = new Font(dgvQuestoMese.Font, FontStyle.Bold);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore nel caricamento della tabella \"Spese del mese\" " + ex);
                conn.Close();
            }
        }
        private int somma(string esclusi)
        {
            try
            {
                conn.Open();

                OleDbCommand Totale = new OleDbCommand();
                Totale.Connection = conn;
                Totale.CommandText = "select sum(Spesa.Importo) as Totale from Spesa where Data between Dateserial(year(date()), month(Date()), 1) and Dateserial(year(Date()), month(Date()), day(Date()))" + esclusi + "";
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
                    MessageBox.Show("Questo mese non sono state effettuate spese");
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
                    caricaSpese.CommandText = "select sum (Importo) as somma from Spesa, Tipologia where Spesa.Tipologia = Tipologia.ID and Tipologia.Nome = '" + readerTipologie["Nome"] + "' and Data between Dateserial(year(date()), month(Date()), 1) and Dateserial(year(Date()), month(Date()), day(Date())) " + esclusi + "";
                    OleDbDataReader readerSpese = caricaSpese.ExecuteReader();
                    readerSpese.Read();

                    chartSpese.Series["0"].Points.AddXY(readerTipologie["Nome"], readerSpese["somma"]);
                }

                conn.Close();

                double tot = 0;

                try
                {
                    conn.Open();

                    OleDbCommand Totale = new OleDbCommand();
                    Totale.Connection = conn;
                    Totale.CommandText = "select sum(Spesa.Importo) as Totale from Spesa where Data between Dateserial(year(date()), month(Date()), 1) and Dateserial(year(Date()), month(Date()), day(Date()))" + esclusi + "";
                    OleDbDataReader readerTotale = Totale.ExecuteReader();
                    readerTotale.Read();

                    tot = Convert.ToDouble(readerTotale["Totale"]);

                    conn.Close();
                }
                catch(Exception ex)
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
                MessageBox.Show("ERRORE CARICAMENTO GRAFICO: " + ex);
                conn.Close();
            }
        }
    }

}
