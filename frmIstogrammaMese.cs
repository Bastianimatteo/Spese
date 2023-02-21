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
    public partial class frmIstogrammaMese : Form
    {
        OleDbConnection conn;
        string escludi;

        double[] vettore_spesa = new double[12];
        int i = 0;
        public frmIstogrammaMese(OleDbConnection pconn)
        {
            InitializeComponent();
            conn = pconn;
        }

        private void frmIstogrammaMese_Load(object sender, EventArgs e)
        {
            chkEscludi.Checked = false;
            escludi = "and Escludi = 0";
            carica(escludi);
            chartMese.ChartAreas["ChartArea1"].AxisX.Interval = 1;
        }

        private void chkEscludi_CheckedChanged(object sender, EventArgs e)
        {

            if(chkEscludi.Checked == true)
            {
                escludi = "";
                carica(escludi);
            }
            else
            {
                escludi = "and Escludi = 0";
                carica(escludi);
            }
        }

        private void carica(string escludi)
        {
            chartMese.Series["Mese"].Points.Clear();

            classQuery cl = new classQuery();
            string a = cl.Anno(conn);
            string b = (Convert.ToDouble(a) + 1).ToString();

            string[] vettore_data = { "settembre " + a, "ottobre " + a, "novembre " + a, "dicembre " + a, "gennaio " + b, "febbraio " + b, "marzo " + b, "aprile " + b, "maggio " + b, "giugno " + b, "luglio " + b, "agosto " + b};

            try
            {
                conn.Open();

                for (i = 0; i < 12; i++)
                { 
                    OleDbCommand Totale = new OleDbCommand();
                    Totale.Connection = conn;
                    Totale.CommandText = "select sum(Importo) as Valore from Spesa where Format(Data, 'MMMM YY') = '" + vettore_data[i] + "'" + escludi;
                    OleDbDataReader readerTotale = Totale.ExecuteReader();
                    readerTotale.Read();

                    if (readerTotale["Valore"] != DBNull.Value)
                    {
                        vettore_spesa[i] = Convert.ToDouble(readerTotale["Valore"]);
                    }
                }

                for (i = 0; i < 12; i++)
                {
                    if (vettore_spesa[i] != 0)
                    {
                        chartMese.Series["Mese"].Points.AddXY(vettore_data[i], vettore_spesa[i]);
                    }
                }                

                conn.Close();
                chartMese.Series["Mese"].IsValueShownAsLabel = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore nel caricamento della lista dei mesi o nel recuper del totale: " + ex);
                conn.Close();
            }
        }
    }
}
