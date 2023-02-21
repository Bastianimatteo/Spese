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
    public partial class frmIstogrammaSpesa : Form
    {
        OleDbConnection conn = new OleDbConnection();

        double[] vettore_spesa = new double[12];
        int i = 0;
        public frmIstogrammaSpesa(OleDbConnection pconn)
        {
            InitializeComponent();
            conn = pconn;
        }

        private void frmStatisticheSpesa_Load(object sender, EventArgs e)
        {
            chartSpese.ChartAreas["ChartArea1"].AxisX.Interval = 1;

            classQuery cl = new classQuery();
            string a = cl.Anno(conn);
            string b = (Convert.ToDouble(a) + 1).ToString();

            string[] vettore_data = { "settembre " + a, "ottobre " + a, "novembre " + a, "dicembre " + a, "gennaio " + b, "febbraio " + b, "marzo " + b, "aprile " + b, "maggio " + b, "giugno " + b, "luglio " + b, "agosto " + b };

            try
            {
                conn.Open();

                for (i = 0; i < 12; i++)
                {
                    OleDbCommand caricaSpese = new OleDbCommand();
                    caricaSpese.Connection = conn;
                    caricaSpese.CommandText = "select sum(Importo) as Valore from Spesa, Tipologia where Spesa.Tipologia  = Tipologia.ID and Tipologia.Nome = 'Spesa' and Format(Data, 'MMMM YY') = '" + vettore_data[i] + "'";
                    OleDbDataReader readerSpese = caricaSpese.ExecuteReader();
                    readerSpese.Read();

                    if(readerSpese["Valore"] != DBNull.Value)
                    {
                        vettore_spesa[i] = Convert.ToDouble(readerSpese["Valore"]);
                    }
                }

                for (i = 0; i < 12; i++)
                {
                    if (vettore_spesa[i] != 0)
                    {
                        chartSpese.Series["Spesa"].Points.AddXY(vettore_data[i], vettore_spesa[i]);
                    }
                }

                conn.Close();
                chartSpese.Series["Spesa"].IsValueShownAsLabel = true;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Errore nel recupero della lista dei mesi o nella somma degli importi:" + ex);
                conn.Close();
            }
        }
    }
}
