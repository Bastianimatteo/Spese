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
    public partial class Medie_mensili : Form
    {
        OleDbConnection conn = new OleDbConnection();
        double[] vettore_media = new double[5];
        double[] vettore_tot_mese = new double[12];

        public Medie_mensili(OleDbConnection pconn)
        {
            InitializeComponent();
            conn = pconn;
        }

        private void frmStatistiche_mensili_Load(object sender, EventArgs e)
        {
            string[] vettore_tipologia = { "Spesa", "Mangiare fuori", "Caffe", "Festa", "Abbigliamento" };
            int i,j;
            classQuery cl = new classQuery();

            string a = cl.Anno(conn);
            string b = (Convert.ToDouble(a) + 1).ToString();

            string[] vettore_data = { "Settembre " + a, "Ottobre " + a, "Novembre " + a, "Dicembre " + a, "Gennaio " + b, "Febbraio " + b, "Marzo " + b, "Aprile " + b, "Maggio " + b, "Giugno " + b, "Luglio " + b, "Agosto " + b };

            for (i = 0; i < vettore_tipologia.Length; i++)
            {
                vettore_media[i]= cl.Media_tipologia(conn, vettore_tipologia[i]);

                switch(i)
                {
                    case 0: lblSpesa.Text = "SPESA: " + Math.Round(vettore_media[i], 2) + " €"; break;
                    case 1: lblMangiareFuori.Text = "MANGIARE FUORI: " + Math.Round(vettore_media[i], 2) + " €"; break;
                    case 2: lblCaffè.Text = "CAFFÈ: " + Math.Round(vettore_media[i], 2) + " €"; break;
                    case 3: lblFesta.Text = "FESTA: " + Math.Round(vettore_media[i], 2) + " €"; break;
                    case 4: lblAbbigliamento.Text = "ABBIGLIAMENTO: " + Math.Round(vettore_media[i], 2) + " €"; break;
                }
            }

            for (i = 0; i < 5; i++)
            {
                for (j = 0; j < 12; j++)
                {
                    OleDbDataReader reader_tot_mese = cl.Mese_tipologia(conn, vettore_tipologia[i], vettore_data[j]);
                    reader_tot_mese.Read();
                    if (reader_tot_mese["tot"] != DBNull.Value)
                    {
                        vettore_tot_mese[j] = Convert.ToDouble(reader_tot_mese["tot"]);
                    }
                    else
                    {
                        vettore_tot_mese[j] = 0;
                    }
                    conn.Close();

                    switch(i)
                    {
                        case 0:
                            if (vettore_tot_mese[j] > vettore_media[i])
                                lstSpesa.Items.Add(vettore_data[j] + " --- " + vettore_tot_mese[j] + " €");
                            break;

                        case 1:
                            if (vettore_tot_mese[j] > vettore_media[i])
                                lstMangiareFuori.Items.Add(vettore_data[j] + " --- " + vettore_tot_mese[j] + " €");
                            break;

                        case 2:
                            if (vettore_tot_mese[j] > vettore_media[i])
                                lstCaffè.Items.Add(vettore_data[j] + " --- " + vettore_tot_mese[j] + " €");
                            break;

                        case 3:
                            if (vettore_tot_mese[j] > vettore_media[i])
                                lstFesta.Items.Add(vettore_data[j] + " --- " + vettore_tot_mese[j] + " €");
                            break;

                        case 4:
                            if (vettore_tot_mese[j] > vettore_media[i])
                                lstAbbigliamento.Items.Add(vettore_data[j] + " --- " + vettore_tot_mese[j] + " €");
                            break;
                    }
                }
            }
        }

        private void btnIStogrammaSpesa_Click(object sender, EventArgs e)
        {
            frmIstogrammaSpesa frmIstSpesa = new frmIstogrammaSpesa(conn);
            frmIstSpesa.Show();
        }
    }
}
