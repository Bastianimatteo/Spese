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
    public partial class frmModifica_ultima : Form
    {
        OleDbConnection conn = new OleDbConnection();
        public frmModifica_ultima(OleDbConnection pconn)
        {
            InitializeComponent();
            conn = pconn;
        }

        private void frmModifica_spesa_Load(object sender, EventArgs e)
        {
            string importo, data, descrizione;
            try
            {
                conn.Open();

                OleDbCommand Ultima = new OleDbCommand();
                Ultima.Connection = conn;
                Ultima.CommandText = "select Importo, Format(Data, \"dd/MM/yyyy\") as Data, Descrizione from Spesa where ID = (select max(ID) from Spesa)";
                OleDbDataReader readerUltima = Ultima.ExecuteReader();

                if (readerUltima.Read())
                {
                    importo = readerUltima["Importo"].ToString();
                    data = readerUltima["Data"].ToString();
                    descrizione = readerUltima["Descrizione"].ToString();
                }
                else
                {
                    importo = "";
                    data = "";
                    descrizione = "";
                }

                txtImporto.Text = importo;
                txtData.Text = data;
                txtDescrizione.Text = descrizione;

                conn.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show("Errore nel caricamento dell'importo e della data dell'ultima spesa inserita: " + ex);
                conn.Close();
            }
        }

        private void btnInvia_Click(object sender, EventArgs e)
        {
            if (string.Compare(txtImporto.Text, "") != 0 && string.Compare(txtData.Text, "") != 0 && string.Compare(txtDescrizione.Text, "") != 0)
            {
                try
                {
                    conn.Open();

                    OleDbCommand updateUltima = new OleDbCommand();
                    updateUltima.Connection = conn;
                    updateUltima.CommandText = "update Spesa set Importo= '" + txtImporto.Text + "', Data= '" + txtData.Text + "', Descrizione = '" + txtDescrizione.Text + "' where ID = (select max(ID) from Spesa)";
                    updateUltima.ExecuteNonQuery();

                    conn.Close();
                    MessageBox.Show("Valori dell'ultima spesa inserita aggiornati");

                    this.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Errore nell'aggiornamento dei valori dell'ultima spesa inserita: " + ex);
                    conn.Close();
                }
            }
            else
            {
                MessageBox.Show("Valori non validi");
            }
        }
    }
}
