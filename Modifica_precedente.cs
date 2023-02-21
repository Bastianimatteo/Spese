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
    public partial class frmModifica_precedente : Form
    {
        OleDbConnection conn = new OleDbConnection();
        public frmModifica_precedente(OleDbConnection pconn)
        {
            InitializeComponent();
            conn = pconn;
        }

        private void frmModifica_precedente_Load(object sender, EventArgs e)
        {
            txtID.ReadOnly = true;

            try
            {
                conn.Open();

                OleDbCommand Spese = new OleDbCommand();
                Spese.Connection = conn;
                Spese.CommandText = "select top 33 Spesa.ID as ID, Tipologia.Nome as TIPOLOGIA, Importo as IMPORTO, Data as DATA, Descrizione as DESCRIZIONE, Escludi as ESCLUDI from Tipologia, Spesa where Tipologia.ID = Spesa.Tipologia order by (Data) desc";

                OleDbDataAdapter dataAdap = new OleDbDataAdapter(Spese);
                DataTable dataTab = new DataTable();
                dataAdap.Fill(dataTab);
                dgvSpese.DataSource = dataTab;

                conn.Close();

                dgvSpese.Columns["IMPORTO"].DefaultCellStyle.Format = "c";
                dgvSpese.Columns["IMPORTO"].DefaultCellStyle.Font = new Font(dgvSpese.Font, FontStyle.Bold);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore nel caricamento della tabella \"Spese\" " + ex);
                conn.Close();
            }
        }

        private void btnInvia_Click(object sender, EventArgs e)
        {
            if(string.Compare(txtID.Text, "") == 0 || string.Compare(txtImporto.Text, "") == 0 || string.Compare(txtData.Text, "") == 0)
            {
                MessageBox.Show("I campi Importo e Data sono obbligatori");
            }
            else
            {
                int escludi;
                if (chkEscludi.Checked == false)
                    escludi = 0;
                else
                    escludi = 1;

                try
                {
                    conn.Open();

                    OleDbCommand updatePrecedente = new OleDbCommand();
                    updatePrecedente.Connection = conn;
                    updatePrecedente.CommandText = "update Spesa set Importo = '" + txtImporto.Text + "', Data = '" + txtData.Text + "', Descrizione = '" + txtDescrizione.Text + "', Escludi = '" + escludi + "' where ID = " + txtID.Text + "";
                    updatePrecedente.ExecuteNonQuery();

                    conn.Close();
                    MessageBox.Show("Valori della spesa aggiornati");

                    this.Close();
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Errore nell'aggiornamento dei valori della spesa inserita: " + ex);
                    conn.Close();
                }
            }
        }

        private void dgvSpese_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string valore = dgvSpese.Rows[e.RowIndex].Cells[0].Value.ToString();

            txtID.Text = valore;
        }

        private void txtID_TextChanged(object sender, EventArgs e)
        {
            try
            {
                conn.Open();

                OleDbCommand caricaValori = new OleDbCommand();
                caricaValori.Connection = conn;
                caricaValori.CommandText = "select Importo, Format(Data, \"dd/MM/yyyy\") as Data, Descrizione, Escludi from Spesa where ID = " + txtID.Text + "";
                OleDbDataReader readerValori = caricaValori.ExecuteReader();
                readerValori.Read();

                txtImporto.Text = readerValori["Importo"].ToString();
                txtData.Text = readerValori["Data"].ToString();
                txtDescrizione.Text = readerValori["Descrizione"].ToString();

                if (Convert.ToInt32(readerValori["Escludi"]) == 0)
                    chkEscludi.Checked = false;
                else
                    chkEscludi.Checked = true;


                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore nel caricamento dei valori \"Importo\" e \"Data\" relativi a questo ID: " + ex);
                conn.Close();
            }
        }
    }
}
