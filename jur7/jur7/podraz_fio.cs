using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace jur7
{
    public partial class podraz_fio : Form
    {
        string kod_pod = "";
        string podraz_naim = "";

        Connect sql = new Connect();
        public podraz_fio(string kod_pod,string podraz_naim)
        {
            InitializeComponent();

            this.kod_pod = kod_pod;
            this.podraz_naim = podraz_naim;

            sql.Connection();

            run_main();
        }

        public void run_main()
        {
            var query = " SELECT id,kod_pod,podraz_naim,fio FROM podraz where podraz_kod='" + kod_pod+"'  ";
            sql.myReader = sql.return_MySqlCommand(query).ExecuteReader();
            while (sql.myReader.Read())
            {
                //gruppa,kod_gruppa,naim,schet,prosent_izn,debet,subschet,kredit
                podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Add()].Cells[0].Value = (sql.myReader["id"] != DBNull.Value ? sql.myReader.GetString("id") : "");
                podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[1].Value = kod_pod;
                podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[2].Value = podraz_naim;
                podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[3].Value = (sql.myReader["fio"] != DBNull.Value ? sql.myReader.GetString("fio") : "");
            }
            sql.myReader.Close();
        }

        private void podraz_fio_Load(object sender, EventArgs e)
        {
            this.podrazdelenie_dataGridView.RowsDefaultCellStyle.BackColor = Color.White;
            this.podrazdelenie_dataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(233, 233, 234);

            podrazdelenie_dataGridView.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
        }

        private void podrazdelenie_dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //try
            //{
            if (podrazdelenie_dataGridView.CurrentRow != null)
            {
                DataGridViewRow dgvRow = podrazdelenie_dataGridView.CurrentRow;




                if (dgvRow.Cells[0].Value == null)
                {
                    Console.WriteLine("insert");

                    sql.return_MySqlCommand("insert into podraz (podraz_naim,podraz_kod,fio) values" +
                                        "('" + (podraz_naim) + "', " +
                                         "'" + (kod_pod) + "', " +
                                        "'" + (podrazdelenie_dataGridView.CurrentRow.Cells[3].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[3].Value : "") + "' " +
                                        ") ").ExecuteNonQuery();

                    this.podrazdelenie_dataGridView.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.podrazdelenie_dataGridView_CellValueChanged);
                    sql.myReader = sql.return_MySqlCommand("select max(id) as id from podraz").ExecuteReader();
                    while (sql.myReader.Read())
                    {
                        podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.CurrentRow.Index].Cells[0].Value = sql.myReader.GetString("id");
                    }
                    sql.myReader.Close();
                    this.podrazdelenie_dataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.podrazdelenie_dataGridView_CellValueChanged);
                }
                else
                {
                    Console.WriteLine("update " + dgvRow.Cells[0].Value);

                    sql.return_MySqlCommand("update podraz set " +
                     "podraz_kod = '" + (kod_pod) + "', " +
                     "podraz_naim = '" + (podraz_naim) + "', " +
                     "fio = '" + (podrazdelenie_dataGridView.CurrentRow.Cells[3].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[3].Value : "") + "' " +
                     " where id = '" + podrazdelenie_dataGridView.CurrentRow.Cells[0].Value + "' ").ExecuteNonQuery();
                }
            }


            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("gruppa_dataGridView_CellValueChanged " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void podrazdelenie_dataGridView_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            try
            {
                DialogResult dialogResult = MessageBox.Show("Вы действительно хотите удалить данные?", "Удаление", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    foreach (DataGridViewRow row in podrazdelenie_dataGridView.SelectedRows)
                    {
                        if (row.Cells[0].Value != null)
                        {

                            sql.return_MySqlCommand("delete from podraz where id = " + row.Cells[0].Value + "").ExecuteNonQuery();
                        }
                    }
                }
                else
                {

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("prixod_dataGridView_UserDeletingRow " + ex.Message);
            }
        }
    }
}
