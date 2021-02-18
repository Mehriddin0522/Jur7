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
    public partial class Podrazdelenie : Form
    {

        Connect sql = new Connect();
        Connect sql1 = new Connect();
        public Podrazdelenie()
        {
            InitializeComponent();

            sql.Connection();
            sql1.Connection();
            

            run_main();
        }

        
        public void run_main()
        {
            var query = "SELECT * FROM podrazdeleniya where podraz is not null";
            sql1.myReader = sql1.return_MySqlCommand(query).ExecuteReader();
            while (sql1.myReader.Read())
            {
                //gruppa,kod_gruppa,naim,schet,prosent_izn,debet,subschet,kredit
                podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Add()].Cells[0].Value = (sql1.myReader["id"] != DBNull.Value ? sql1.myReader.GetString("id") : "");
                podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[1].Value = (sql1.myReader["type"] != DBNull.Value ? sql1.myReader.GetString("type") : "");
                podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[2].Value = (sql1.myReader["podraz"] != DBNull.Value ? sql1.myReader.GetString("podraz") : "");
                podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[3].Value = (sql1.myReader["podraz_fio"] != DBNull.Value ? sql1.myReader.GetString("podraz_fio") : "");
            }
            sql1.myReader.Close();
        }

        private void Komu_Load(object sender, EventArgs e)
        {
            this.podrazdelenie_dataGridView.RowsDefaultCellStyle.BackColor = Color.White;
            this.podrazdelenie_dataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(233, 233, 234);

            podrazdelenie_dataGridView.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

            txt_type.Items.Add("От кого");
            txt_type.Items.Add("Кому");
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

                    sql.return_MySqlCommand("insert into podrazdeleniya (type,podraz,podraz_fio) values" +
                                        "('" + (podrazdelenie_dataGridView.CurrentRow.Cells[1].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[1].Value : "") + "', " +
                                        "'" + (podrazdelenie_dataGridView.CurrentRow.Cells[2].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[2].Value : "") + "', " +
                                        "'" + (podrazdelenie_dataGridView.CurrentRow.Cells[3].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[3].Value : "") + "' " +
                                        ") ").ExecuteNonQuery();

                    this.podrazdelenie_dataGridView.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.podrazdelenie_dataGridView_CellValueChanged);
                    sql1.myReader = sql1.return_MySqlCommand("select max(id) as id from podrazdeleniya").ExecuteReader();
                    while (sql1.myReader.Read())
                    {
                        podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.CurrentRow.Index].Cells[0].Value = sql1.myReader.GetString("id");
                    }
                    sql1.myReader.Close();
                    this.podrazdelenie_dataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.podrazdelenie_dataGridView_CellValueChanged);
                }
                else
                {
                    Console.WriteLine("update " + dgvRow.Cells[0].Value);

                    sql.return_MySqlCommand("update podrazdeleniya set " +
                     "type = '" + (podrazdelenie_dataGridView.CurrentRow.Cells[1].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[1].Value : "") + "', " +
                     "podraz = '" + (podrazdelenie_dataGridView.CurrentRow.Cells[2].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[2].Value : "") + "', " +
                     "podraz_fio = '" + (podrazdelenie_dataGridView.CurrentRow.Cells[3].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[3].Value : "") + "' " +
                     " where id = '" + podrazdelenie_dataGridView.CurrentRow.Cells[0].Value + "' ").ExecuteNonQuery();
                }
            }


            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("gruppa_dataGridView_CellValueChanged " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

       public string id = "";
       public string podrazdelenie = "";
        private void podrazdelenie_dataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void podrazdelenie_dataGridView_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            
           
        }
    }
}
