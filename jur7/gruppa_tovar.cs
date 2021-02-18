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
    public partial class gruppa_tovar : Form
    {
        Connect sql = new Connect();
        Connect sql1 = new Connect();
        public gruppa_tovar()
        {
            InitializeComponent();

            sql.Connection();
            sql1.Connection();

            run_main();
        }

        public string refresh_strings_to_mysql(string mystring)
        {
            string str = string.Format("{0:#0.00}", Convert.ToDouble(mystring.Replace('.', ','))).Replace(',', '.');
            Console.WriteLine(str);
            return str;
        }
        public void run_main()
        {
            try
            {
                var query = " select * from gruppa_jur7 group by kod_gruppa order by cast(kod_gruppa as  unsigned) ";
                sql1.myReader = sql1.return_MySqlCommand(query).ExecuteReader();
                while (sql1.myReader.Read())
                {
                    podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Add()].Cells[0].Value = (sql1.myReader["id"] != DBNull.Value ? sql1.myReader.GetString("id") : "");
                    podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[1].Value = (sql1.myReader["gruppa"] != DBNull.Value ? sql1.myReader.GetString("gruppa") : "");
                    podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[2].Value = (sql1.myReader["kod_gruppa"] != DBNull.Value ? sql1.myReader.GetString("kod_gruppa") : "");
                    podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[3].Value = (sql1.myReader["naim"] != DBNull.Value ? sql1.myReader.GetString("naim") : "");
                    podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[4].Value = (sql1.myReader["schet"] != DBNull.Value ? sql1.myReader.GetString("schet") : "");
                    podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[5].Value = (sql1.myReader["prosent_izn"] != DBNull.Value ? (Convert.ToDouble(sql1.myReader.GetString("prosent_izn").Replace(".", ","))) : 0);
                    podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[6].Value = (sql1.myReader["debet"] != DBNull.Value ? sql1.myReader.GetString("debet") : "");
                    podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[7].Value = (sql1.myReader["subschet"] != DBNull.Value ? sql1.myReader.GetString("subschet") : "");
                    podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[8].Value = (sql1.myReader["kredit"] != DBNull.Value ? sql1.myReader.GetString("kredit") : "");
                }
                sql1.myReader.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show("run_main " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void gruppa_tovar_Load(object sender, EventArgs e)
        {
            try
            {
                this.podrazdelenie_dataGridView.RowsDefaultCellStyle.BackColor = Color.White;
                this.podrazdelenie_dataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(233, 233, 234);

                podrazdelenie_dataGridView.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                podrazdelenie_dataGridView.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                podrazdelenie_dataGridView.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                podrazdelenie_dataGridView.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                podrazdelenie_dataGridView.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                podrazdelenie_dataGridView.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                podrazdelenie_dataGridView.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                podrazdelenie_dataGridView.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                txt_gruppa.Items.Clear();
                sql.myReader = sql.return_MySqlCommand("SELECT distinct gruppa FROM gruppa0_jur7").ExecuteReader();

                while (sql.myReader.Read())
                {
                    txt_gruppa.Items.Add(sql.myReader.GetString("gruppa"));
                }
                sql.myReader.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show("gruppa_tovar_Load " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void podrazdelenie_dataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridViewRow row = podrazdelenie_dataGridView.CurrentRow;
                gruppa0 gruppa0 = new gruppa0();
                if (e.ColumnIndex == 3)
                {
                    if (gruppa0.ShowDialog() == DialogResult.OK)
                    {

                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("gruppa_dataGridView_CellDoubleClick " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void podrazdelenie_dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                //gruppa_dataGridView.Rows.Clear();



                if (podrazdelenie_dataGridView.CurrentRow != null)
                {
                    DataGridViewRow dgvRow = podrazdelenie_dataGridView.CurrentRow;



                    if (dgvRow.Cells[0].Value == null)
                    {
                        Console.WriteLine("insert");

                        sql.return_MySqlCommand("insert into gruppa_jur7 (gruppa,kod_gruppa,naim,schet,prosent_izn,debet,subschet,kredit) values" +
                                            "('" + (podrazdelenie_dataGridView.CurrentRow.Cells[1].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[1].Value : "") + "', " +
                                            "'" + (podrazdelenie_dataGridView.CurrentRow.Cells[2].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[2].Value : "") + "', " +
                                            "'" + (podrazdelenie_dataGridView.CurrentRow.Cells[3].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[3].Value : "") + "', " +
                                            "'" + (podrazdelenie_dataGridView.CurrentRow.Cells[4].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[4].Value : "") + "', " +
                                            "'" + (podrazdelenie_dataGridView.CurrentRow.Cells[5].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                            "'" + (podrazdelenie_dataGridView.CurrentRow.Cells[6].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[6].Value : "") + "', " +
                                            "'" + (podrazdelenie_dataGridView.CurrentRow.Cells[7].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[7].Value : "") + "', " +
                                            "'" + (podrazdelenie_dataGridView.CurrentRow.Cells[8].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[8].Value : "") + "' " +
                                            ") ").ExecuteNonQuery();

                        this.podrazdelenie_dataGridView.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.podrazdelenie_dataGridView_CellValueChanged);
                        sql1.myReader = sql1.return_MySqlCommand("select max(id) as id from gruppa_jur7").ExecuteReader();
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

                        sql.return_MySqlCommand("update gruppa_jur7 set " +
                         "gruppa = '" + (podrazdelenie_dataGridView.CurrentRow.Cells[1].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[1].Value : "") + "', " +
                         "kod_gruppa = '" + (podrazdelenie_dataGridView.CurrentRow.Cells[2].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[2].Value : "") + "', " +
                         "naim = '" + (podrazdelenie_dataGridView.CurrentRow.Cells[3].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[3].Value : "") + "', " +
                         "schet = '" + (podrazdelenie_dataGridView.CurrentRow.Cells[4].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[4].Value : "") + "', " +
                         "prosent_izn = '" + (podrazdelenie_dataGridView.CurrentRow.Cells[5].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[5].Value.ToString().Replace(",", ".") : "0") + "', " +
                         "debet = '" + (podrazdelenie_dataGridView.CurrentRow.Cells[6].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[6].Value : "") + "', " +
                         "subschet = '" + (podrazdelenie_dataGridView.CurrentRow.Cells[7].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[7].Value : "") + "', " +
                         "kredit = '" + (podrazdelenie_dataGridView.CurrentRow.Cells[8].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[8].Value : "") + "' " +
                         " where id = '" + podrazdelenie_dataGridView.CurrentRow.Cells[0].Value + "' ").ExecuteNonQuery();

                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("gruppa_dataGridView_CellValueChanged " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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

                            sql.return_MySqlCommand("delete from gruppa_jur7 where id = " + row.Cells[0].Value + "").ExecuteNonQuery();
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

        private void podrazdelenie_dataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.Exception.Message == "DataGridViewComboBoxCell value is not valid.")
            {
                object value = podrazdelenie_dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (!((DataGridViewComboBoxColumn)podrazdelenie_dataGridView.Columns[e.ColumnIndex]).Items.Contains(value))
                {
                    ((DataGridViewComboBoxColumn)podrazdelenie_dataGridView.Columns[e.ColumnIndex]).Items.Add(value);
                    e.ThrowException = false;
                }
            }
        }
    }
}
