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
    public partial class gruppa0 : Form
    {
        Connect sql = new Connect();
        Connect sql1 = new Connect();
        public gruppa0()
        {
            InitializeComponent();

            sql.Connection();
            sql1.Connection();

            run_main();
        }

        public void run_main()
        {
            try
            {

                var query = " SELECT * FROM gruppa0_jur7  ";
                sql1.myReader = sql1.return_MySqlCommand(query).ExecuteReader();
                while (sql1.myReader.Read())
                {
                    //gruppa,first,two,three,four,five,six,seven,eight
                    gruppa0_dataGridView.Rows[gruppa0_dataGridView.Rows.Add()].Cells[0].Value = (sql1.myReader["id"] != DBNull.Value ? sql1.myReader.GetString("id") : "");
                    gruppa0_dataGridView.Rows[gruppa0_dataGridView.Rows.Count - 2].Cells[1].Value = (sql1.myReader["gruppa"] != DBNull.Value ? sql1.myReader.GetString("gruppa") : "");
                    gruppa0_dataGridView.Rows[gruppa0_dataGridView.Rows.Count - 2].Cells[2].Value = (sql1.myReader["first"] != DBNull.Value ? (Convert.ToDouble(sql1.myReader.GetString("first").Replace(".", ","))) : 0);
                    gruppa0_dataGridView.Rows[gruppa0_dataGridView.Rows.Count - 2].Cells[3].Value = (sql1.myReader["two"] != DBNull.Value ? (Convert.ToDouble(sql1.myReader.GetString("two").Replace(".", ","))) : 0);
                    gruppa0_dataGridView.Rows[gruppa0_dataGridView.Rows.Count - 2].Cells[4].Value = (sql1.myReader["three"] != DBNull.Value ? (Convert.ToDouble(sql1.myReader.GetString("three").Replace(".", ","))) : 0);
                    gruppa0_dataGridView.Rows[gruppa0_dataGridView.Rows.Count - 2].Cells[5].Value = (sql1.myReader["four"] != DBNull.Value ? (Convert.ToDouble(sql1.myReader.GetString("four").Replace(".", ","))) : 0);
                    gruppa0_dataGridView.Rows[gruppa0_dataGridView.Rows.Count - 2].Cells[6].Value = (sql1.myReader["five"] != DBNull.Value ? (Convert.ToDouble(sql1.myReader.GetString("five").Replace(".", ","))) : 0);
                    gruppa0_dataGridView.Rows[gruppa0_dataGridView.Rows.Count - 2].Cells[7].Value = (sql1.myReader["six"] != DBNull.Value ? (Convert.ToDouble(sql1.myReader.GetString("six").Replace(".", ","))) : 0);
                    gruppa0_dataGridView.Rows[gruppa0_dataGridView.Rows.Count - 2].Cells[8].Value = (sql1.myReader["seven"] != DBNull.Value ? (Convert.ToDouble(sql1.myReader.GetString("seven").Replace(".", ","))) : 0);
                    gruppa0_dataGridView.Rows[gruppa0_dataGridView.Rows.Count - 2].Cells[9].Value = (sql1.myReader["eight"] != DBNull.Value ? (Convert.ToDouble(sql1.myReader.GetString("eight").Replace(".", ","))) : 0);
                }
                sql1.myReader.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show("run_main " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void gruppa0_Load(object sender, EventArgs e)
        {
            this.gruppa0_dataGridView.RowsDefaultCellStyle.BackColor = Color.White;
            this.gruppa0_dataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(233, 233, 234);

            gruppa0_dataGridView.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            gruppa0_dataGridView.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            gruppa0_dataGridView.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            gruppa0_dataGridView.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            gruppa0_dataGridView.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            gruppa0_dataGridView.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            gruppa0_dataGridView.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            gruppa0_dataGridView.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            gruppa0_dataGridView.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

        }

        private void gruppa0_dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (gruppa0_dataGridView.CurrentRow != null)
                {
                    DataGridViewRow dgvRow = gruppa0_dataGridView.CurrentRow;




                    if (dgvRow.Cells[0].Value == null)
                    {
                        Console.WriteLine("insert");

                        sql.return_MySqlCommand("insert into gruppa0_jur7 (gruppa,first,two,three,four,five,six,seven,eight) values" +
                                            "('" + (gruppa0_dataGridView.CurrentRow.Cells[1].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[1].Value : "") + "', " +
                                             "'" + (gruppa0_dataGridView.CurrentRow.Cells[2].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[2].Value.ToString().Replace(',', '.') : "0") + "'," +
                                             "'" + (gruppa0_dataGridView.CurrentRow.Cells[3].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[3].Value.ToString().Replace(',', '.') : "0") + "'," +
                                             "'" + (gruppa0_dataGridView.CurrentRow.Cells[4].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[4].Value.ToString().Replace(',', '.') : "0") + "'," +
                                             "'" + (gruppa0_dataGridView.CurrentRow.Cells[5].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                             "'" + (gruppa0_dataGridView.CurrentRow.Cells[6].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[6].Value.ToString().Replace(',', '.') : "0") + "'," +
                                             "'" + (gruppa0_dataGridView.CurrentRow.Cells[7].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                             "'" + (gruppa0_dataGridView.CurrentRow.Cells[8].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[8].Value.ToString().Replace(',', '.') : "0") + "'," +
                                             "'" + (gruppa0_dataGridView.CurrentRow.Cells[9].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[9].Value.ToString().Replace(',', '.') : "0") + "'" +
                                            ") ").ExecuteNonQuery();

                        this.gruppa0_dataGridView.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.gruppa0_dataGridView_CellValueChanged);
                        sql1.myReader = sql1.return_MySqlCommand("select max(id) as id from gruppa0_jur7").ExecuteReader();
                        while (sql1.myReader.Read())
                        {
                            gruppa0_dataGridView.Rows[gruppa0_dataGridView.CurrentRow.Index].Cells[0].Value = sql1.myReader.GetString("id");
                        }
                        sql1.myReader.Close();
                        this.gruppa0_dataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.gruppa0_dataGridView_CellValueChanged);
                    }
                    else
                    {
                        Console.WriteLine("update " + dgvRow.Cells[0].Value);

                        sql.return_MySqlCommand("update gruppa0_jur7 set " +
                         "gruppa = '" + (gruppa0_dataGridView.CurrentRow.Cells[1].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[1].Value : "") + "', " +
                         "first = '" + (gruppa0_dataGridView.CurrentRow.Cells[2].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[2].Value.ToString().Replace(",", ".") : "0") + "', " +
                          "two = '" + (gruppa0_dataGridView.CurrentRow.Cells[3].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[3].Value.ToString().Replace(",", ".") : "0") + "', " +
                           "three = '" + (gruppa0_dataGridView.CurrentRow.Cells[4].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[4].Value.ToString().Replace(",", ".") : "0") + "', " +
                            "four = '" + (gruppa0_dataGridView.CurrentRow.Cells[5].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[5].Value.ToString().Replace(",", ".") : "0") + "', " +
                             "five = '" + (gruppa0_dataGridView.CurrentRow.Cells[6].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[6].Value.ToString().Replace(",", ".") : "0") + "', " +
                              "six = '" + (gruppa0_dataGridView.CurrentRow.Cells[7].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[7].Value.ToString().Replace(",", ".") : "0") + "', " +
                               "seven = '" + (gruppa0_dataGridView.CurrentRow.Cells[8].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[8].Value.ToString().Replace(",", ".") : "0") + "', " +
                                "eight = '" + (gruppa0_dataGridView.CurrentRow.Cells[9].Value != null ? gruppa0_dataGridView.CurrentRow.Cells[9].Value.ToString().Replace(",", ".") : "0") + "' " +
                         " where id = '" + gruppa0_dataGridView.CurrentRow.Cells[0].Value + "' ").ExecuteNonQuery();
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("gruppa_dataGridView_CellValueChanged " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void label2_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void gruppa0_dataGridView_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            try
            {
                DialogResult dialogResult = MessageBox.Show("Вы действительно хотите удалить данные?", "Удаление", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    foreach (DataGridViewRow row in gruppa0_dataGridView.SelectedRows)
                    {
                        if (row.Cells[0].Value != null)
                        {

                            sql.return_MySqlCommand("delete from gruppa0_jur7 where id = " + row.Cells[0].Value + "").ExecuteNonQuery();
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

        private void label4_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try
            {

                add_period gruppa0 = new add_period();
                gruppa0.Show();

            }
            catch (Exception ex)
            {
                MessageBox.Show("gruppa_dataGridView_CellDoubleClick " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void gruppa0_dataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.Exception.Message == "DataGridViewComboBoxCell value is not valid.")
            {
                object value = gruppa0_dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (!((DataGridViewComboBoxColumn)gruppa0_dataGridView.Columns[e.ColumnIndex]).Items.Contains(value))
                {
                    ((DataGridViewComboBoxColumn)gruppa0_dataGridView.Columns[e.ColumnIndex]).Items.Add(value);
                    e.ThrowException = false;
                }
            }
        }
    }
}
