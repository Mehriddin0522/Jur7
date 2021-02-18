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
    public partial class add_period : Form
    {
        Connect sql = new Connect();
        Connect sql1 = new Connect();
        public add_period()
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
                var query = " SELECT * FROM pereosenka_data_jur7 ";
                sql1.myReader = sql1.return_MySqlCommand(query).ExecuteReader();
                while (sql1.myReader.Read())
                {
                    //gruppa,kod_gruppa,naim,schet,prosent_izn,debet,subschet,kredit
                    podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Add()].Cells[0].Value = (sql1.myReader["id"] != DBNull.Value ? sql1.myReader.GetString("id") : "");
                    podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[1].Value = (sql1.myReader["number"] != DBNull.Value ? sql1.myReader.GetString("number") : "");
                    podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[2].Value = (sql1.myReader["date_start"] != DBNull.Value ? (DateTime.Parse(sql1.myReader.GetString("date_start")).ToString("dd.MM.yyyy")) : null);
                    podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[3].Value = (sql1.myReader["date_finish"] != DBNull.Value ? (DateTime.Parse(sql1.myReader.GetString("date_finish")).ToString("dd.MM.yyyy")) : null);

                }
                sql1.myReader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("run_main " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void add_period_Load(object sender, EventArgs e)
        {
            this.podrazdelenie_dataGridView.RowsDefaultCellStyle.BackColor = Color.White;
            this.podrazdelenie_dataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(233, 233, 234);

            podrazdelenie_dataGridView.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            podrazdelenie_dataGridView.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            podrazdelenie_dataGridView.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            txt_num.Items.Clear();
            txt_num.Items.Add("1");
            txt_num.Items.Add("2");
            txt_num.Items.Add("3");
            txt_num.Items.Add("4");
            txt_num.Items.Add("5");
            txt_num.Items.Add("6");
            txt_num.Items.Add("7");
            txt_num.Items.Add("8");


        }

        private void podrazdelenie_dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (podrazdelenie_dataGridView.CurrentRow != null)
                {
                    DataGridViewRow dgvRow = podrazdelenie_dataGridView.CurrentRow;

                    ////"" + (prixod_dataGridView.Rows[i].Cells[17].Value != null ? "'" + DateTime.Parse(prixod_dataGridView.Rows[i].Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +


                    if (dgvRow.Cells[0].Value == null)
                    {
                        Console.WriteLine("insert");
                        var ins = "insert into pereosenka_data_jur7 (number,date_start,date_finish) values" +
                                            "('" + (podrazdelenie_dataGridView.CurrentRow.Cells[1].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[1].Value : "") + "', " +
                                            "" + (podrazdelenie_dataGridView.CurrentRow.Cells[2].Value != null ? "'" + DateTime.Parse(podrazdelenie_dataGridView.CurrentRow.Cells[2].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                                            "" + (podrazdelenie_dataGridView.CurrentRow.Cells[3].Value != null ? "'" + DateTime.Parse(podrazdelenie_dataGridView.CurrentRow.Cells[3].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + " " +
                                            ") ";
                        sql.return_MySqlCommand(ins).ExecuteNonQuery();

                        this.podrazdelenie_dataGridView.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.podrazdelenie_dataGridView_CellValueChanged);
                        sql1.myReader = sql1.return_MySqlCommand("select max(id) as id from pereosenka_data_jur7").ExecuteReader();
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

                        var update = "update pereosenka_data_jur7 set " +
                         "number = '" + (podrazdelenie_dataGridView.CurrentRow.Cells[1].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[1].Value : "") + "', " +
                         "date_start = " + (podrazdelenie_dataGridView.CurrentRow.Cells[2].Value != null ? "'" + DateTime.Parse(podrazdelenie_dataGridView.CurrentRow.Cells[2].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                         "date_finish = " + (podrazdelenie_dataGridView.CurrentRow.Cells[3].Value != null ? "'" + DateTime.Parse(podrazdelenie_dataGridView.CurrentRow.Cells[3].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + " " +
                         " where id = '" + podrazdelenie_dataGridView.CurrentRow.Cells[0].Value + "' ";
                        sql.return_MySqlCommand(update).ExecuteNonQuery();
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("podrazdelenie_dataGridView_CellValueChanged " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                            sql.return_MySqlCommand("delete from pereosenka_data_jur7 where id = " + row.Cells[0].Value + "").ExecuteNonQuery();
                        }
                    }
                }
                else
                {

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("podrazdelenie_dataGridView_UserDeletingRow " + ex.Message);
            }
        }
    }
}
