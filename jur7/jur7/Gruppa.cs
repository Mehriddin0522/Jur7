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
    public partial class Gruppa : Form
    {
        Connect sql = new Connect();
        Connect sql1 = new Connect();
        Connect sql2 = new Connect();

        //public string id = "";
        public string kod_gruppa = "";
        public string naim = "";
        public string schet = "";
        public string debet = "";
        public string kredit = "";

        public Gruppa()
        {
            InitializeComponent();

            sql.Connection();
            sql1.Connection();
            sql2.Connection();

            
        }

        public string refresh_strings_to_mysql(string mystring)
        {
            string str = string.Format("{0:#0.00}", Convert.ToDouble(mystring.Replace('.', ','))).Replace(',', '.');
            Console.WriteLine(str);
            return str;
        }
        public void run_main()
        {
            var query = "select * from gruppa";
            sql2.myReader = sql2.return_MySqlCommand(query).ExecuteReader();
            while (sql2.myReader.Read())
            {
                //gruppa,kod_gruppa,naim,schet,prosent_izn,debet,subschet,kredit
                gruppa__dataGridView.Rows[gruppa__dataGridView.Rows.Add()].Cells[0].Value = (sql2.myReader["id"] != DBNull.Value ? sql2.myReader.GetString("id") : "");
                gruppa__dataGridView.Rows[gruppa__dataGridView.Rows.Count - 2].Cells[1].Value = (sql2.myReader["gruppa"] != DBNull.Value ? sql2.myReader.GetString("gruppa") : "");
                gruppa__dataGridView.Rows[gruppa__dataGridView.Rows.Count - 2].Cells[2].Value = (sql2.myReader["kod_gruppa"] != DBNull.Value ? sql2.myReader.GetString("kod_gruppa") : "");
                gruppa__dataGridView.Rows[gruppa__dataGridView.Rows.Count - 2].Cells[3].Value = (sql2.myReader["naim"] != DBNull.Value ? sql2.myReader.GetString("naim") : "");
                gruppa__dataGridView.Rows[gruppa__dataGridView.Rows.Count - 2].Cells[4].Value = (sql2.myReader["schet"] != DBNull.Value ? sql2.myReader.GetString("schet") : "");
                gruppa__dataGridView.Rows[gruppa__dataGridView.Rows.Count - 2].Cells[5].Value = (sql2.myReader["prosent_izn"] != DBNull.Value ? (Convert.ToDouble(sql2.myReader.GetString("prosent_izn").Replace(".", ","))) : 0);
                gruppa__dataGridView.Rows[gruppa__dataGridView.Rows.Count - 2].Cells[6].Value = (sql2.myReader["debet"] != DBNull.Value ? sql2.myReader.GetString("debet") : "");
                gruppa__dataGridView.Rows[gruppa__dataGridView.Rows.Count - 2].Cells[7].Value = (sql2.myReader["subschet"] != DBNull.Value ? sql2.myReader.GetString("subschet") : "");
                gruppa__dataGridView.Rows[gruppa__dataGridView.Rows.Count - 2].Cells[8].Value = (sql2.myReader["kredit"] != DBNull.Value ? sql2.myReader.GetString("kredit") : "");
            }
            sql2.myReader.Close();
        }
        private void Gruppa_Load(object sender, EventArgs e)
        {
            this.gruppa__dataGridView.RowsDefaultCellStyle.BackColor = Color.White;
            this.gruppa__dataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(233, 233, 234);

            gruppa__dataGridView.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            gruppa__dataGridView.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            gruppa__dataGridView.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            gruppa__dataGridView.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            gruppa__dataGridView.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            gruppa__dataGridView.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            gruppa__dataGridView.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            gruppa__dataGridView.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            txt_gruppa.Items.Clear();
            sql.myReader = sql.return_MySqlCommand("SELECT distinct gruppa FROM gruppa0").ExecuteReader();

            while (sql.myReader.Read())
            {
                txt_gruppa.Items.Add(sql.myReader.GetString("gruppa"));
            }
            sql.myReader.Close();

            run_main();

        }

        private void gruppa_dataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridViewRow row = gruppa__dataGridView.CurrentRow;
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

        private void gruppa__dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //try
            //{
            //gruppa_dataGridView.Rows.Clear();



            //if (gruppa__dataGridView.CurrentRow != null)
            //{
            //    DataGridViewRow dgvRow = gruppa__dataGridView.CurrentRow;



            //    if (dgvRow.Cells[0].Value == null)
            //    {
            //        Console.WriteLine("insert");

            //        sql.return_MySqlCommand("insert into gruppa (gruppa,kod_gruppa,naim,schet,prosent_izn,debet,subschet,kredit) values" +
            //                            "('" + (gruppa__dataGridView.CurrentRow.Cells[1].Value != null ? gruppa__dataGridView.CurrentRow.Cells[1].Value : "") + "', " +
            //                            "'" + (gruppa__dataGridView.CurrentRow.Cells[2].Value != null ? gruppa__dataGridView.CurrentRow.Cells[2].Value : "") + "', " +
            //                            "'" + (gruppa__dataGridView.CurrentRow.Cells[3].Value != null ? gruppa__dataGridView.CurrentRow.Cells[3].Value : "") + "', " +
            //                            "'" + (gruppa__dataGridView.CurrentRow.Cells[4].Value != null ? gruppa__dataGridView.CurrentRow.Cells[4].Value : "") + "', " +
            //                            "'" + (gruppa__dataGridView.CurrentRow.Cells[5].Value != null ? gruppa__dataGridView.CurrentRow.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
            //                            "'" + (gruppa__dataGridView.CurrentRow.Cells[6].Value != null ? gruppa__dataGridView.CurrentRow.Cells[6].Value : "") + "', " +
            //                            "'" + (gruppa__dataGridView.CurrentRow.Cells[7].Value != null ? gruppa__dataGridView.CurrentRow.Cells[7].Value : "") + "', " +
            //                            "'" + (gruppa__dataGridView.CurrentRow.Cells[8].Value != null ? gruppa__dataGridView.CurrentRow.Cells[8].Value : "") + "' " +
            //                            ") ").ExecuteNonQuery();

            //        this.gruppa__dataGridView.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.gruppa__dataGridView_CellValueChanged);
            //        sql1.myReader = sql1.return_MySqlCommand("select max(id) as id from gruppa").ExecuteReader();
            //        while (sql1.myReader.Read())
            //        {
            //            gruppa__dataGridView.Rows[gruppa__dataGridView.CurrentRow.Index].Cells[0].Value = sql1.myReader.GetString("id");
            //        }
            //        sql1.myReader.Close();
            //        this.gruppa__dataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.gruppa__dataGridView_CellValueChanged);
            //    }
            //    else
            //    {
            //        Console.WriteLine("update " + dgvRow.Cells[0].Value);

            //        sql.return_MySqlCommand("update gruppa set " +
            //         "gruppa = '" + (gruppa__dataGridView.CurrentRow.Cells[1].Value != null ? gruppa__dataGridView.CurrentRow.Cells[1].Value : "") + "', " +
            //         "kod_gruppa = '" + (gruppa__dataGridView.CurrentRow.Cells[2].Value != null ? gruppa__dataGridView.CurrentRow.Cells[2].Value : "") + "', " +
            //         "naim = '" + (gruppa__dataGridView.CurrentRow.Cells[3].Value != null ? gruppa__dataGridView.CurrentRow.Cells[3].Value : "") + "', " +
            //         "schet = '" + (gruppa__dataGridView.CurrentRow.Cells[4].Value != null ? gruppa__dataGridView.CurrentRow.Cells[4].Value : "") + "', " +
            //         "prosent_izn = '" + (gruppa__dataGridView.CurrentRow.Cells[5].Value != null ? gruppa__dataGridView.CurrentRow.Cells[5].Value.ToString().Replace(",", ".") : "0") + "', " +
            //         "debet = '" + (gruppa__dataGridView.CurrentRow.Cells[6].Value != null ? gruppa__dataGridView.CurrentRow.Cells[6].Value : "") + "', " +
            //         "subschet = '" + (gruppa__dataGridView.CurrentRow.Cells[7].Value != null ? gruppa__dataGridView.CurrentRow.Cells[7].Value : "") + "', " +
            //         "kredit = '" + (gruppa__dataGridView.CurrentRow.Cells[8].Value != null ? gruppa__dataGridView.CurrentRow.Cells[8].Value : "") + "' " +
            //         " where id = '" + gruppa__dataGridView.CurrentRow.Cells[0].Value + "' ").ExecuteNonQuery();

            //    }
            //}


            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("gruppa_dataGridView_CellValueChanged " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}

        }
    }
}
