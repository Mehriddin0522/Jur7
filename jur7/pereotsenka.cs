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
    public partial class pereotsenka : Form
    {
        Connect sql = new Connect();
        Connect sql2 = new Connect();
        Connect sql3 = new Connect();

        public string string_for_otdels;
        public string month_global;
        public string year_global;
        public pereotsenka(string string_for_otdels, string year_global, string month_global)
        {
            InitializeComponent();

            sql.Connection();
            sql2.Connection();
            sql3.Connection();

            this.string_for_otdels = string_for_otdels;
            this.month_global = month_global;
            this.year_global = year_global;

            komu_combo();
        }


        public void komu_combo()
        {
            komu_pri_comboBox.Items.Clear();
            sql.myReader = sql.return_MySqlCommand("SELECT distinct podraz_naim FROM podraz_jur7").ExecuteReader();

            while (sql.myReader.Read())
            {
                komu_pri_comboBox.Items.Add(sql.myReader.GetString("podraz_naim"));
            }
            sql.myReader.Close();

            komu_pri_comboBox2.Items.Clear();
            sql.myReader = sql.return_MySqlCommand("SELECT distinct podraz_naim FROM podraz_jur7").ExecuteReader();

            while (sql.myReader.Read())
            {
                komu_pri_comboBox2.Items.Add(sql.myReader.GetString("podraz_naim"));
            }
            sql.myReader.Close();
        }

        public string refresh_strings_to_mysql(string mystring)
        {
            string str = string.Format("{0:#0.00}", Convert.ToDouble(mystring.Replace('.', ','))).Replace(',', '.');
            Console.WriteLine(str);
            return str;
        }

        public string refresh_string_currency(string test_string)
        {
            string str = "";
            try
            {
                str = string.Format("{0:#,0.00}", (object)Convert.ToDouble(test_string.ToString().Replace('.', ','))); //"{0:#,0}"
            }
            catch (Exception ex)
            {
                Console.WriteLine("   ------------- refresh_string_currency :" + ex.Message);
            }
            return str;
        }

        String getmonth_String2;
        public string set_month_name2(int getmonth)
        {
            switch (getmonth)
            {
                case 1:
                    {
                        getmonth_String2 = "январь";
                        break;
                    }
                case 2:
                    {
                        getmonth_String2 = "февраль";
                        break;
                    }
                case 3:
                    {
                        getmonth_String2 = "март";
                        break;
                    }
                case 4:
                    {
                        getmonth_String2 = "апрель";
                        break;
                    }
                case 5:
                    {
                        getmonth_String2 = "май";
                        break;
                    }
                case 6:
                    {
                        getmonth_String2 = "июнь";
                        break;
                    }
                case 7:
                    {
                        getmonth_String2 = "июль";
                        break;
                    }
                case 8:
                    {
                        getmonth_String2 = "августь";
                        break;
                    }
                case 9:
                    {
                        getmonth_String2 = "сентябрь";
                        break;
                    }
                case 10:
                    {
                        getmonth_String2 = "октябрь";
                        break;
                    }
                case 11:
                    {
                        getmonth_String2 = "ноябрь";
                        break;
                    }
                case 12:
                    {
                        getmonth_String2 = "декабрь";
                        break;
                    }
            }
            return getmonth_String2;
        }

        private void komu_pri_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                komu_pri_comboBox2.Text = "";
                komu_pri_comboBox2.Items.Clear();
                var select = "SELECT * FROM podraz_jur7 where podraz_naim='" + komu_pri_comboBox.Text + "'";
                sql3.myReader = sql3.return_MySqlCommand(select).ExecuteReader();
                while (sql3.myReader.Read())
                {
                    komu_pri_comboBox2.Items.Add(sql3.myReader["fio"] != DBNull.Value ? sql3.myReader.GetString("fio") : "");
                }
                sql3.myReader.Close();

            }
            catch (Exception ex)
            {

                MessageBox.Show("komu_vnut_per_comboBox_SelectedIndexChanged " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }



        public void label_update_prixod()
        {
            double summa = 0;
            double iznos = 0;


            foreach (DataGridViewRow row in pereosenka_dataGridView.Rows)
            {
                summa = summa + (row.Cells[9].Value != null ? Double.Parse(row.Cells[9].Value.ToString()) : 0);

                iznos = iznos + (row.Cells[16].Value != null ? Double.Parse(row.Cells[16].Value.ToString()) : 0);

            }
            if (summa.ToString().Length <= 3)
            {
                prixod_obshiy_summa_label.Text = string.Format("{0:#0.00}", summa);
            }
            if (summa.ToString().Length > 3)
            {
                prixod_obshiy_summa_label.Text = string.Format("{0:#0,000.00}", summa);
            }

            if (iznos.ToString().Length <= 3)
            {
                iznos_sum_lbl.Text = string.Format("{0:#0.00}", iznos);
            }
            if (iznos.ToString().Length > 3)
            {
                iznos_sum_lbl.Text = string.Format("{0:#0,000.00}", iznos);
            }

        }

        private void pereosenka_dataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void pereosenka_dataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                string ot_kogo_1 = komu_pri_comboBox.Text;
                string ot_kogo_2 = komu_pri_comboBox2.Text;

                DataGridViewRow row_ost = pereosenka_dataGridView.CurrentRow;

                //if (ot_kogo_ras_ComboBox.Text != "" && ot_kogo_ras_comboBox2.Text != "")
                //{
                if (e.ColumnIndex == 3)
                {

                    string pereosenka_visible = "1";
                    Sklad ostatok = new Sklad(string_for_otdels, year_global, month_global, ot_kogo_1, ot_kogo_2, pereosenka_visible);
                    ostatok.WindowState = FormWindowState.Maximized;

                    //public string product_id = "";
                    //public string schet = "";
                    //public string naim = "";
                    //public string edin = "";
                    //public string gruppa = "";
                    //public string seria_num = "";
                    //public string inv_num = "";

                    if (ostatok.ShowDialog() == DialogResult.OK)
                    {

                    }
                }



                label_update_prixod();

            }
            catch (Exception ex)
            {
                MessageBox.Show("rasxod_dataGridView_CellDoubleClick " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pereosenka_dataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.Exception.Message == "DataGridViewComboBoxCell value is not valid.")
            {
                object value = pereosenka_dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (!((DataGridViewComboBoxColumn)pereosenka_dataGridView.Columns[e.ColumnIndex]).Items.Contains(value))
                {
                    ((DataGridViewComboBoxColumn)pereosenka_dataGridView.Columns[e.ColumnIndex]).Items.Add(value);
                    e.ThrowException = false;
                }
            }
        }

        private void pereosenka_dataGridView_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            try
            {
                DialogResult dialogResult = MessageBox.Show("Вы действительно хотите удалить данные?", "Удаление", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    foreach (DataGridViewRow row in pereosenka_dataGridView.SelectedRows)
                    {
                        if (row.Cells[0].Value != null)
                        {

                            //sql.return_MySqlCommand("delete from gruppa where id = " + row.Cells[0].Value + "").ExecuteNonQuery();
                        }
                    }
                }
                else
                {

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("pereosenka_dataGridView_UserDeletingRow " + ex.Message);
            }
        }
    }
}
