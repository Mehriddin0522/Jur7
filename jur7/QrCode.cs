using DataGridViewMultiColumnComboColumnDemo;
using MySql.Data.MySqlClient;
using QRCoder;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace jur7
{
    public partial class QrCode : Form
    {

        Number_To_Words_russian number_russian = new Number_To_Words_russian();

        Connect sql = new Connect();
        Connect sql1 = new Connect();
        Connect sql2 = new Connect();

        public string string_for_otdels;
        public string month_global;
        public string year_global;

        public QrCode(string string_for_otdels, string year_global, string month_global)
        {
            InitializeComponent();

            sql.Connection();
            sql1.Connection();
            sql2.Connection();

            this.string_for_otdels = string_for_otdels;
            this.year_global = year_global;
            this.month_global = month_global;

            add_items();
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

        DataTable multi_col = new DataTable();
        private void QrCode_Load(object sender, EventArgs e)
        {
            this.qrcode_dataGridView.RowsDefaultCellStyle.BackColor = Color.White;
            this.qrcode_dataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(233, 233, 234);

            qrcode_dataGridView.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            qrcode_dataGridView.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            qrcode_dataGridView.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            qrcode_dataGridView.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            qrcode_dataGridView.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            qrcode_dataGridView.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            qrcode_dataGridView.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            qrcode_dataGridView.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            qrcode_dataGridView.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            qrcode_dataGridView.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            qrcode_dataGridView.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            qrcode_dataGridView.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            qrcode_dataGridView.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            qrcode_dataGridView.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            qrcode_dataGridView.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            qrcode_dataGridView.Columns[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewMultiColumnComboColumn newColumn = (DataGridViewMultiColumnComboColumn)qrcode_dataGridView.Columns[2];

            sql.mydataAdapter = new MySqlDataAdapter();
            multi_col.Clear();
            sql.mydataAdapter.SelectCommand = this.sql.return_MySqlCommand(" SELECT kod_gruppa,schet,naim FROM gruppa_jur7 order by naim asc");
            sql.mydataAdapter.Fill(multi_col);

            newColumn.DataSource = multi_col;

            newColumn.DropDownWidth = 600;
            newColumn.Width = 100;

            newColumn.DataPropertyName = "kod_gruppa";
            newColumn.DisplayMember = "kod_gruppa";
            newColumn.ValueMember = "kod_gruppa";
            newColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;

        }

        private void qrcode_dataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.Exception.Message == "DataGridViewComboBoxCell value is not valid.")
            {
                object value = qrcode_dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (!((DataGridViewComboBoxColumn)qrcode_dataGridView.Columns[e.ColumnIndex]).Items.Contains(value))
                {
                    ((DataGridViewComboBoxColumn)qrcode_dataGridView.Columns[e.ColumnIndex]).Items.Add(value);
                    e.ThrowException = false;
                }
            }
        }

        public string convert_date_main_function_INDATAGRIDVIEW(string sample)
        {
            try
            {
                sample = Regex.Replace(sample, "[^0-9.]", "");
                string[] strArray = sample.Replace(',', '.').Split('.');
                string s1 = strArray[0].Trim();
                string s2 = strArray[1].Trim();
                string s3 = strArray[2].Trim();
                if (int.Parse(s2) <= 12 && int.Parse(s3) < 3000)
                {
                    if (s1.Length == 1)
                        s1 = "0" + s1;
                    if (s2.Length == 1)
                        s2 = "0" + s2;
                    if (s3.Length == 2)
                        s3 = "20" + s3;
                    sample = s1 + "." + s2 + "." + s3;
                }
                else
                    sample = null;
            }
            catch (Exception ex)
            {
                sample = null;
                Console.WriteLine(ex.Message);
            }


            return sample;
        }

        double summa_iznos = 0;
        private void qrcode_dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridViewRow dgvRow = qrcode_dataGridView.CurrentRow;

                if (qrcode_dataGridView.SelectedCells.Count > 0)
                {


                    if (e.ColumnIndex == 2)
                    {
                        sql.myReader = sql.return_MySqlCommand(" SELECT * FROM gruppa_jur7 where kod_gruppa='" + dgvRow.Cells[2].Value + "' ").ExecuteReader();
                        while (sql.myReader.Read())
                        {
                            //schet,subschet/16
                            dgvRow.Cells[10].Value = sql.myReader.GetString("schet");
                            dgvRow.Cells[11].Value = sql.myReader.GetString("subschet");

                            summa_iznos = sql.myReader["prosent_izn"] != DBNull.Value ? Convert.ToDouble(sql.myReader.GetString("prosent_izn")) : 0;


                            if (jur_order_qrcode_textBox.Text == "3")
                            {
                                dgvRow.Cells[12].Value = "159";
                            }
                            else
                            {
                                dgvRow.Cells[12].Value = "280";
                            }

                            //dgvRow.Cells[16].Value = sql.myReader.GetString("subschet");

                        }
                        sql.myReader.Close();


                    }


                    if (e.ColumnIndex == 3)
                    {
                        add_items();
                    }


                    if (e.ColumnIndex == 7)
                    {
                        // Console.WriteLine(dgvRow.Cells[7].Value);
                        dgvRow.Cells[9].Value = string.Format("{0:#0.00}", (dgvRow.Cells[7].Value != null ? (Convert.ToDouble(dgvRow.Cells[7].Value.ToString().Replace(".", ","))) : 0) *
                                                                           (dgvRow.Cells[8].Value != null ? (Convert.ToDouble(dgvRow.Cells[8].Value.ToString().Replace(".", ","))) : 0)
                                                                           );
                        //dgvRow.Cells[16].Value = string.Format("{0:#0.00}", (dgvRow.Cells[7].Value != null ? (Convert.ToDouble(dgvRow.Cells[7].Value.ToString().Replace(".", ","))) : 0) *
                        //                                                   (dgvRow.Cells[16].Value != null ? (Convert.ToDouble(dgvRow.Cells[16].Value.ToString().Replace(".", ","))) : 0)
                        //                                                   );

                    }

                    if (e.ColumnIndex == 8)
                    {
                        dgvRow.Cells[9].Value = string.Format("{0:#0.00}", (dgvRow.Cells[7].Value != null ? (Convert.ToDouble(dgvRow.Cells[7].Value.ToString().Replace(".", ","))) : 0) *
                                                                           (dgvRow.Cells[8].Value != null ? (Convert.ToDouble(dgvRow.Cells[8].Value.ToString().Replace(".", ","))) : 0));


                    }

                    if (e.ColumnIndex == 17)
                    {
                        dgvRow.Cells[17].Value = String.Format("{0:dd.MM.yyyy}", convert_date_main_function_INDATAGRIDVIEW(dgvRow.Cells[17].Value.ToString()));
                    }

                }

                label_update_prixod();

            }
            catch (Exception ex)
            {
                //sql.myReader.Close();
                //sql2.myReader.Close();
                MessageBox.Show("prixod_dataGridView_CellValueChanged " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        public void label_update_prixod()
        {
            double summa = 0;
            double iznos = 0;


            foreach (DataGridViewRow row in qrcode_dataGridView.Rows)
            {
                summa = summa + (row.Cells[9].Value != null ? Double.Parse(row.Cells[9].Value.ToString()) : 0);

                iznos = iznos + (row.Cells[16].Value != null ? Double.Parse(row.Cells[16].Value.ToString()) : 0);

            }
            if (summa.ToString().Length <= 3)
            {
                qrcode_obshiy_summa_label.Text = string.Format("{0:#0.00}", summa);
            }
            if (summa.ToString().Length > 3)
            {
                qrcode_obshiy_summa_label.Text = string.Format("{0:#0,000.00}", summa);
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



        DateTime _lastKeystroke = new DateTime(0);
        string _barcode = string.Empty;

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            bool res = processKey(keyData);
            return keyData == Keys.Enter ? res : base.ProcessCmdKey(ref msg, keyData);
        }

        
        string vid_doc = "";
        public void clear()
        {

        }
        bool processKey(Keys key)
        {


            TimeSpan elapsed = (DateTime.Now - _lastKeystroke);
            if (elapsed.TotalMilliseconds > 50)
            {
                _barcode = string.Empty;
            }

            _barcode += (char)key;
            _lastKeystroke = DateTime.Now;

            if (key == Keys.Enter && _barcode.Length > 1)
            {
                string msg = new String(_barcode.ToArray());
                //textBox1.Text = msg;
                string spl = msg.Replace('\u0010', ' ');

                string spl2 = spl.Replace('\r', ' ');
                string spl3 = spl2.Replace(" ", "").Replace("½", "-");
                string nw = msg;

                try
                {
                    this.qrcode_dataGridView.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.qrcode_dataGridView_CellValueChanged);

                    string debet_01 = "";
                    double debet_01_sum = 0;
                    double debet_06_sum = 0;
                    double debet_07_sum = 0;

                    jur_order_qrcode_textBox.Text = "";
                    num_qrcode_textBox.Text = "";
                    primech_qrcode_textBox.Text = "";
                    doveren_qrcode_textBox.Text = "";
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                    textBox4.Text = "";

                    var query = "SELECT * FROM products_jur7 where concat(user,'-',year,'-',month,'-',kod_doc,'-',vid_doc) like '%" + spl3 + "%' group by kod_doc";
                    sql.myReader = sql.return_MySqlCommand(query).ExecuteReader();
                    while (sql.myReader.Read())
                    {
                        jur_order_qrcode_textBox.Text = (sql.myReader["jur_order"] != DBNull.Value ? sql.myReader.GetString("jur_order") : "");
                        num_qrcode_textBox.Text = (sql.myReader["num_doc"] != DBNull.Value ? sql.myReader.GetString("num_doc") : "");
                        data_qrcode_DateTimePicker.Value = (sql.myReader["date_doc"] != DBNull.Value ? sql.myReader.GetDateTime("date_doc") : DateTime.Now);
                        primech_qrcode_textBox.Text = (sql.myReader["primech"] != DBNull.Value ? sql.myReader.GetString("primech") : "");
                        doveren_qrcode_textBox.Text = (sql.myReader["doverennost"] != DBNull.Value ? sql.myReader.GetString("doverennost") : "");
                        textBox1.Text = (sql.myReader["ot_kogo"] != DBNull.Value ? sql.myReader.GetString("ot_kogo") : "");
                        textBox2.Text = (sql.myReader["ot_kogo_2"] != DBNull.Value ? sql.myReader.GetString("ot_kogo_2") : "");
                        textBox3.Text = (sql.myReader["komu_1"] != DBNull.Value ? sql.myReader.GetString("komu_1") : "");
                        textBox4.Text = (sql.myReader["komu_2"] != DBNull.Value ? sql.myReader.GetString("komu_2") : "");

                        kod_num_textBox.Text = (sql.myReader["kod_doc"] != DBNull.Value ? sql.myReader.GetString("kod_doc") : "");

                        year_textBox.Text = (sql.myReader["year"] != DBNull.Value ? sql.myReader.GetString("year") : "");
                        month_textBox.Text = (sql.myReader["month"] != DBNull.Value ? sql.myReader.GetString("month") : "");

                        vid_doc = (sql.myReader["vid_doc"] != DBNull.Value ? sql.myReader.GetString("vid_doc") : "");
                    }
                    sql.myReader.Close();

                    
                    

                    qrcode_dataGridView.Rows.Clear();

                    int id_product = 0;


                    var select_ras = " SELECT * FROM products_jur7 where concat(user,'-',year,'-',month,'-',kod_doc,'-',vid_doc) like '%" + spl3 + "%' ";
                    sql.myReader = sql.return_MySqlCommand(select_ras).ExecuteReader();
                    while (sql.myReader.Read())
                    {

                        //kod_tov,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,

                        int index = qrcode_dataGridView.Rows.Add();
                        qrcode_dataGridView.Rows[index].Cells[0].Value = (sql.myReader["id"] != DBNull.Value ? sql.myReader.GetString("id") : "");
                        qrcode_dataGridView.Rows[index].Cells[1].Value = (sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetString("product_id") : "");
                        qrcode_dataGridView.Rows[index].Cells[2].Value = (sql.myReader["gruppa"] != DBNull.Value ? sql.myReader.GetString("gruppa") : "");
                        qrcode_dataGridView.Rows[index].Cells[3].Value = (sql.myReader["naim_tov"] != DBNull.Value ? sql.myReader.GetString("naim_tov") : "");
                        qrcode_dataGridView.Rows[index].Cells[4].Value = (sql.myReader["edin"] != DBNull.Value ? sql.myReader.GetString("edin") : "");
                        qrcode_dataGridView.Rows[index].Cells[5].Value = (sql.myReader["inventar_num"] != DBNull.Value ? sql.myReader.GetString("inventar_num") : "");
                        qrcode_dataGridView.Rows[index].Cells[6].Value = (sql.myReader["seria_num"] != DBNull.Value ? sql.myReader.GetString("seria_num") : "");
                        string kols = sql.myReader["kol"] != DBNull.Value ? sql.myReader.GetString("kol") : "";

                        if (kols.Length <= 3)
                        {
                            qrcode_dataGridView.Rows[qrcode_dataGridView.Rows.Count - 2].Cells[7].Value = string.Format("{0:#0.00}", (sql.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("kol").Replace(".", ","))) : 0));
                        }
                        if (kols.Length > 3)
                        {
                            qrcode_dataGridView.Rows[qrcode_dataGridView.Rows.Count - 2].Cells[7].Value = string.Format("{0:#,###.00}", (sql.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("kol").Replace(".", ","))) : 0));
                        }

                        string sena = sql.myReader["sena"] != DBNull.Value ? sql.myReader.GetString("sena") : "";

                        if (sena.Length <= 3)
                        {
                            qrcode_dataGridView.Rows[qrcode_dataGridView.Rows.Count - 2].Cells[8].Value = string.Format("{0:#0.00}", (sql.myReader["sena"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("sena").Replace(".", ","))) : 0));
                        }
                        if (sena.Length > 3)
                        {
                            qrcode_dataGridView.Rows[qrcode_dataGridView.Rows.Count - 2].Cells[8].Value = string.Format("{0:#,###.00}", (sql.myReader["sena"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("sena").Replace(".", ","))) : 0));
                        }

                        string summa = sql.myReader["summa"] != DBNull.Value ? sql.myReader.GetString("summa") : "";

                        if (summa.Length <= 3)
                        {
                            qrcode_dataGridView.Rows[qrcode_dataGridView.Rows.Count - 2].Cells[9].Value = string.Format("{0:#0.00}", (sql.myReader["summa"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa").Replace(".", ","))) : 0));
                        }
                        if (summa.Length > 3)
                        {
                            qrcode_dataGridView.Rows[qrcode_dataGridView.Rows.Count - 2].Cells[9].Value = string.Format("{0:#,###.00}", (sql.myReader["summa"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa").Replace(".", ","))) : 0));
                        }

                        //deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,id_products

                        qrcode_dataGridView.Rows[index].Cells[10].Value = (sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch") : "");


                        debet_01 = (sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch") : "");
                        string first = debet_01.Substring(0, 2);

                        if (first == "01")
                        {
                            debet_01_sum += (qrcode_dataGridView.Rows[index].Cells[9].Value != null ? Double.Parse(qrcode_dataGridView.Rows[index].Cells[9].Value.ToString()) : 0);
                        }
                        else if (first == "06")
                        {
                            debet_06_sum += (qrcode_dataGridView.Rows[index].Cells[9].Value != null ? Double.Parse(qrcode_dataGridView.Rows[index].Cells[9].Value.ToString()) : 0);
                        }
                        else if (first == "07")
                        {
                            debet_07_sum += (qrcode_dataGridView.Rows[index].Cells[9].Value != null ? Double.Parse(qrcode_dataGridView.Rows[index].Cells[9].Value.ToString()) : 0);
                        }

                        qrcode_dataGridView.Rows[index].Cells[11].Value = (sql.myReader["deb_sch_2"] != DBNull.Value ? sql.myReader.GetString("deb_sch_2") : "");
                        qrcode_dataGridView.Rows[index].Cells[12].Value = (sql.myReader["kre_sch"] != DBNull.Value ? sql.myReader.GetString("kre_sch") : "");
                        qrcode_dataGridView.Rows[index].Cells[13].Value = (sql.myReader["kre_sch_2"] != DBNull.Value ? sql.myReader.GetString("kre_sch_2") : "");
                        qrcode_dataGridView.Rows[index].Cells[14].Value = (sql.myReader["provodka_iznos"] != DBNull.Value ? sql.myReader.GetString("provodka_iznos") : "");
                        qrcode_dataGridView.Rows[index].Cells[15].Value = (sql.myReader["provodka_iznos_2"] != DBNull.Value ? sql.myReader.GetString("provodka_iznos_2") : "");

                        string summa_iznos = sql.myReader["summa_iznos"] != DBNull.Value ? sql.myReader.GetString("summa_iznos") : "";

                        if (summa_iznos.Length <= 3)
                        {
                            qrcode_dataGridView.Rows[qrcode_dataGridView.Rows.Count - 2].Cells[16].Value = string.Format("{0:#0.00}", (sql.myReader["summa_iznos"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa_iznos").Replace(".", ","))) : 0));
                        }
                        if (summa_iznos.Length > 3)
                        {
                            qrcode_dataGridView.Rows[qrcode_dataGridView.Rows.Count - 2].Cells[16].Value = string.Format("{0:#,###.00}", (sql.myReader["summa_iznos"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa_iznos").Replace(".", ","))) : 0));
                        }

                        qrcode_dataGridView.Rows[index].Cells[17].Value = (sql.myReader["date_pr"] != DBNull.Value ? (DateTime.Parse(sql.myReader.GetString("date_pr")).ToString("dd.MM.yyyy")) : "");

                        

                        var id_sklad_products = " SELECT id FROM products_prixod_jur7 where id_sklad_products = '"+ (sql.myReader["id"] != DBNull.Value ? sql.myReader.GetString("id") : "") + "' ";
                        sql2.myReader = sql2.return_MySqlCommand(id_sklad_products).ExecuteReader();
                        while (sql2.myReader.Read())
                        {
                            id_product = (sql2.myReader["id"] != DBNull.Value ? sql2.myReader.GetInt32("id") : 0);
                            qrcode_dataGridView.Rows[index].Cells[18].Value = id_product;
                        }
                        sql2.myReader.Close();

                        var id_sklad_products2 = " SELECT id FROM products_rasxod_jur7 where concat(user,'-',year,'-',month,'-',kod_doc,'-',vid_doc) like '%" + spl3 + "%' ";
                        sql2.myReader = sql2.return_MySqlCommand(id_sklad_products2).ExecuteReader();
                        while (sql2.myReader.Read())
                        {
                            id_product = (sql2.myReader["id"] != DBNull.Value ? sql2.myReader.GetInt32("id") : 0);
                            qrcode_dataGridView.Rows[index].Cells[18].Value = id_product;
                        }
                        sql2.myReader.Close();

                        var id_sklad_products3 = " SELECT id FROM products_vnut_per_jur7 where concat(user,'-',year,'-',month,'-',kod_doc,'-',vid_doc) like '%" + spl3 + "%' ";
                        sql2.myReader = sql2.return_MySqlCommand(id_sklad_products3).ExecuteReader();
                        while (sql2.myReader.Read())
                        {
                            id_product = (sql2.myReader["id"] != DBNull.Value ? sql2.myReader.GetInt32("id") : 0);
                            qrcode_dataGridView.Rows[index].Cells[18].Value = id_product;
                        }
                        sql2.myReader.Close();

                       
                        //sklad_dataGridView.Rows[index].Cells[3].Value = refresh_strings_to_mysql(sql.myReader["sena"] != DBNull.Value ? string.Format("{0:#0.00}", sql.myReader.GetDouble("sena")) : "0");
                        //qrcode_dataGridView.Rows[index].Cells[18].Value = (sql.myReader["id_products"] != DBNull.Value ? sql.myReader.GetString("id_products") : "");


                    }
                    sql.myReader.Close();


                    if (debet_01_sum.ToString().Length <= 3)
                    {
                        nol_bir_lbl.Text = string.Format("{0:#0.00}", debet_01_sum);
                    }
                    if (debet_01_sum.ToString().Length > 3)
                    {
                        nol_bir_lbl.Text = string.Format("{0:#0,000.00}", debet_01_sum);
                    }

                    if (debet_06_sum.ToString().Length <= 3)
                    {
                        nol_olti_lbl.Text = string.Format("{0:#0.00}", debet_06_sum);
                    }
                    if (debet_06_sum.ToString().Length > 3)
                    {
                        nol_olti_lbl.Text = string.Format("{0:#0,000.00}", debet_06_sum);
                    }
                    if (debet_07_sum.ToString().Length <= 3)
                    {
                        nol_7_lbl.Text = string.Format("{0:#0.00}", debet_07_sum);
                    }
                    if (debet_07_sum.ToString().Length > 3)
                    {
                        nol_7_lbl.Text = string.Format("{0:#0,000.00}", debet_07_sum);
                    }


                    this.qrcode_dataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.qrcode_dataGridView_CellValueChanged);
                    label_update_prixod();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Хато маълумот киритилган (" + ex.Message + ")");
                }


                _barcode = string.Empty;
                return true;
            }
            return false;
        }

        //DialogResult dialogResult = MessageBox.Show("Бу маълумот киритилган !!!", "Поставщик бўйича",
        //MessageBoxButtons.OK, MessageBoxIcon.Error);
        public void KeyPress_scanner_preview(object sender, KeyPressEventArgs e)
        {

        }

        string fio_gl_bugalter = "";
        string fio_bugalter = "";
        string inspektor = "";
        private void vnut_per_document_btn_Click(object sender, EventArgs e)
        {
            //try
            //{
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.LeftMargin = 0.2;
            sheet.PageSetup.RightMargin = 0.2;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;


            sheet.PageSetup.Orientation = PageOrientationType.Portrait;


            sheet.Range["a1:a1"].ColumnWidth = 3;
            sheet.Range["b1:b1"].ColumnWidth = 29;
            sheet.Range["c1:c1"].ColumnWidth = 4.86;
            sheet.Range["d1:d1"].ColumnWidth = 8;
            sheet.Range["e1:e1"].ColumnWidth = 9;
            sheet.Range["f1:f1"].ColumnWidth = 12;
            sheet.Range["g1:g1"].ColumnWidth = 4;
            sheet.Range["h1:h1"].ColumnWidth = 7;
            sheet.Range["i1:i1"].ColumnWidth = 4;
            sheet.Range["j1:j1"].ColumnWidth = 7;
            sheet.Range["k1:k1"].ColumnWidth = 9;

            string name_organ = "";
            var name_org = "SELECT * FROM spravochnik_main_jur7 where user_jur7='" + string_for_otdels + "'";

            sql.myReader = sql.return_MySqlCommand(name_org).ExecuteReader();
            while (sql.myReader.Read())
            {
                name_organ = (sql.myReader["naim_org"] != DBNull.Value ? sql.myReader.GetString("naim_org") : "");
            }
            sql.myReader.Close();

            sheet.Range["a1:k1"].Style.Font.IsBold = true;
            sheet.Range["a1:k1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a1:k1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a1:k1"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a1:k1"].Style.Font.Size = 14;
            sheet.Range["a1:k1"].Merge(); // birlashtirish
            sheet.Range["a1:k1"].Text = name_organ;
            sheet.Range["a1:k1"].Style.Font.Color = Color.DarkBlue;
            sheet.SetRowHeight(1, 21);



            sheet.Range["a2:k2"].Style.Font.IsBold = true;
            sheet.Range["a2:k2"].Style.Font.IsItalic = true;
            sheet.Range["a2:k2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a2:k2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a2:k2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a2:k2"].Style.Font.Size = 14;
            sheet.Range["a2:k2"].Merge(); // birlashtirish
            sheet.Range["a2:k2"].Text = "ПРИЕМНЫЙ АКТ № " + num_qrcode_textBox.Text + " от " + data_qrcode_DateTimePicker.Value.ToString("dd.MM.yyyy");
            sheet.Range["a2:k2"].Style.WrapText = true;
            sheet.SetRowHeight(2, 21);

            sheet.Range["b3:j3"].Style.Font.IsBold = true;
            sheet.Range["b3:j3"].Style.Font.IsItalic = true;
            sheet.Range["b3:j3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b3:j3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b3:j3"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b3:j3"].Style.Font.Size = 12;
            sheet.Range["b3:j3"].Merge(); // birlashtirish
            sheet.Range["b3:j3"].Text = "Откуда: " + textBox1.Text;
            sheet.Range["b3:j3"].Style.WrapText = true;
            sheet.SetRowHeight(3, 18);

            sheet.Range["b4:j4"].Style.Font.IsBold = true;
            sheet.Range["b4:j4"].Style.Font.IsItalic = true;
            sheet.Range["b4:j4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b4:j4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b4:j4"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b4:j4"].Style.Font.Size = 10;
            sheet.Range["b4:j4"].Merge(); // birlashtirish
            sheet.Range["b4:j4"].Text = "Кому: " + textBox4.Text + "  " + textBox3.Text;
            sheet.Range["b4:j4"].Style.WrapText = true;
            sheet.SetRowHeight(4, 18);

            sheet.Range["a5:a5"].Style.Font.IsBold = true;
            sheet.Range["a5:a5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a5:a5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a5:a5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a5:a5"].Style.Font.Size = 11;
            sheet.Range["a5:a5"].Merge(); // birlashtirish
            sheet.Range["a5:a5"].Text = "№";
            sheet.Range["a5:a5"].Style.WrapText = true;
            sheet.Range["a5:a5"].BorderAround(LineStyleType.Thin);

            sheet.Range["b5:b5"].Style.Font.IsBold = true;
            sheet.Range["b5:b5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b5:b5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b5:b5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b5:b5"].Style.Font.Size = 11;
            sheet.Range["b5:b5"].Merge(); // birlashtirish
            sheet.Range["b5:b5"].Text = "Наименование";
            sheet.Range["b5:b5"].Style.WrapText = true;
            sheet.Range["b5:b5"].BorderAround(LineStyleType.Thin);

            sheet.Range["c5:c5"].Style.Font.IsBold = true;
            sheet.Range["c5:c5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c5:c5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c5:c5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c5:c5"].Style.Font.Size = 10;
            sheet.Range["c5:c5"].Merge(); // birlashtirish
            sheet.Range["c5:c5"].Text = "Ед.из";
            sheet.Range["c5:c5"].BorderAround(LineStyleType.Thin);
            sheet.Range["c5:c5"].Style.WrapText = true;

            sheet.Range["d5:d5"].Style.Font.IsBold = true;
            sheet.Range["d5:d5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d5:d5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d5:d5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d5:d5"].Style.Font.Size = 11;
            sheet.Range["d5:d5"].Merge(); // birlashtirish
            sheet.Range["d5:d5"].Text = "Кол.";
            sheet.Range["d5:d5"].Style.WrapText = true;
            sheet.Range["d5:d5"].BorderAround(LineStyleType.Thin);

            sheet.Range["e5:e5"].Style.Font.IsBold = true;
            sheet.Range["e5:e5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e5:e5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e5:e5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e5:e5"].Style.Font.Size = 11;
            sheet.Range["e5:e5"].Merge(); // birlashtirish
            sheet.Range["e5:e5"].Text = "Цена";
            sheet.Range["e5:e5"].Style.WrapText = true;
            sheet.Range["e5:e5"].BorderAround(LineStyleType.Thin);

            sheet.Range["f5:f5"].Style.Font.IsBold = true;
            sheet.Range["f5:f5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f5:f5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f5:f5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f5:f5"].Style.Font.Size = 11;
            sheet.Range["f5:f5"].Merge(); // birlashtirish
            sheet.Range["f5:f5"].Text = "Сумма";
            sheet.Range["f5:f5"].Style.WrapText = true;
            sheet.Range["f5:f5"].BorderAround(LineStyleType.Thin);

            sheet.Range["g5:h5"].Style.Font.IsBold = true;
            sheet.Range["g5:h5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g5:h5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g5:h5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g5:h5"].Style.Font.Size = 11;
            sheet.Range["g5:h5"].Merge(); // birlashtirish
            sheet.Range["g5:h5"].Text = "Дебет";
            sheet.Range["g5:h5"].Style.WrapText = true;
            sheet.Range["g5:h5"].BorderAround(LineStyleType.Thin);

            sheet.Range["i5:j5"].Style.Font.IsBold = true;
            sheet.Range["i5:j5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["i5:j5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["i5:j5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["i5:j5"].Style.Font.Size = 11;
            sheet.Range["i5:j5"].Merge(); // birlashtirish
            sheet.Range["i5:j5"].Text = "Кредит";
            sheet.Range["i5:j5"].Style.WrapText = true;
            sheet.Range["i5:j5"].BorderAround(LineStyleType.Thin);

            sheet.Range["k5:k5"].Style.Font.IsBold = true;
            sheet.Range["k5:k5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["k5:k5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["k5:k5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["k5:k5"].Style.Font.Size = 11;
            sheet.Range["k5:k5"].Merge(); // birlashtirish
            sheet.Range["k5:k5"].Text = "Дата";
            sheet.Range["k5:k5"].Style.WrapText = true;
            sheet.Range["k5:k5"].BorderAround(LineStyleType.Thin);


            int i = 0;
            int myrow = 6;
            int j = 0;



            var top = " SELECT * FROM products_jur7 where kod_doc='" + kod_num_textBox.Text + "' and user='" + string_for_otdels + "' and month='" + month_textBox.Text + "' and year='" + year_textBox.Text + "'  ";
            sql.myReader = sql.return_MySqlCommand(top).ExecuteReader();
            while (sql.myReader.Read())
            {
                j = i;
                j = j + 1;

                sheet.Range["a" + myrow + ":a" + myrow].Merge();
                sheet.Range["a" + myrow + ":a" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["a" + myrow + ":a" + myrow].Style.VerticalAlignment = VerticalAlignType.Bottom;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.Size = 10;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["a" + myrow + ":a" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["a" + myrow + ":a" + myrow].Text = (String)j.ToString();

                sheet.Range["b" + myrow + ":b" + myrow].Merge();
                sheet.Range["b" + myrow + ":b" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["b" + myrow + ":b" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["b" + myrow + ":b" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.Size = 10;
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["b" + myrow + ":b" + myrow].Text = sql.myReader["naim_tov"] != DBNull.Value ? sql.myReader.GetString("naim_tov").ToString() : "";
                sheet.Range["b" + myrow + ":b" + myrow].Style.WrapText = true;

                sheet.Range["c" + myrow + ":c" + myrow].Merge();
                sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 10;
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["c" + myrow + ":c" + myrow].Text = sql.myReader["edin"] != DBNull.Value ? sql.myReader.GetString("edin").ToString() : "";
                sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;

                sheet.Range["d" + myrow + ":d" + myrow].Merge();
                sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
                sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["d" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 10;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["d" + myrow + ":d" + myrow].Value = sql.myReader["kol"] != DBNull.Value ? sql.myReader.GetString("kol").ToString() : "";

                sheet.Range["e" + myrow + ":e" + myrow].Merge();
                sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 10;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["e" + myrow + ":e" + myrow].Value = sql.myReader["sena"] != DBNull.Value ? sql.myReader.GetString("sena").ToString() : "";

                sheet.Range["f" + myrow + ":f" + myrow].Merge();
                sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":f" + myrow].Value = sql.myReader["summa"] != DBNull.Value ? sql.myReader.GetString("summa").ToString() : "";

                sheet.Range["g" + myrow + ":g" + myrow].Merge();
                sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
                sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["g" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 10;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["g" + myrow + ":g" + myrow].Text = sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch").ToString() : "";

                sheet.Range["h" + myrow + ":h" + myrow].Merge();
                sheet.Range["h" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["h" + myrow + ":h" + myrow].Style.WrapText = true;
                sheet.Range["h" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["h" + myrow + ":h" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.Size = 10;
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["h" + myrow + ":h" + myrow].Text = sql.myReader["deb_sch_2"] != DBNull.Value ? sql.myReader.GetString("deb_sch_2").ToString() : "";

                sheet.Range["i" + myrow + ":i" + myrow].Merge();
                sheet.Range["i" + myrow + ":i" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["i" + myrow + ":i" + myrow].Style.WrapText = true;
                sheet.Range["i" + myrow + ":i" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["i" + myrow + ":i" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["i" + myrow + ":i" + myrow].Style.Font.Size = 10;
                sheet.Range["i" + myrow + ":i" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["i" + myrow + ":i" + myrow].Text = sql.myReader["kre_sch"] != DBNull.Value ? sql.myReader.GetString("kre_sch").ToString() : "";

                sheet.Range["j" + myrow + ":j" + myrow].Merge();
                sheet.Range["j" + myrow + ":j" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["j" + myrow + ":j" + myrow].Style.WrapText = true;
                sheet.Range["j" + myrow + ":j" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["j" + myrow + ":j" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["j" + myrow + ":j" + myrow].Style.Font.Size = 10;
                sheet.Range["j" + myrow + ":j" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["j" + myrow + ":j" + myrow].Text = sql.myReader["kre_sch_2"] != DBNull.Value ? sql.myReader.GetString("kre_sch_2").ToString() : "";

                sheet.Range["k" + myrow + ":k" + myrow].Merge();
                sheet.Range["k" + myrow + ":k" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["k" + myrow + ":k" + myrow].Style.WrapText = true;
                sheet.Range["k" + myrow + ":k" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["k" + myrow + ":k" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["k" + myrow + ":k" + myrow].Style.Font.Size = 10;
                sheet.Range["k" + myrow + ":k" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["k" + myrow + ":k" + myrow].Text = sql.myReader["date_pr"] != DBNull.Value ? Convert.ToDateTime(sql.myReader.GetString("date_pr").ToString()).ToString("dd.MM.yyyy") : "";

                myrow = myrow + 1;
                i = i + 1;

            }
            sql.myReader.Close();


            sheet.Range["e" + myrow + ":e" + myrow].Merge();
            sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.IsBold = true;
            sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 11;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["e" + myrow + ":e" + myrow].Text = "Всего:";
            sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);

            sheet.Range["f" + myrow + ":g" + myrow].Merge();
            sheet.Range["f" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["f" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.Size = 11;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["f" + myrow + ":g" + myrow].Value = qrcode_obshiy_summa_label.Text;
            sheet.Range["f" + myrow + ":g" + myrow].Style.WrapText = true;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.IsBold = true;
            sheet.Range["f" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);
            myrow++;


            var spravocnik = "SELECT fio_gl_bugalter,fio_bugalter,inspektor FROM spravochnik_main_jur7 where user_jur7='" + string_for_otdels + "'";

            sql.myReader = sql.return_MySqlCommand(spravocnik).ExecuteReader();
            while (sql.myReader.Read())
            {
                fio_gl_bugalter = sql.myReader["fio_gl_bugalter"] != DBNull.Value ? sql.myReader.GetString("fio_gl_bugalter").ToString() : "";
                fio_bugalter = sql.myReader["fio_bugalter"] != DBNull.Value ? sql.myReader.GetString("fio_bugalter").ToString() : "";
                inspektor = sql.myReader["inspektor"] != DBNull.Value ? sql.myReader.GetString("inspektor").ToString() : "";
            }
            sql.myReader.Close();


            sheet.Range["b" + myrow + ":d" + myrow].Merge();
            sheet.Range["b" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":d" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":d" + myrow].Text = fio_gl_bugalter;
            sheet.SetRowHeight(myrow, 18);
            myrow++;

            sheet.Range["b" + myrow + ":d" + myrow].Merge();
            sheet.Range["b" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":d" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":d" + myrow].Text = fio_bugalter;

            sheet.Range["e" + myrow + ":e" + myrow].Merge();
            sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.IsBold = true;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.IsItalic = true;
            sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 10;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["e" + myrow + ":e" + myrow].Text = "Деб.сч.";
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Color = Color.DarkBlue;

            sheet.Range["f" + myrow + ":f" + myrow].Merge();
            sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.IsBold = true;
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.IsItalic = true;
            sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["f" + myrow + ":f" + myrow].Text = "Деб.сч.";
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Color = Color.DarkBlue;

            sheet.Range["g" + myrow + ":h" + myrow].Merge();
            sheet.Range["g" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g" + myrow + ":h" + myrow].Style.WrapText = true;
            sheet.Range["g" + myrow + ":h" + myrow].Style.Font.IsBold = true;
            sheet.Range["g" + myrow + ":h" + myrow].Style.Font.IsItalic = true;
            sheet.Range["g" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g" + myrow + ":h" + myrow].Style.Font.Size = 10;
            sheet.Range["g" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["g" + myrow + ":h" + myrow].Text = "Кре.сч.";
            sheet.Range["g" + myrow + ":h" + myrow].Style.Font.Color = Color.DarkBlue;

            sheet.Range["j" + myrow + ":k" + myrow].Merge();
            sheet.Range["j" + myrow + ":k" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["j" + myrow + ":k" + myrow].Style.WrapText = true;
            sheet.Range["j" + myrow + ":k" + myrow].Style.Font.IsBold = true;
            sheet.Range["j" + myrow + ":k" + myrow].Style.Font.IsItalic = true;
            sheet.Range["j" + myrow + ":k" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["j" + myrow + ":k" + myrow].Style.Font.Size = 10;
            sheet.Range["j" + myrow + ":k" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["j" + myrow + ":k" + myrow].Text = "Сумма";
            sheet.Range["j" + myrow + ":k" + myrow].Style.Font.Color = Color.DarkBlue;

            sheet.SetRowHeight(myrow, 18);

            myrow++;

            sheet.Range["b" + myrow + ":d" + myrow].Merge();
            sheet.Range["b" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":d" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":d" + myrow].Text = inspektor;

            var schet = "SELECT deb_sch,deb_sch_2,kre_sch,sum(summa) as summa FROM products_jur7 where kod_doc='" + kod_num_textBox.Text + "' and user='" + string_for_otdels + "' and month='" + month_textBox.Text + "' and year='" + year_textBox.Text + "' group by deb_sch";

            sql.myReader = sql.return_MySqlCommand(schet).ExecuteReader();
            while (sql.myReader.Read())
            {
                sheet.Range["e" + myrow + ":e" + myrow].Merge();
                sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.IsBold = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.IsItalic = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 10;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["e" + myrow + ":e" + myrow].Text = sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch").ToString() : "";

                sheet.Range["f" + myrow + ":f" + myrow].Merge();
                sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.IsBold = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.IsItalic = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":f" + myrow].Text = sql.myReader["deb_sch_2"] != DBNull.Value ? sql.myReader.GetString("deb_sch_2").ToString() : "";

                sheet.Range["g" + myrow + ":h" + myrow].Merge();
                sheet.Range["g" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["g" + myrow + ":h" + myrow].Style.WrapText = true;
                sheet.Range["g" + myrow + ":h" + myrow].Style.Font.IsBold = true;
                sheet.Range["g" + myrow + ":h" + myrow].Style.Font.IsItalic = true;
                sheet.Range["g" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["g" + myrow + ":h" + myrow].Style.Font.Size = 10;
                sheet.Range["g" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["g" + myrow + ":h" + myrow].Text = sql.myReader["kre_sch"] != DBNull.Value ? sql.myReader.GetString("kre_sch").ToString() : "";

                sheet.Range["j" + myrow + ":k" + myrow].Merge();
                sheet.Range["j" + myrow + ":k" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["j" + myrow + ":k" + myrow].Style.WrapText = true;
                sheet.Range["j" + myrow + ":k" + myrow].Style.Font.IsBold = true;
                sheet.Range["j" + myrow + ":k" + myrow].Style.Font.IsItalic = true;
                sheet.Range["j" + myrow + ":k" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["j" + myrow + ":k" + myrow].Style.Font.Size = 10;
                sheet.Range["j" + myrow + ":k" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["j" + myrow + ":k" + myrow].Value = sql.myReader["summa"] != DBNull.Value ? sql.myReader.GetString("summa").ToString() : "";

                sheet.SetRowHeight(myrow, 18);

                myrow++;
            }
            sql.myReader.Close();


            sheet.Range["b" + myrow + ":d" + myrow].Merge();
            sheet.Range["b" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":d" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Top;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":d" + myrow].Text = "Получил:________________________ ";
            sheet.SetRowHeight(myrow, 18);
            sheet.Range["d5:" + myrow + "k"].NumberFormat = "#,##0.00";


            workbook.SaveToFile(Environment.CurrentDirectory + "\\docs\\Документ.xlsx", Spire.Xls.FileFormat.Version2007);
            System.Diagnostics.Process.Start(workbook.FileName);



            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Документ_excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void vnut_per_izv_btn_Click(object sender, EventArgs e)
        {
            //try
            //{
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.LeftMargin = 0.2;
            sheet.PageSetup.RightMargin = 0.2;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;


            sheet.PageSetup.Orientation = PageOrientationType.Portrait;


            sheet.Range["a1:a1"].ColumnWidth = 3;
            sheet.Range["b1:b1"].ColumnWidth = 29;
            sheet.Range["c1:c1"].ColumnWidth = 4.86;
            sheet.Range["d1:d1"].ColumnWidth = 8;
            sheet.Range["e1:e1"].ColumnWidth = 9;
            sheet.Range["f1:f1"].ColumnWidth = 12;
            sheet.Range["g1:g1"].ColumnWidth = 4;
            sheet.Range["h1:h1"].ColumnWidth = 7;
            sheet.Range["i1:i1"].ColumnWidth = 4;
            sheet.Range["j1:j1"].ColumnWidth = 7;
            sheet.Range["k1:k1"].ColumnWidth = 9;


            string name_organ = "";
            var name_org = "SELECT * FROM spravochnik_main_jur7 where user_jur7='" + string_for_otdels + "'";
            sql.myReader.Close();
            sql.myReader = sql.return_MySqlCommand(name_org).ExecuteReader();
            while (sql.myReader.Read())
            {
                name_organ = (sql.myReader["naim_org"] != DBNull.Value ? sql.myReader.GetString("naim_org") : "");
            }
            sql.myReader.Close();

            sheet.Range["b1:k1"].Style.Font.IsBold = true;
            sheet.Range["b1:k1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b1:k1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b1:k1"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b1:k1"].Style.Font.Size = 14;
            sheet.Range["b1:k1"].Merge(); // birlashtirish
            sheet.Range["b1:k1"].Text = name_organ;
            sheet.Range["b1:k1"].Style.Font.Color = Color.DarkBlue;
            sheet.SetRowHeight(1, 21);


            sheet.Range["a2:k2"].Style.Font.IsBold = true;
            //sheet.Range["a2:k2"].Style.Font.IsItalic = true;
            sheet.Range["a2:k2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a2:k2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a2:k2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a2:k2"].Style.Font.Size = 14;
            sheet.Range["a2:k2"].Merge(); // birlashtirish
            sheet.Range["a2:k2"].Text = "Извещение";
            sheet.Range["a2:k2"].Style.WrapText = true;
            sheet.SetRowHeight(2, 21);

            sheet.Range["b3:k3"].Style.Font.IsBold = true;
            //sheet.Range["b3:k3"].Style.Font.IsItalic = true;
            sheet.Range["b3:k3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b3:k3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b3:k3"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b3:k3"].Style.Font.Size = 12;
            sheet.Range["b3:k3"].Merge(); // birlashtirish
            sheet.Range["b3:k3"].Text = "о безвозмездной передаче основныих средств № " + num_qrcode_textBox.Text + " от " + data_qrcode_DateTimePicker.Value.ToString("dd.MM.yyyy");
            sheet.Range["b3:k3"].Style.WrapText = true;
            sheet.SetRowHeight(3, 18);

            //sheet.Range["a4:k4"].Style.Font.IsBold = true;
            //sheet.Range["a4:k4"].Style.Font.IsItalic = true;
            sheet.Range["a4:k4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a4:k4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a4:k4"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a4:k4"].Style.Font.Size = 11;
            sheet.Range["a4:k4"].Merge(); // birlashtirish
            sheet.Range["a4:k4"].Text = "Кому: " + textBox4.Text;
            sheet.Range["a4:k4"].Style.WrapText = true;
            sheet.Range["a4:k4"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.SetRowHeight(4, 20);

            //sheet.Range["a5:d5"].Style.Font.IsBold = true;
            //sheet.Range["a5:d5"].Style.Font.IsItalic = true;
            sheet.Range["a5:d5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a5:d5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a5:d5"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a5:d5"].Style.Font.Size = 11;
            sheet.Range["a5:d5"].Merge(); // birlashtirish
            sheet.Range["a5:d5"].Text = "Отправителъ: " + name_organ;
            sheet.Range["a5:d5"].Style.WrapText = true;
            sheet.Range["a5:d5"].Style.Font.Underline = FontUnderlineType.SingleAccounting;

            //sheet.Range["e5:k5"].Style.Font.IsBold = true;
            //sheet.Range["e5:k5"].Style.Font.IsItalic = true;
            sheet.Range["e5:k5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e5:k5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e5:k5"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["e5:k5"].Style.Font.Size = 11;
            sheet.Range["e5:k5"].Merge(); // birlashtirish
            sheet.Range["e5:k5"].Text = "Получателъ: " + textBox4.Text;
            sheet.Range["e5:k5"].Style.WrapText = true;
            sheet.Range["e5:k5"].Style.Font.Underline = FontUnderlineType.SingleAccounting;

            sheet.SetRowHeight(5, 20);

            //sheet.Range["a6:k6"].Style.Font.IsBold = true;
            //sheet.Range["a6:k6"].Style.Font.IsItalic = true;
            sheet.Range["a6:k6"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a6:k6"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a6:k6"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a6:k6"].Style.Font.Size = 11;
            sheet.Range["a6:k6"].Merge(); // birlashtirish
            sheet.Range["a6:k6"].Text = "Основание на передачу(распоряжение № и дата): " + primech_qrcode_textBox.Text;
            sheet.Range["a6:k6"].Style.WrapText = true;
            sheet.Range["a6:k6"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.SetRowHeight(6, 20);

            //sheet.Range["a7:c7"].Style.Font.IsBold = true;
            //sheet.Range["a7:c7"].Style.Font.IsItalic = true;
            sheet.Range["a7:c7"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a7:c7"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a7:c7"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a7:c7"].Style.Font.Size = 11;
            sheet.Range["a7:c7"].Merge(); // birlashtirish
            sheet.Range["a7:c7"].Text = "№ Доверенност: " + doveren_qrcode_textBox.Text;
            sheet.Range["a7:c7"].Style.WrapText = true;
            sheet.Range["a7:c7"].Style.Font.Underline = FontUnderlineType.SingleAccounting;

            //sheet.Range["d7:f7"].Style.Font.IsBold = true;
            //sheet.Range["d7:f7"].Style.Font.IsItalic = true;
            sheet.Range["d7:g7"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d7:g7"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d7:g7"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["d7:g7"].Style.Font.Size = 11;
            sheet.Range["d7:g7"].Merge(); // birlashtirish
            sheet.Range["d7:g7"].Text = "№ Требование: " + num_qrcode_textBox.Text;
            sheet.Range["d7:g7"].Style.WrapText = true;
            sheet.Range["d7:g7"].Style.Font.Underline = FontUnderlineType.SingleAccounting;

            //sheet.Range["g7:k7"].Style.Font.IsBold = true;
            //sheet.Range["g7:k7"].Style.Font.IsItalic = true;
            sheet.Range["h7:k7"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h7:k7"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h7:k7"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["h7:k7"].Style.Font.Size = 11;
            sheet.Range["h7:k7"].Merge(); // birlashtirish
            sheet.Range["h7:k7"].Text = "Дата: ";
            sheet.Range["h7:k7"].Style.WrapText = true;
            sheet.Range["h7:k7"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.SetRowHeight(7, 21);

            sheet.Range["a8:a8"].Style.Font.IsBold = true;
            sheet.Range["a8:a8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a8:a8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a8:a8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a8:a8"].Style.Font.Size = 11;
            sheet.Range["a8:a8"].Merge(); // birlashtirish
            sheet.Range["a8:a8"].Text = "№";
            sheet.Range["a8:a8"].Style.WrapText = true;
            sheet.Range["a8:a8"].BorderAround(LineStyleType.Thin);

            sheet.Range["b8:b8"].Style.Font.IsBold = true;
            sheet.Range["b8:b8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b8:b8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b8:b8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b8:b8"].Style.Font.Size = 11;
            sheet.Range["b8:b8"].Merge(); // birlashtirish
            sheet.Range["b8:b8"].Text = "Наименование";
            sheet.Range["b8:b8"].Style.WrapText = true;
            sheet.Range["b8:b8"].BorderAround(LineStyleType.Thin);

            sheet.Range["c8:c8"].Style.Font.IsBold = true;
            sheet.Range["c8:c8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c8:c8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c8:c8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c8:c8"].Style.Font.Size = 10;
            sheet.Range["c8:c8"].Merge(); // birlashtirish
            sheet.Range["c8:c8"].Text = "Ед.из";
            sheet.Range["c8:c8"].BorderAround(LineStyleType.Thin);
            sheet.Range["c8:c8"].Style.WrapText = true;

            sheet.Range["d8:d8"].Style.Font.IsBold = true;
            sheet.Range["d8:d8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d8:d8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d8:d8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d8:d8"].Style.Font.Size = 11;
            sheet.Range["d8:d8"].Merge(); // birlashtirish
            sheet.Range["d8:d8"].Text = "Кол.";
            sheet.Range["d8:d8"].Style.WrapText = true;
            sheet.Range["d8:d8"].BorderAround(LineStyleType.Thin);

            sheet.Range["e8:e8"].Style.Font.IsBold = true;
            sheet.Range["e8:e8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e8:e8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e8:e8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e8:e8"].Style.Font.Size = 11;
            sheet.Range["e8:e8"].Merge(); // birlashtirish
            sheet.Range["e8:e8"].Text = "Цена";
            sheet.Range["e8:e8"].Style.WrapText = true;
            sheet.Range["e8:e8"].BorderAround(LineStyleType.Thin);

            sheet.Range["f8:f8"].Style.Font.IsBold = true;
            sheet.Range["f8:f8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f8:f8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f8:f8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f8:f8"].Style.Font.Size = 11;
            sheet.Range["f8:f8"].Merge(); // birlashtirish
            sheet.Range["f8:f8"].Text = "Сумма";
            sheet.Range["f8:f8"].Style.WrapText = true;
            sheet.Range["f8:f8"].BorderAround(LineStyleType.Thin);

            sheet.Range["g8:h8"].Style.Font.IsBold = true;
            sheet.Range["g8:h8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g8:h8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g8:h8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g8:h8"].Style.Font.Size = 11;
            sheet.Range["g8:h8"].Merge(); // birlashtirish
            sheet.Range["g8:h8"].Text = "Дебет";
            sheet.Range["g8:h8"].Style.WrapText = true;
            sheet.Range["g8:h8"].BorderAround(LineStyleType.Thin);

            sheet.Range["i8:j8"].Style.Font.IsBold = true;
            sheet.Range["i8:j8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["i8:j8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["i8:j8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["i8:j8"].Style.Font.Size = 11;
            sheet.Range["i8:j8"].Merge(); // birlashtirish
            sheet.Range["i8:j8"].Text = "Кредит";
            sheet.Range["i8:j8"].Style.WrapText = true;
            sheet.Range["i8:j8"].BorderAround(LineStyleType.Thin);

            sheet.Range["k8:k8"].Style.Font.IsBold = true;
            sheet.Range["k8:k8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["k8:k8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["k8:k8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["k8:k8"].Style.Font.Size = 11;
            sheet.Range["k8:k8"].Merge(); // birlashtirish
            sheet.Range["k8:k8"].Text = "Дата";
            sheet.Range["k8:k8"].Style.WrapText = true;
            sheet.Range["k8:k8"].BorderAround(LineStyleType.Thin);


            int i = 0;
            int myrow = 9;
            int j = 0;

            var top = " SELECT * FROM products_jur7 where kod_doc='" + kod_num_textBox.Text + "' and user='" + string_for_otdels + "' and month='" + month_textBox.Text + "' and year='" + year_textBox.Text + "'  ";
            sql.myReader = sql.return_MySqlCommand(top).ExecuteReader();
            while (sql.myReader.Read())
            {
                j = i;
                j = j + 1;

                sheet.Range["a" + myrow + ":a" + myrow].Merge();
                sheet.Range["a" + myrow + ":a" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["a" + myrow + ":a" + myrow].Style.VerticalAlignment = VerticalAlignType.Bottom;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.Size = 10;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["a" + myrow + ":a" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["a" + myrow + ":a" + myrow].Text = (String)j.ToString();

                sheet.Range["b" + myrow + ":b" + myrow].Merge();
                sheet.Range["b" + myrow + ":b" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["b" + myrow + ":b" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["b" + myrow + ":b" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.Size = 10;
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["b" + myrow + ":b" + myrow].Text = sql.myReader["naim_tov"] != DBNull.Value ? sql.myReader.GetString("naim_tov").ToString() : "";
                sheet.Range["b" + myrow + ":b" + myrow].Style.WrapText = true;

                sheet.Range["c" + myrow + ":c" + myrow].Merge();
                sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 10;
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["c" + myrow + ":c" + myrow].Text = sql.myReader["edin"] != DBNull.Value ? sql.myReader.GetString("edin").ToString() : "";
                sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;

                sheet.Range["d" + myrow + ":d" + myrow].Merge();
                sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
                sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["d" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 10;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["d" + myrow + ":d" + myrow].Value = sql.myReader["kol"] != DBNull.Value ? sql.myReader.GetString("kol").ToString() : "";

                sheet.Range["e" + myrow + ":e" + myrow].Merge();
                sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 10;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["e" + myrow + ":e" + myrow].Value = sql.myReader["sena"] != DBNull.Value ? sql.myReader.GetString("sena").ToString() : "";

                sheet.Range["f" + myrow + ":f" + myrow].Merge();
                sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":f" + myrow].Value = sql.myReader["summa"] != DBNull.Value ? sql.myReader.GetString("summa").ToString() : "";

                sheet.Range["g" + myrow + ":g" + myrow].Merge();
                sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
                sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["g" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 10;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["g" + myrow + ":g" + myrow].Text = sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch").ToString() : "";

                sheet.Range["h" + myrow + ":h" + myrow].Merge();
                sheet.Range["h" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["h" + myrow + ":h" + myrow].Style.WrapText = true;
                sheet.Range["h" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["h" + myrow + ":h" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.Size = 10;
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["h" + myrow + ":h" + myrow].Text = sql.myReader["deb_sch_2"] != DBNull.Value ? sql.myReader.GetString("deb_sch_2").ToString() : "";

                sheet.Range["i" + myrow + ":i" + myrow].Merge();
                sheet.Range["i" + myrow + ":i" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["i" + myrow + ":i" + myrow].Style.WrapText = true;
                sheet.Range["i" + myrow + ":i" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["i" + myrow + ":i" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["i" + myrow + ":i" + myrow].Style.Font.Size = 10;
                sheet.Range["i" + myrow + ":i" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["i" + myrow + ":i" + myrow].Text = sql.myReader["kre_sch"] != DBNull.Value ? sql.myReader.GetString("kre_sch").ToString() : "";

                sheet.Range["j" + myrow + ":j" + myrow].Merge();
                sheet.Range["j" + myrow + ":j" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["j" + myrow + ":j" + myrow].Style.WrapText = true;
                sheet.Range["j" + myrow + ":j" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["j" + myrow + ":j" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["j" + myrow + ":j" + myrow].Style.Font.Size = 10;
                sheet.Range["j" + myrow + ":j" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["j" + myrow + ":j" + myrow].Text = sql.myReader["kre_sch_2"] != DBNull.Value ? sql.myReader.GetString("kre_sch_2").ToString() : "";

                sheet.Range["k" + myrow + ":k" + myrow].Merge();
                sheet.Range["k" + myrow + ":k" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["k" + myrow + ":k" + myrow].Style.WrapText = true;
                sheet.Range["k" + myrow + ":k" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["k" + myrow + ":k" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["k" + myrow + ":k" + myrow].Style.Font.Size = 10;
                sheet.Range["k" + myrow + ":k" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["k" + myrow + ":k" + myrow].Text = sql.myReader["date_pr"] != DBNull.Value ? Convert.ToDateTime(sql.myReader.GetString("date_pr").ToString()).ToString("dd.MM.yyyy") : "";


                myrow = myrow + 1;
                i = i + 1;


            }


            sheet.Range["e" + myrow + ":e" + myrow].Merge();
            sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.IsBold = true;
            sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 11;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["e" + myrow + ":e" + myrow].Text = "Всего:";
            sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);

            sheet.Range["f" + myrow + ":g" + myrow].Merge();
            sheet.Range["f" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["f" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.Size = 11;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["f" + myrow + ":g" + myrow].Value = qrcode_obshiy_summa_label.Text;
            sheet.Range["f" + myrow + ":g" + myrow].Style.WrapText = true;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.IsBold = true;
            sheet.Range["f" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);
            myrow++;

            sheet.Range["a" + myrow + ":h" + myrow].Merge();
            sheet.Range["a" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a" + myrow + ":h" + myrow].Style.WrapText = true;
            //sheet.Rangea"b" + myrow + ":h" + myrow].Style.Font.IsBold = true;
            sheet.Range["a" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.Size = 11;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["a" + myrow + ":h" + myrow].Text = "    Приложение";

            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.SetRowHeight(myrow, 18);
            myrow++;

            sheet.Range["a" + myrow + ":c" + myrow].Merge();
            sheet.Range["a" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a" + myrow + ":c" + myrow].Style.WrapText = true;
            sheet.Range["a" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["a" + myrow + ":c" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["a" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["a" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["a" + myrow + ":c" + myrow].Text = "Началник отдела";

            sheet.Range["d" + myrow + ":h" + myrow].Merge();
            sheet.Range["d" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["d" + myrow + ":h" + myrow].Style.WrapText = true;
            sheet.Range["d" + myrow + ":h" + myrow].Style.Font.IsBold = true;
            sheet.Range["d" + myrow + ":h" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["d" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d" + myrow + ":h" + myrow].Style.Font.Size = 11;
            sheet.Range["d" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["d" + myrow + ":h" + myrow].Text = "Бухгалтер";



            sheet.SetRowHeight(myrow, 18);

            myrow++;

            sheet.Range["a" + myrow + ":h" + myrow].Merge();
            sheet.Range["a" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a" + myrow + ":h" + myrow].Style.WrapText = true;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.IsBold = true;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["a" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.Size = 11;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["a" + myrow + ":h" + myrow].Text = " ";

            sheet.SetRowHeight(myrow, 18);

            myrow++;

            sheet.Range["a" + myrow + ":h" + myrow].Merge();
            sheet.Range["a" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a" + myrow + ":h" + myrow].Style.WrapText = true;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.IsBold = true;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.IsItalic = true;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["a" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.Size = 11;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["a" + myrow + ":h" + myrow].Text = "Линия отреза";

            sheet.SetRowHeight(myrow, 18);

            myrow++;

            sheet.Range["a" + myrow + ":h" + myrow].Merge();
            sheet.Range["a" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a" + myrow + ":h" + myrow].Style.WrapText = true;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.IsBold = true;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["a" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.Size = 11;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["a" + myrow + ":h" + myrow].Text = textBox4.Text;

            sheet.SetRowHeight(myrow, 18);

            myrow++;

            sheet.Range["a" + myrow + ":b" + myrow].Merge();
            sheet.Range["a" + myrow + ":b" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a" + myrow + ":b" + myrow].Style.WrapText = true;
            sheet.Range["a" + myrow + ":b" + myrow].Style.Font.IsBold = true;
            sheet.Range["a" + myrow + ":b" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["a" + myrow + ":b" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a" + myrow + ":b" + myrow].Style.Font.Size = 11;
            sheet.Range["a" + myrow + ":b" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["a" + myrow + ":b" + myrow].Text = " ";

            sheet.SetRowHeight(myrow, 18);

            myrow++;

            sheet.Range["a" + myrow + ":b" + myrow].Merge();
            sheet.Range["a" + myrow + ":b" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a" + myrow + ":b" + myrow].Style.WrapText = true;
            sheet.Range["a" + myrow + ":b" + myrow].Style.Font.IsBold = true;
            sheet.Range["a" + myrow + ":b" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["a" + myrow + ":b" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a" + myrow + ":b" + myrow].Style.Font.Size = 11;
            sheet.Range["a" + myrow + ":b" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["a" + myrow + ":b" + myrow].Text = "№";

            sheet.Range["c" + myrow + ":h" + myrow].Merge();
            sheet.Range["c" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["c" + myrow + ":h" + myrow].Style.WrapText = true;
            sheet.Range["c" + myrow + ":h" + myrow].Style.Font.IsBold = true;
            //sheet.Range["c" + myrow + ":h" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["c" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c" + myrow + ":h" + myrow].Style.Font.Size = 11;
            sheet.Range["c" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["c" + myrow + ":h" + myrow].Text = "Подтверждение к извещению №";

            sheet.SetRowHeight(myrow, 18);

            myrow++;


            sheet.Range["b" + myrow + ":h" + myrow].Merge();
            sheet.Range["b" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":h" + myrow].Style.WrapText = true;
            //sheet.Range["b" + myrow + ":h" + myrow].Style.Font.IsBold = true;
            //sheet.Range["c" + myrow + ":h" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["b" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":h" + myrow].Style.Font.Size = 10;
            sheet.Range["b" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":h" + myrow].Text = "Перечисленные в извещении материалъные ценности получены и взяты на балансовый";

            sheet.SetRowHeight(myrow, 18);

            myrow++;

            sheet.Range["a" + myrow + ":h" + myrow].Merge();
            sheet.Range["a" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a" + myrow + ":h" + myrow].Style.WrapText = true;
            //sheet.Range["a" + myrow + ":h" + myrow].Style.Font.IsBold = true;
            //sheet.Range["c" + myrow + ":h" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["a" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.Size = 10;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["a" + myrow + ":h" + myrow].Text = "учет в___________квартале 20___ г. в сум______________________________________";

            sheet.SetRowHeight(myrow, 18);

            myrow++;

            sheet.Range["a" + myrow + ":h" + myrow].Merge();
            sheet.Range["a" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a" + myrow + ":h" + myrow].Style.WrapText = true;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.IsBold = true;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["a" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.Size = 11;
            sheet.Range["a" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["a" + myrow + ":h" + myrow].Text = " ";

            sheet.SetRowHeight(myrow, 18);

            myrow++;

            sheet.Range["b" + myrow + ":b" + myrow].Merge();
            sheet.Range["b" + myrow + ":b" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b" + myrow + ":b" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":b" + myrow].Style.Font.IsBold = true;
            //sheet.Range["c" + myrow + ":h" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["b" + myrow + ":b" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":b" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":b" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":b" + myrow].Text = "Началъник ФЕО";


            sheet.Range["e" + myrow + ":f" + myrow].Merge();
            sheet.Range["e" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e" + myrow + ":f" + myrow].Style.WrapText = true;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Font.IsBold = true;
            //sheet.Range["c" + myrow + ":h" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["e" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Font.Size = 11;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["e" + myrow + ":f" + myrow].Text = "Ст.бухгалтер";

            sheet.Range["d5:" + myrow + "k"].NumberFormat = "#,##0.00";


            workbook.SaveToFile(Environment.CurrentDirectory + "\\docs\\Извещение.xlsx", Spire.Xls.FileFormat.Version2007);
            System.Diagnostics.Process.Start(workbook.FileName);



            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Извещение_excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void vnut_per_spisanie_btn_Click(object sender, EventArgs e)
        {
            //try
            //{
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.LeftMargin = 0.2;
            sheet.PageSetup.RightMargin = 0.2;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;


            sheet.PageSetup.Orientation = PageOrientationType.Portrait;


            sheet.Range["a1:a1"].ColumnWidth = 3;
            sheet.Range["b1:b1"].ColumnWidth = 35.57;
            sheet.Range["c1:c1"].ColumnWidth = 4.86;
            sheet.Range["d1:d1"].ColumnWidth = 9;
            sheet.Range["e1:e1"].ColumnWidth = 10;
            sheet.Range["f1:f1"].ColumnWidth = 13;
            sheet.Range["g1:g1"].ColumnWidth = 4;
            sheet.Range["h1:h1"].ColumnWidth = 7;
            sheet.Range["i1:i1"].ColumnWidth = 4;
            sheet.Range["j1:j1"].ColumnWidth = 7;


            string name_organ = "";
            var name_org = "SELECT * FROM spravochnik_main_jur7 where user_jur7='" + string_for_otdels + "'";
            sql.myReader.Close();
            sql.myReader = sql.return_MySqlCommand(name_org).ExecuteReader();
            while (sql.myReader.Read())
            {
                name_organ = (sql.myReader["naim_org"] != DBNull.Value ? sql.myReader.GetString("naim_org") : "");
            }
            sql.myReader.Close();

            sheet.Range["a1:e1"].Style.Font.IsBold = true;
            sheet.Range["a1:e1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a1:e1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a1:e1"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a1:e1"].Style.Font.Size = 14;
            sheet.Range["a1:e1"].Merge(); // birlashtirish
            sheet.Range["a1:e1"].Text = name_organ;
            sheet.Range["a1:e1"].Style.Font.Color = Color.DarkBlue;

            sheet.Range["f1:j1"].Style.Font.IsBold = true;
            sheet.Range["f1:j1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f1:j1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f1:j1"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f1:j1"].Style.Font.Size = 14;
            sheet.Range["f1:j1"].Merge(); // birlashtirish
            sheet.Range["f1:j1"].Text = "\"УТВЕРЖДАЮ\"";
            sheet.Range["f1:j1"].Style.Font.Color = Color.DarkBlue;

            sheet.SetRowHeight(1, 21);


            //sheet.Range["f2:j2"].Style.Font.IsBold = true;
            sheet.Range["f2:j2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f2:j2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f2:j2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f2:j2"].Style.Font.Size = 14;
            sheet.Range["f2:j2"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["f2:j2"].Merge(); // birlashtirish
            sheet.Range["f2:j2"].Text = " ";
            sheet.Range["f2:j2"].Style.Font.Color = Color.DarkBlue;
            sheet.SetRowHeight(2, 18);

            //sheet.Range["f2:j2"].Style.Font.IsBold = true;
            sheet.Range["f3:j3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f3:j3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f3:j3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f3:j3"].Style.Font.Size = 12;
            sheet.Range["f3:j3"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["f3:j3"].Merge(); // birlashtirish
            sheet.Range["f3:j3"].Text = " ";
            //sheet.Range["f3:j3"].Style.Font.Color = Color.DarkBlue;
            sheet.SetRowHeight(3, 18);

            //sheet.Range["f2:j2"].Style.Font.IsBold = true;
            sheet.Range["f4:j4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f4:j4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f4:j4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f4:j4"].Style.Font.Size = 14;
            sheet.Range["f4:j4"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["f4:j4"].Merge(); // birlashtirish
            sheet.Range["f4:j4"].Text = " ";
            sheet.Range["f4:j4"].Style.Font.Color = Color.DarkBlue;
            sheet.SetRowHeight(4, 18);

            sheet.Range["a5:j5"].Style.Font.IsBold = true;
            sheet.Range["a5:j5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a5:j5"].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["a5:j5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a5:j5"].Style.Font.Size = 12;
            sheet.Range["a5:j5"].Merge(); // birlashtirish
            sheet.Range["a5:j5"].Text = "Прием передача № " + num_qrcode_textBox.Text + " от " + data_qrcode_DateTimePicker.Value.ToString("dd.MM.yyyy");
            sheet.Range["a5:j5"].Style.WrapText = true;
            //sheet.Range["a5:j5"].BorderAround(LineStyleType.Thin);

            sheet.SetRowHeight(5, 24);

            sheet.Range["b6:j6"].Style.Font.IsBold = true;
            sheet.Range["b6:j6"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b6:j6"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b6:j6"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b6:j6"].Style.Font.Size = 11;
            sheet.Range["b6:j6"].Merge(); // birlashtirish
            sheet.Range["b6:j6"].Text = "Выдатъ(откуда) :" + textBox1.Text;
            sheet.Range["b6:j6"].Style.WrapText = true;
            //sheet.Range["a6:j6"].BorderAround(LineStyleType.Thin);

            sheet.SetRowHeight(6, 18);

            sheet.Range["b7:j7"].Style.Font.IsBold = true;
            sheet.Range["b7:j7"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b7:j7"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b7:j7"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b7:j7"].Style.Font.Size = 11;
            sheet.Range["b7:j7"].Merge(); // birlashtirish
            sheet.Range["b7:j7"].Text = "Кому:" + textBox4.Text + " " + textBox3.Text;
            sheet.Range["b7:j7"].Style.WrapText = true;
            //sheet.Range["a7:j7"].BorderAround(LineStyleType.Thin);

            sheet.SetRowHeight(7, 20);

            sheet.Range["a8:a8"].Style.Font.IsBold = true;
            sheet.Range["a8:a8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a8:a8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a8:a8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a8:a8"].Style.Font.Size = 11;
            sheet.Range["a8:a8"].Merge(); // birlashtirish
            sheet.Range["a8:a8"].Text = "№";
            sheet.Range["a8:a8"].Style.WrapText = true;
            sheet.Range["a8:a8"].BorderAround(LineStyleType.Thin);

            sheet.Range["b8:b8"].Style.Font.IsBold = true;
            sheet.Range["b8:b8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b8:b8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b8:b8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b8:b8"].Style.Font.Size = 11;
            sheet.Range["b8:b8"].Merge(); // birlashtirish
            sheet.Range["b8:b8"].Text = "Наименование";
            sheet.Range["b8:b8"].Style.WrapText = true;
            sheet.Range["b8:b8"].BorderAround(LineStyleType.Thin);

            sheet.Range["c8:c8"].Style.Font.IsBold = true;
            sheet.Range["c8:c8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c8:c8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c8:c8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c8:c8"].Style.Font.Size = 10;
            sheet.Range["c8:c8"].Merge(); // birlashtirish
            sheet.Range["c8:c8"].Text = "Ед.из";
            sheet.Range["c8:c8"].BorderAround(LineStyleType.Thin);
            sheet.Range["c8:c8"].Style.WrapText = true;

            sheet.Range["d8:d8"].Style.Font.IsBold = true;
            sheet.Range["d8:d8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d8:d8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d8:d8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d8:d8"].Style.Font.Size = 11;
            sheet.Range["d8:d8"].Merge(); // birlashtirish
            sheet.Range["d8:d8"].Text = "Кол.";
            sheet.Range["d8:d8"].Style.WrapText = true;
            sheet.Range["d8:d8"].BorderAround(LineStyleType.Thin);

            sheet.Range["e8:e8"].Style.Font.IsBold = true;
            sheet.Range["e8:e8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e8:e8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e8:e8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e8:e8"].Style.Font.Size = 11;
            sheet.Range["e8:e8"].Merge(); // birlashtirish
            sheet.Range["e8:e8"].Text = "Цена";
            sheet.Range["e8:e8"].Style.WrapText = true;
            sheet.Range["e8:e8"].BorderAround(LineStyleType.Thin);

            sheet.Range["f8:f8"].Style.Font.IsBold = true;
            sheet.Range["f8:f8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f8:f8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f8:f8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f8:f8"].Style.Font.Size = 11;
            sheet.Range["f8:f8"].Merge(); // birlashtirish
            sheet.Range["f8:f8"].Text = "Сумма";
            sheet.Range["f8:f8"].Style.WrapText = true;
            sheet.Range["f8:f8"].BorderAround(LineStyleType.Thin);

            sheet.Range["g8:h8"].Style.Font.IsBold = true;
            sheet.Range["g8:h8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g8:h8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g8:h8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g8:h8"].Style.Font.Size = 11;
            sheet.Range["g8:h8"].Merge(); // birlashtirish
            sheet.Range["g8:h8"].Text = "Дебет";
            sheet.Range["g8:h8"].Style.WrapText = true;
            sheet.Range["g8:h8"].BorderAround(LineStyleType.Thin);

            sheet.Range["i8:j8"].Style.Font.IsBold = true;
            sheet.Range["i8:j8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["i8:j8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["i8:j8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["i8:j8"].Style.Font.Size = 11;
            sheet.Range["i8:j8"].Merge(); // birlashtirish
            sheet.Range["i8:j8"].Text = "Кредит";
            sheet.Range["i8:j8"].Style.WrapText = true;
            sheet.Range["i8:j8"].BorderAround(LineStyleType.Thin);



            int i = 0;
            int myrow = 9;
            int j = 0;

            var top = " SELECT * FROM products_jur7 where kod_doc='" + kod_num_textBox.Text + "' and user='" + string_for_otdels + "' and month='" + month_textBox.Text + "' and year='" + year_textBox.Text + "'  ";
            sql.myReader = sql.return_MySqlCommand(top).ExecuteReader();
            while (sql.myReader.Read())
            {
                j = i;
                j = j + 1;

                sheet.Range["a" + myrow + ":a" + myrow].Merge();
                sheet.Range["a" + myrow + ":a" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["a" + myrow + ":a" + myrow].Style.VerticalAlignment = VerticalAlignType.Bottom;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.Size = 10;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["a" + myrow + ":a" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["a" + myrow + ":a" + myrow].Text = (String)j.ToString();

                sheet.Range["b" + myrow + ":b" + myrow].Merge();
                sheet.Range["b" + myrow + ":b" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["b" + myrow + ":b" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["b" + myrow + ":b" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.Size = 10;
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["b" + myrow + ":b" + myrow].Text = sql.myReader["naim_tov"] != DBNull.Value ? sql.myReader.GetString("naim_tov").ToString() : "";
                sheet.Range["b" + myrow + ":b" + myrow].Style.WrapText = true;

                sheet.Range["c" + myrow + ":c" + myrow].Merge();
                sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 10;
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["c" + myrow + ":c" + myrow].Text = sql.myReader["edin"] != DBNull.Value ? sql.myReader.GetString("edin").ToString() : "";
                sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;

                sheet.Range["d" + myrow + ":d" + myrow].Merge();
                sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
                sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["d" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 10;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["d" + myrow + ":d" + myrow].Value = sql.myReader["kol"] != DBNull.Value ? sql.myReader.GetString("kol").ToString() : "";

                sheet.Range["e" + myrow + ":e" + myrow].Merge();
                sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 10;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["e" + myrow + ":e" + myrow].Value = sql.myReader["sena"] != DBNull.Value ? sql.myReader.GetString("sena").ToString() : "";

                sheet.Range["f" + myrow + ":f" + myrow].Merge();
                sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":f" + myrow].Value = sql.myReader["summa"] != DBNull.Value ? sql.myReader.GetString("summa").ToString() : "";

                sheet.Range["g" + myrow + ":g" + myrow].Merge();
                sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
                sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["g" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 10;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["g" + myrow + ":g" + myrow].Text = sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch").ToString() : "";

                sheet.Range["h" + myrow + ":h" + myrow].Merge();
                sheet.Range["h" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["h" + myrow + ":h" + myrow].Style.WrapText = true;
                sheet.Range["h" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["h" + myrow + ":h" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.Size = 10;
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["h" + myrow + ":h" + myrow].Text = sql.myReader["deb_sch_2"] != DBNull.Value ? sql.myReader.GetString("deb_sch_2").ToString() : "";

                sheet.Range["i" + myrow + ":i" + myrow].Merge();
                sheet.Range["i" + myrow + ":i" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["i" + myrow + ":i" + myrow].Style.WrapText = true;
                sheet.Range["i" + myrow + ":i" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["i" + myrow + ":i" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["i" + myrow + ":i" + myrow].Style.Font.Size = 10;
                sheet.Range["i" + myrow + ":i" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["i" + myrow + ":i" + myrow].Text = sql.myReader["kre_sch"] != DBNull.Value ? sql.myReader.GetString("kre_sch").ToString() : "";

                sheet.Range["j" + myrow + ":j" + myrow].Merge();
                sheet.Range["j" + myrow + ":j" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["j" + myrow + ":j" + myrow].Style.WrapText = true;
                sheet.Range["j" + myrow + ":j" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["j" + myrow + ":j" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["j" + myrow + ":j" + myrow].Style.Font.Size = 10;
                sheet.Range["j" + myrow + ":j" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["j" + myrow + ":j" + myrow].Text = sql.myReader["kre_sch_2"] != DBNull.Value ? sql.myReader.GetString("kre_sch_2").ToString() : "";


                myrow = myrow + 1;
                i = i + 1;

            }
            sql.myReader.Close();



            sheet.Range["e" + myrow + ":e" + myrow].Merge();
            sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.IsBold = true;
            sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 11;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["e" + myrow + ":e" + myrow].Text = "Всего:";
            sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);

            sheet.Range["f" + myrow + ":f" + myrow].Merge();
            sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["f" + myrow + ":f" + myrow].Value = qrcode_obshiy_summa_label.Text;
            sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.IsBold = true;
            sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
            myrow++;
            myrow++;

            sheet.Range["b" + myrow + ":c" + myrow].Merge();
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Нач.ФЕО: ";

            sheet.Range["d" + myrow + ":e" + myrow].Merge();
            sheet.Range["d" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["d" + myrow + ":e" + myrow].Style.WrapText = true;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.IsBold = true;
            sheet.Range["d" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.Size = 11;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["d" + myrow + ":e" + myrow].Text = "Члены комиссии";

            sheet.Range["f" + myrow + ":i" + myrow].Merge();
            sheet.Range["f" + myrow + ":i" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["f" + myrow + ":i" + myrow].Style.WrapText = true;
            sheet.Range["f" + myrow + ":i" + myrow].Style.Font.IsBold = true;
            sheet.Range["f" + myrow + ":i" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["f" + myrow + ":i" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f" + myrow + ":i" + myrow].Style.Font.Size = 11;
            sheet.Range["f" + myrow + ":i" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["f" + myrow + ":i" + myrow].Text = " ";

            sheet.SetRowHeight(myrow, 20);
            myrow++;

            sheet.Range["b" + myrow + ":c" + myrow].Merge();
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Нач.Отдел: ";

            sheet.Range["f" + myrow + ":i" + myrow].Merge();
            sheet.Range["f" + myrow + ":i" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["f" + myrow + ":i" + myrow].Style.WrapText = true;
            sheet.Range["f" + myrow + ":i" + myrow].Style.Font.IsBold = true;
            sheet.Range["f" + myrow + ":i" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["f" + myrow + ":i" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f" + myrow + ":i" + myrow].Style.Font.Size = 11;
            sheet.Range["f" + myrow + ":i" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["f" + myrow + ":i" + myrow].Text = " ";
            sheet.SetRowHeight(myrow, 18);

            myrow++;

            sheet.Range["b" + myrow + ":c" + myrow].Merge();
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Гл.Спец: ";

            sheet.Range["f" + myrow + ":i" + myrow].Merge();
            sheet.Range["f" + myrow + ":i" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["f" + myrow + ":i" + myrow].Style.WrapText = true;
            sheet.Range["f" + myrow + ":i" + myrow].Style.Font.IsBold = true;
            sheet.Range["f" + myrow + ":i" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["f" + myrow + ":i" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f" + myrow + ":i" + myrow].Style.Font.Size = 11;
            sheet.Range["f" + myrow + ":i" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["f" + myrow + ":i" + myrow].Text = " ";


            sheet.SetRowHeight(myrow, 18);

            myrow++;

            sheet.Range["b" + myrow + ":c" + myrow].Merge();
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Получил:___________________________";

            sheet.Range["d" + myrow + ":e" + myrow].Merge();
            sheet.Range["d" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["d" + myrow + ":e" + myrow].Style.WrapText = true;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.IsBold = true;
            sheet.Range["d" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.Size = 11;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["d" + myrow + ":e" + myrow].Text = "Отпустил:";

            sheet.Range["f" + myrow + ":i" + myrow].Merge();
            sheet.Range["f" + myrow + ":i" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["f" + myrow + ":i" + myrow].Style.WrapText = true;
            sheet.Range["f" + myrow + ":i" + myrow].Style.Font.IsBold = true;
            sheet.Range["f" + myrow + ":i" + myrow].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.Range["f" + myrow + ":i" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f" + myrow + ":i" + myrow].Style.Font.Size = 11;
            sheet.Range["f" + myrow + ":i" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["f" + myrow + ":i" + myrow].Text = " ";


            sheet.SetRowHeight(myrow, 18);
            sheet.Range["d5:" + myrow + "k"].NumberFormat = "#,##0.00";


            workbook.SaveToFile(Environment.CurrentDirectory + "\\docs\\Списание.xlsx", Spire.Xls.FileFormat.Version2007);
            System.Diagnostics.Process.Start(workbook.FileName);



            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Списание_excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        int[] b = new int[2];
        private void vnut_per_naklad_btn_Click(object sender, EventArgs e)
        {
            //try
            //{
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.LeftMargin = 0.2;
            sheet.PageSetup.RightMargin = 0.2;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;


            sheet.PageSetup.Orientation = PageOrientationType.Portrait;


            sheet.Range["a1:a1"].ColumnWidth = 4;
            sheet.Range["b1:b1"].ColumnWidth = 35.14;
            sheet.Range["c1:c1"].ColumnWidth = 4.57;
            sheet.Range["d1:d1"].ColumnWidth = 8;
            sheet.Range["e1:e1"].ColumnWidth = 10;
            sheet.Range["f1:f1"].ColumnWidth = 13;
            sheet.Range["g1:g1"].ColumnWidth = 3;
            sheet.Range["h1:h1"].ColumnWidth = 3;
            sheet.Range["i1:i1"].ColumnWidth = 3;
            sheet.Range["j1:j1"].ColumnWidth = 3;
            sheet.Range["k1:k1"].ColumnWidth = 10;



            sheet.Range["a1:e1"].Style.Font.IsBold = true;
            sheet.Range["a1:e1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a1:e1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a1:e1"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["a1:e1"].Style.Font.Size = 14;
            sheet.Range["a1:e1"].Merge(); // birlashtirish
            sheet.Range["a1:e1"].Text = "СЧЕТ-ФАКТУРА-НАКЛАДНАЯ";
            // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);
            sheet.SetRowHeight(1, 20);

            sheet.Range["a2:e2"].Style.Font.IsBold = false;
            sheet.Range["a2:e2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a2:e2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a2:e2"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["a2:e2"].Style.Font.Size = 12;
            sheet.Range["a2:e2"].Merge(); // birlashtirish
            sheet.Range["a2:e2"].Text = "№ " + num_qrcode_textBox.Text + "от " + data_qrcode_DateTimePicker.Value.ToString("dd.MM.yyyy") + "         ";
            sheet.SetRowHeight(2, 20);

            sheet.Range["b3:e3"].Style.Font.IsBold = false;
            sheet.Range["b3:e3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b3:e3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b3:e3"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["b3:e3"].Style.Font.Size = 12;
            sheet.Range["b3:e3"].Merge(); // birlashtirish
            sheet.Range["b3:e3"].Text = "к товаро-отгрузочным документам № ";
            sheet.SetRowHeight(3, 20);

            sheet.Range["f3:f3"].Style.Font.IsBold = false;
            sheet.Range["f3:f3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f3:f3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f3:f3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f3:f3"].Style.Font.Size = 12;
            sheet.Range["f3:f3"].Merge(); // birlashtirish
            sheet.Range["f3:f3"].Text = "";
            sheet.Range["f3:f3"].Style.Font.Underline = FontUnderlineType.SingleAccounting;

            sheet.Range["b4:j4"].Style.Font.IsBold = false;
            sheet.Range["b4:j4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b4:j4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b4:j4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b4:j4"].Style.Font.Size = 11;
            sheet.Range["b4:j4"].Merge(); // birlashtirish
            sheet.Range["b4:j4"].Text = "     ";
            sheet.Range["b4:j4"].Style.Font.Underline = FontUnderlineType.SingleAccounting;

            var send = "SELECT * FROM spravochnik_main_jur7 where user_jur7='" + string_for_otdels + "'";

            sql1.myReader = sql1.return_MySqlCommand(send).ExecuteReader();
            while (sql1.myReader.Read())
            {
                sheet.Range["a5:d5"].Style.Font.IsBold = false;
                sheet.Range["a5:d5"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a5:d5"].Style.VerticalAlignment = VerticalAlignType.Bottom;
                sheet.Range["a5:d5"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a5:d5"].Style.Font.Size = 11;
                sheet.Range["a5:d5"].Merge(); // birlashtirish
                sheet.Range["a5:d5"].Text = "Поставщик :" + (sql1.myReader["naim_org"] != DBNull.Value ? sql1.myReader.GetString("naim_org") : "");
                sheet.Range["a5:d5"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);

                //naim_org,adres,ras_s,bank,inn,okxn
                sheet.Range["a6:d6"].Style.Font.IsBold = false;
                sheet.Range["a6:d6"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a6:d6"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a6:d6"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a6:d6"].Style.Font.Size = 11;
                sheet.Range["a6:d6"].Merge(); // birlashtirish
                sheet.Range["a6:d6"].Text = "Адрес :" + (sql1.myReader["adres"] != DBNull.Value ? sql1.myReader.GetString("adres") : "");
                sheet.Range["a6:d6"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);



                sheet.Range["a7:d7"].Style.Font.IsBold = false;
                sheet.Range["a7:d7"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a7:d7"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a7:d7"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a7:d7"].Style.Font.Size = 11;
                sheet.Range["a7:d7"].Merge(); // birlashtirish
                sheet.Range["a7:d7"].Text = "Р/с :" + (sql1.myReader["ras_s"] != DBNull.Value ? sql1.myReader.GetString("ras_s") : "");
                sheet.Range["a7:d7"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);



                sheet.Range["a8:d8"].Style.Font.IsBold = false;
                sheet.Range["a8:d8"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a8:d8"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a8:d8"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a8:d8"].Style.Font.Size = 11;
                sheet.Range["a8:d8"].Merge(); // birlashtirish
                sheet.Range["a8:d8"].Text = "Банк : " + (sql1.myReader["bank"] != DBNull.Value ? sql1.myReader.GetString("bank") : "");
                sheet.Range["a8:d8"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);



                sheet.Range["a9:d9"].Style.Font.IsBold = false;
                sheet.Range["a9:d9"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a9:d9"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a9:d9"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a9:d9"].Style.Font.Size = 11;
                sheet.Range["a9:d9"].Merge(); // birlashtirish
                sheet.Range["a9:d9"].Text = "МФО :" + (sql1.myReader["mfo"] != DBNull.Value ? sql1.myReader.GetString("mfo") : "");
                sheet.Range["a9:d9"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);

                sheet.Range["a10:d10"].Style.Font.IsBold = false;
                sheet.Range["a10:d10"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a10:d10"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a10:d10"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a10:d10"].Style.Font.Size = 11;
                sheet.Range["a10:d10"].Merge(); // birlashtirish
                sheet.Range["a10:d10"].Text = "ИНН : " + (sql1.myReader["inn"] != DBNull.Value ? sql1.myReader.GetString("inn") : "") + "  " + "ОКЭТ :" + (sql1.myReader["okxn"] != DBNull.Value ? sql1.myReader.GetString("okxn") : "");
                sheet.Range["a10:d10"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);
            }
            sql1.myReader.Close();


            ///poluchatel
            /// 
            sheet.Range["e5:k5"].Style.Font.IsBold = false;
            sheet.Range["e5:k5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e5:k5"].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["e5:k5"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["e5:k5"].Style.Font.Size = 11;
            sheet.Range["e5:k5"].Merge(); // birlashtirish
            sheet.Range["e5:k5"].Text = "Получателъ :";// + postavshik_comboBox.Text;
            sheet.Range["e5:k5"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);


            sheet.Range["e6:k6"].Style.Font.IsBold = false;
            sheet.Range["e6:k6"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e6:k6"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e6:k6"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["e6:k6"].Style.Font.Size = 11;
            sheet.Range["e6:k6"].Merge(); // birlashtirish
            sheet.Range["e6:k6"].Text = "Адрес :";// + comboBox_adres_rasxod.Text;
            sheet.Range["e6:k6"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);

            sheet.Range["e7:k7"].Style.Font.IsBold = false;
            sheet.Range["e7:k7"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e7:k7"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e7:k7"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["e7:k7"].Style.Font.Size = 11;
            sheet.Range["e7:k7"].Merge(); // birlashtirish
            sheet.Range["e7:k7"].Text = "Р/с :";// + comboBox_pc_rasxod.Text;
            sheet.Range["e7:k7"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);

            sheet.Range["e8:k8"].Style.Font.IsBold = false;
            sheet.Range["e8:k8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e8:k8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e8:k8"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["e8:k8"].Style.Font.Size = 11;
            sheet.Range["e8:k8"].Merge(); // birlashtirish
            sheet.Range["e8:k8"].Text = "Банк : ";// + comboBox_bank_rasxod.Text;
            sheet.Range["e8:k8"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);

            sheet.Range["e9:k9"].Style.Font.IsBold = false;
            sheet.Range["e9:k9"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e9:k9"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e9:k9"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["e9:k9"].Style.Font.Size = 11;
            sheet.Range["e9:k9"].Merge(); // birlashtirish
            sheet.Range["e9:k9"].Text = "МФО :";// + comboBox_mfo_rasxod.Text;
            sheet.Range["e9:k9"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);

            sheet.Range["e10:k10"].Style.Font.IsBold = false;
            sheet.Range["e10:k10"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e10:k10"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e10:k10"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["e10:k10"].Style.Font.Size = 11;
            sheet.Range["e10:k10"].Merge(); // birlashtirish
            sheet.Range["e10:k10"].Text = "ИНН : ";// + comboBox_inn_rasxod.Text;
            sheet.Range["e10:k10"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);



            sheet.SetRowHeight(2, 21);
            sheet.SetRowHeight(3, 18);
            sheet.SetRowHeight(4, 20);
            sheet.SetRowHeight(5, 18);
            sheet.SetRowHeight(6, 18);
            sheet.SetRowHeight(7, 18);
            sheet.SetRowHeight(8, 18);
            sheet.SetRowHeight(9, 18);
            sheet.SetRowHeight(10, 18);

            sheet.Range["a11:a11"].Merge(); // birlashtirish
            sheet.Range["a11:a11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["a11:a11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["a11:a11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["a11:a11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["a11:a11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a11:a11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a11:a11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a11:a11"].Style.Font.Size = 11;
            sheet.Range["a11:a11"].Style.WrapText = true;
            sheet.Range["a11:a11"].Text = "№";

            sheet.Range["b11:b11"].Merge(); // birlashtirish
            sheet.Range["b11:b11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["b11:b11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["b11:b11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["b11:b11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["b11:b11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b11:b11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b11:b11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b11:b11"].Style.Font.Size = 11;
            sheet.Range["b11:b11"].Style.WrapText = true;
            sheet.Range["b11:b11"].Text = "Наименование товара (работ, услуг)";

            sheet.Range["c11:c11"].Merge(); // birlashtirish
            sheet.Range["c11:c11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["c11:c11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["c11:c11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["c11:c11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["c11:c11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c11:c11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c11:c11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c11:c11"].Style.Font.Size = 11;
            sheet.Range["c11:c11"].Style.WrapText = true;
            sheet.Range["c11:c11"].Text = "Едизм";

            sheet.Range["d11:d11"].Merge(); // birlashtirish
            sheet.Range["d11:d11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["d11:d11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["d11:d11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["d11:d11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["d11:d11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d11:d11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d11:d11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d11:d11"].Style.Font.Size = 11;
            sheet.Range["d11:d11"].Style.WrapText = true;
            sheet.Range["d11:d11"].Text = "Кол-во";

            sheet.Range["e11:e11"].Merge(); // birlashtirish
            sheet.Range["e11:e11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["e11:e11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["e11:e11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["e11:e11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["e11:e11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e11:e11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e11:e11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e11:e11"].Style.Font.Size = 11;
            sheet.Range["e11:e11"].Style.WrapText = true;
            sheet.Range["e11:e11"].Text = "Цена";

            sheet.Range["f11:f11"].Merge(); // birlashtirish
            sheet.Range["f11:f11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["f11:f11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["f11:f11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["f11:f11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["f11:f11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f11:f11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f11:f11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f11:f11"].Style.Font.Size = 11;
            sheet.Range["f11:f11"].Text = "Стомость поставка";
            sheet.Range["f11:f11"].Style.WrapText = true;

            sheet.Range["g11:h11"].Merge(); // birlashtirish
            sheet.Range["g11:h11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["g11:h11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["g11:h11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["g11:h11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["g11:h11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g11:h11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g11:h11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g11:h11"].Style.Font.Size = 11;
            sheet.Range["g11:h11"].Style.WrapText = true;
            sheet.Range["g11:h11"].Text = "Акцизный налог";

            sheet.Range["i11:j11"].Merge(); // birlashtirish
            sheet.Range["i11:j11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["i11:j11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["i11:j11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["i11:j11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["i11:j11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["i11:j11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["i11:j11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["i11:j11"].Style.Font.Size = 11;
            sheet.Range["i11:j11"].Style.WrapText = true;
            sheet.Range["i11:j11"].Text = "НДС";

            sheet.Range["k11:k11"].Merge(); // birlashtirish
            sheet.Range["k11:k11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["k11:k11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["k11:k11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["k11:k11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["k11:k11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["k11:k11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["k11:k11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["k11:k11"].Style.Font.Size = 11;
            sheet.Range["k11:k11"].Style.WrapText = true;
            sheet.Range["k11:k11"].Text = "Стоимость поставки с НДС";


            sheet.Range["a12:a12"].Merge(); // birlashtirish
            sheet.Range["a12:a12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["a12:a12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["a12:a12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["a12:a12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["a12:a12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a12:a12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a12:a12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a12:a12"].Style.Font.Size = 11;
            sheet.Range["a12:a12"].Style.WrapText = true;
            sheet.Range["a12:a12"].Text = "1";

            sheet.Range["b12:b12"].Merge(); // birlashtirish
            sheet.Range["b12:b12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["b12:b12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["b12:b12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["b12:b12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["b12:b12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b12:b12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b12:b12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b12:b12"].Style.Font.Size = 11;
            sheet.Range["b12:b12"].Style.WrapText = true;
            sheet.Range["b12:b12"].Text = "2";

            sheet.Range["c12:c12"].Merge(); // birlashtirish
            sheet.Range["c12:c12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["c12:c12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["c12:c12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["c12:c12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["c12:c12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c12:c12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c12:c12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c12:c12"].Style.Font.Size = 11;
            sheet.Range["c12:c12"].Style.WrapText = true;
            sheet.Range["c12:c12"].Text = "3";

            sheet.Range["d12:d12"].Merge(); // birlashtirish
            sheet.Range["d12:d12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["d12:d12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["d12:d12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["d12:d12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["d12:d12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d12:d12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d12:d12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d12:d12"].Style.Font.Size = 11;
            sheet.Range["d12:d12"].Style.WrapText = true;
            sheet.Range["d12:d12"].Text = "4";

            sheet.Range["e12:e12"].Merge(); // birlashtirish
            sheet.Range["e12:e12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["e12:e12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["e12:e12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["e12:e12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["e12:e12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e12:e12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e12:e12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e12:e12"].Style.Font.Size = 11;
            sheet.Range["e12:e12"].Style.WrapText = true;
            sheet.Range["e12:e12"].Text = "5";

            sheet.Range["f12:f12"].Merge(); // birlashtirish
            sheet.Range["f12:f12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["f12:f12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["f12:f12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["f12:f12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["f12:f12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f12:f12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f12:f12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f12:f12"].Style.Font.Size = 11;
            sheet.Range["f12:f12"].Text = "6";
            sheet.Range["f12:f12"].Style.WrapText = true;

            sheet.Range["g12:g12"].Merge(); // birlashtirish
            sheet.Range["g12:g12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["g12:g12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["g12:g12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["g12:g12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["g12:g12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g12:g12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g12:g12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g12:g12"].Style.Font.Size = 11;
            sheet.Range["g12:g12"].Style.WrapText = true;
            sheet.Range["g12:g12"].Text = "7";

            sheet.Range["h12:h12"].Merge(); // birlashtirish
            sheet.Range["h12:h12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["h12:h12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["h12:h12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["h12:h12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["h12:h12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h12:h12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h12:h12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["h12:h12"].Style.Font.Size = 11;
            sheet.Range["h12:h12"].Style.WrapText = true;
            sheet.Range["h12:h12"].Text = "8";

            sheet.Range["i12:i12"].Merge(); // birlashtirish
            sheet.Range["i12:i12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["i12:i12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["i12:i12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["i12:i12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["i12:i12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["i12:i12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["i12:i12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["i12:i12"].Style.Font.Size = 11;
            sheet.Range["i12:i12"].Style.WrapText = true;
            sheet.Range["i12:i12"].Text = "9";

            sheet.Range["j12:j12"].Merge(); // birlashtirish
            sheet.Range["j12:j12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["j12:j12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["j12:j12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["j12:j12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["j12:j12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["j12:j12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["j12:j12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["j12:j12"].Style.Font.Size = 11;
            sheet.Range["j12:j12"].Style.WrapText = true;
            sheet.Range["j12:j12"].Text = "10";

            sheet.Range["k12:k12"].Merge(); // birlashtirish
            sheet.Range["k12:k12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["k12:k12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["k12:k12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["k12:k12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["k12:k12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["k12:k12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["k12:k12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["k12:k12"].Style.Font.Size = 11;
            sheet.Range["k12:k12"].Style.WrapText = true;
            sheet.Range["k12:k12"].Text = "11";

            //////////////////////



            int i = 0;
            int myrow = 13;
            int j = 0;

            double all_summa = 0;

            var send2 = "SELECT * FROM products_jur7 where kod_doc='" + kod_num_textBox.Text + "' and user='" + string_for_otdels + "' and month='" + month_textBox.Text + "' and year='" + year_textBox.Text + "' ";

            sql2.myReader = sql2.return_MySqlCommand(send2).ExecuteReader();
            while (sql2.myReader.Read())
            {
                j = i;
                j = j + 1;

                sheet.Range["a" + myrow + ":a" + myrow].Merge();
                sheet.Range["a" + myrow + ":a" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["a" + myrow + ":a" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.Size = 10;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["a" + myrow + ":a" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["a" + myrow + ":a" + myrow].Text = (String)j.ToString();

                sheet.Range["b" + myrow + ":b" + myrow].Merge();
                sheet.Range["b" + myrow + ":b" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["b" + myrow + ":b" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["b" + myrow + ":b" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.Size = 10;
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["b" + myrow + ":b" + myrow].Text = sql2.myReader["naim_tov"] != DBNull.Value ? sql2.myReader.GetString("naim_tov") : "";
                sheet.Range["b" + myrow + ":b" + myrow].Style.WrapText = true;


                sheet.Range["c" + myrow + ":c" + myrow].Merge();
                sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 10;
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["c" + myrow + ":c" + myrow].Value = sql2.myReader["edin"] != DBNull.Value ? sql2.myReader.GetString("edin") : "";
                sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;

                // naim_tov,edin,kol,sena,summa
                sheet.Range["d" + myrow + ":d" + myrow].Merge();
                sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
                sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["d" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 10;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["d" + myrow + ":d" + myrow].Value = sql2.myReader["kol"] != DBNull.Value ? sql2.myReader.GetString("kol") : "";

                sheet.Range["e" + myrow + ":e" + myrow].Merge();
                sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 10;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["e" + myrow + ":e" + myrow].Value = sql2.myReader["sena"] != DBNull.Value ? sql2.myReader.GetString("sena") : "";

                all_summa += sql2.myReader["summa"] != DBNull.Value ? sql2.myReader.GetDouble("summa") : 0;

                sheet.Range["f" + myrow + ":f" + myrow].Merge();
                sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":f" + myrow].Value = sql2.myReader["summa"] != DBNull.Value ? sql2.myReader.GetString("summa") : "";

                sheet.Range["g" + myrow + ":k" + myrow].Merge();
                sheet.Range["g" + myrow + ":k" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["g" + myrow + ":k" + myrow].Style.WrapText = true;
                sheet.Range["g" + myrow + ":k" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["g" + myrow + ":k" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["g" + myrow + ":k" + myrow].Style.Font.Size = 10;
                sheet.Range["g" + myrow + ":k" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["g" + myrow + ":k" + myrow].Text = "Без акц.налог Без НДС";

                myrow = myrow + 1;
                i = i + 1;

            }
            sql2.myReader.Close();


            //sheet.Range["a" + myrow + ":k" + myrow].Merge(); // birlashtirish
            //sheet.Range["a" + myrow + ":k" + myrow].Style.Font.FontName = "Times New Roman";
            //sheet.Range["a" + myrow + ":k" + myrow].Style.VerticalAlignment = VerticalAlignType.Bottom;
            //sheet.Range["a" + myrow + ":k" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            //sheet.Range["a" + myrow + ":k" + myrow].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            //sheet.Range["a" + myrow + ":k" + myrow].Style.Font.Size = 11;
            //sheet.Range["a" + myrow + ":k" + myrow].Text = "Без акц.налог Без НДС   ";
            //sheet.SetRowHeight(myrow, 18);
            //myrow++;

            sheet.Range["b" + myrow + ":d" + myrow].Merge(); // birlashtirish
            //sheet.Range["b" + myrow + ":d" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":d" + myrow].Text = "Итого : ";

            sheet.Range["e" + myrow + ":f" + myrow].Merge(); // birlashtirish
            //sheet.Range["e" + myrow + ":f" + myrow].Style.Font.IsBold = true;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["e" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Font.Size = 11;
            sheet.Range["e" + myrow + ":f" + myrow].Value = all_summa.ToString();

            sheet.SetRowHeight(myrow, 18);
            myrow++;

            String[] arr = refresh_strings_to_mysql(all_summa.ToString()).Split('.');
            b[0] = Convert.ToInt32(arr[0]);
            b[1] = Convert.ToInt32(arr[1]);

            sheet.Range["b" + myrow + ":k" + myrow].Style.Font.IsBold = true;
            //sheet.Range["b" + myrow + ":k" + myrow].Style.Font.IsItalic = true;
            sheet.Range["b" + myrow + ":k" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":k" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":k" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":k" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":k" + myrow].Merge(); // birlashtirish
            sheet.Range["b" + myrow + ":k" + myrow].Text = "Прописью :" + number_russian.toWords(b[0]) + " сум " + b[1] + " тийин";
            sheet.Range["b" + myrow + ":k" + myrow].Style.Font.Underline = FontUnderlineType.Single;

            sheet.SetRowHeight(myrow, 18);
            myrow++;

            //sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Merge(); // birlashtirish
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Нач. ФЭО: ";
            //  sheet.Range["b" + myrow + ":b" + myrow].Style.Font.Underline = FontUnderlineType.Single;

            myrow++;
            //sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Merge(); // birlashtirish
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Нач.Отдел";

            //sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["e" + myrow + ":j" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["e" + myrow + ":j" + myrow].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["e" + myrow + ":j" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["e" + myrow + ":j" + myrow].Style.Font.Size = 11;
            sheet.Range["e" + myrow + ":j" + myrow].Merge(); // birlashtirish
            sheet.Range["e" + myrow + ":j" + myrow].Text = "Получил :_________________________________";

            sheet.SetRowHeight(myrow, 18);
            myrow++;

            //sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Merge(); // birlashtirish
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Гл.Спец";

            sheet.SetRowHeight(myrow, 18);
            myrow++;

            //sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Merge(); // birlashtirish
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Отпустил : _________________________________";
            sheet.SetRowHeight(myrow, 18);

            sheet.Range["d13:" + myrow + "f"].NumberFormat = "#,##0.00";

            string kod_num = kod_num_textBox.Text;

            string filePath = Environment.CurrentDirectory + "\\docs\\qrcode\\" + string_for_otdels + "-" + year_textBox.Text + "-" + month_textBox.Text + "-" + kod_num + "-" + "1" + ".png";
            FileInfo file = new FileInfo(filePath);
            if (file.Exists)
            {
                string picPath = Environment.CurrentDirectory + "\\docs\\qrcode\\" + string_for_otdels + "-" + year_textBox.Text + "-" + month_textBox.Text + "-" + kod_num + "-" + "1" + ".png";
                ExcelPicture picture = sheet.Pictures.Add(1, 8, picPath);
                picture.Width = 60;
                picture.Height = 60;

            }
            else
            {
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode((string_for_otdels + "-" + year_textBox.Text + "-" + month_textBox.Text + "-" + kod_num + "-" + "1"), QRCodeGenerator.ECCLevel.Q);
                QRCode qrCode = new QRCode(qrCodeData);

                Bitmap bitMap = qrCode.GetGraphic(20);

                bitMap.Save(Environment.CurrentDirectory + "\\docs\\qrcode\\" + string_for_otdels + "-" + year_textBox.Text + "-" + month_textBox.Text + "-" + kod_num + "-" + "1" + ".png");

                string picPath = Environment.CurrentDirectory + "\\docs\\qrcode\\" + string_for_otdels + "-" + year_textBox.Text + "-" + month_textBox.Text + "-" + kod_num + "-" + "1" + ".png";
                ExcelPicture picture = sheet.Pictures.Add(1, 8, picPath);
                picture.Width = 60;
                picture.Height = 60;


            }

            workbook.SaveToFile(Environment.CurrentDirectory + "\\docs\\Накладная.xlsx", Spire.Xls.FileFormat.Version2007);
            System.Diagnostics.Process.Start(workbook.FileName);


            //  }
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Накладная_Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void vnut_per_schet_fac_btn_Click(object sender, EventArgs e)
        {
            //try
            //{
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.LeftMargin = 0.2;
            sheet.PageSetup.RightMargin = 0.2;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;


            sheet.PageSetup.Orientation = PageOrientationType.Portrait;


            sheet.Range["a1:a1"].ColumnWidth = 4;
            sheet.Range["b1:b1"].ColumnWidth = 35.14;
            sheet.Range["c1:c1"].ColumnWidth = 4.57;
            sheet.Range["d1:d1"].ColumnWidth = 8;
            sheet.Range["e1:e1"].ColumnWidth = 10;
            sheet.Range["f1:f1"].ColumnWidth = 13;
            sheet.Range["g1:g1"].ColumnWidth = 3;
            sheet.Range["h1:h1"].ColumnWidth = 3;
            sheet.Range["i1:i1"].ColumnWidth = 3;
            sheet.Range["j1:j1"].ColumnWidth = 3;
            sheet.Range["k1:k1"].ColumnWidth = 10;



            sheet.Range["a1:e1"].Style.Font.IsBold = true;
            sheet.Range["a1:e1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a1:e1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a1:e1"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["a1:e1"].Style.Font.Size = 14;
            sheet.Range["a1:e1"].Merge(); // birlashtirish
            sheet.Range["a1:e1"].Text = "СЧЕТ-ФАКТУРА-НАКЛАДНАЯ";
            // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);
            sheet.SetRowHeight(1, 20);

            sheet.Range["a2:e2"].Style.Font.IsBold = false;
            sheet.Range["a2:e2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a2:e2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a2:e2"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["a2:e2"].Style.Font.Size = 12;
            sheet.Range["a2:e2"].Merge(); // birlashtirish
            sheet.Range["a2:e2"].Text = "№ " + num_qrcode_textBox.Text + "от " + data_qrcode_DateTimePicker.Value.ToString("dd.MM.yyyy") + "         ";
            sheet.SetRowHeight(2, 20);

            sheet.Range["b3:e3"].Style.Font.IsBold = false;
            sheet.Range["b3:e3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b3:e3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b3:e3"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["b3:e3"].Style.Font.Size = 12;
            sheet.Range["b3:e3"].Merge(); // birlashtirish
            sheet.Range["b3:e3"].Text = "к товаро-отгрузочным документам № ";
            sheet.SetRowHeight(3, 20);

            sheet.Range["f3:f3"].Style.Font.IsBold = false;
            sheet.Range["f3:f3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f3:f3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f3:f3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f3:f3"].Style.Font.Size = 12;
            sheet.Range["f3:f3"].Merge(); // birlashtirish
            sheet.Range["f3:f3"].Text = "";
            sheet.Range["f3:f3"].Style.Font.Underline = FontUnderlineType.SingleAccounting;

            sheet.Range["b4:j4"].Style.Font.IsBold = false;
            sheet.Range["b4:j4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b4:j4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b4:j4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b4:j4"].Style.Font.Size = 11;
            sheet.Range["b4:j4"].Merge(); // birlashtirish
            sheet.Range["b4:j4"].Text = "     ";
            sheet.Range["b4:j4"].Style.Font.Underline = FontUnderlineType.SingleAccounting;

            var send = "SELECT * FROM spravochnik_main_jur7 where user_jur7='" + string_for_otdels + "'";

            sql1.myReader = sql1.return_MySqlCommand(send).ExecuteReader();
            while (sql1.myReader.Read())
            {
                sheet.Range["a5:d5"].Style.Font.IsBold = false;
                sheet.Range["a5:d5"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a5:d5"].Style.VerticalAlignment = VerticalAlignType.Bottom;
                sheet.Range["a5:d5"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a5:d5"].Style.Font.Size = 11;
                sheet.Range["a5:d5"].Merge(); // birlashtirish
                sheet.Range["a5:d5"].Text = "Поставщик :" + (sql1.myReader["naim_org"] != DBNull.Value ? sql1.myReader.GetString("naim_org") : "");
                sheet.Range["a5:d5"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);

                //naim_org,adres,ras_s,bank,inn,okxn
                sheet.Range["a6:d6"].Style.Font.IsBold = false;
                sheet.Range["a6:d6"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a6:d6"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a6:d6"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a6:d6"].Style.Font.Size = 11;
                sheet.Range["a6:d6"].Merge(); // birlashtirish
                sheet.Range["a6:d6"].Text = "Адрес :" + (sql1.myReader["adres"] != DBNull.Value ? sql1.myReader.GetString("adres") : "");
                sheet.Range["a6:d6"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);



                sheet.Range["a7:d7"].Style.Font.IsBold = false;
                sheet.Range["a7:d7"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a7:d7"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a7:d7"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a7:d7"].Style.Font.Size = 11;
                sheet.Range["a7:d7"].Merge(); // birlashtirish
                sheet.Range["a7:d7"].Text = "Р/с :" + (sql1.myReader["ras_s"] != DBNull.Value ? sql1.myReader.GetString("ras_s") : "");
                sheet.Range["a7:d7"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);



                sheet.Range["a8:d8"].Style.Font.IsBold = false;
                sheet.Range["a8:d8"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a8:d8"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a8:d8"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a8:d8"].Style.Font.Size = 11;
                sheet.Range["a8:d8"].Merge(); // birlashtirish
                sheet.Range["a8:d8"].Text = "Банк : " + (sql1.myReader["bank"] != DBNull.Value ? sql1.myReader.GetString("bank") : "");
                sheet.Range["a8:d8"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);



                sheet.Range["a9:d9"].Style.Font.IsBold = false;
                sheet.Range["a9:d9"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a9:d9"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a9:d9"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a9:d9"].Style.Font.Size = 11;
                sheet.Range["a9:d9"].Merge(); // birlashtirish
                sheet.Range["a9:d9"].Text = "МФО :" + (sql1.myReader["mfo"] != DBNull.Value ? sql1.myReader.GetString("mfo") : "");
                sheet.Range["a9:d9"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);

                sheet.Range["a10:d10"].Style.Font.IsBold = false;
                sheet.Range["a10:d10"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a10:d10"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a10:d10"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a10:d10"].Style.Font.Size = 11;
                sheet.Range["a10:d10"].Merge(); // birlashtirish
                sheet.Range["a10:d10"].Text = "ИНН : " + (sql1.myReader["inn"] != DBNull.Value ? sql1.myReader.GetString("inn") : "") + "  " + "ОКЭТ :" + (sql1.myReader["okxn"] != DBNull.Value ? sql1.myReader.GetString("okxn") : "");
                sheet.Range["a10:d10"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);
            }
            sql1.myReader.Close();


            ///poluchatel
            /// 
            sheet.Range["e5:k5"].Style.Font.IsBold = false;
            sheet.Range["e5:k5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e5:k5"].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["e5:k5"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["e5:k5"].Style.Font.Size = 11;
            sheet.Range["e5:k5"].Merge(); // birlashtirish
            sheet.Range["e5:k5"].Text = "Получателъ :";// + postavshik_comboBox.Text;
            sheet.Range["e5:k5"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);


            sheet.Range["e6:k6"].Style.Font.IsBold = false;
            sheet.Range["e6:k6"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e6:k6"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e6:k6"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["e6:k6"].Style.Font.Size = 11;
            sheet.Range["e6:k6"].Merge(); // birlashtirish
            sheet.Range["e6:k6"].Text = "Адрес :";// + comboBox_adres_rasxod.Text;
            sheet.Range["e6:k6"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);

            sheet.Range["e7:k7"].Style.Font.IsBold = false;
            sheet.Range["e7:k7"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e7:k7"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e7:k7"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["e7:k7"].Style.Font.Size = 11;
            sheet.Range["e7:k7"].Merge(); // birlashtirish
            sheet.Range["e7:k7"].Text = "Р/с :";// + comboBox_pc_rasxod.Text;
            sheet.Range["e7:k7"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);

            sheet.Range["e8:k8"].Style.Font.IsBold = false;
            sheet.Range["e8:k8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e8:k8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e8:k8"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["e8:k8"].Style.Font.Size = 11;
            sheet.Range["e8:k8"].Merge(); // birlashtirish
            sheet.Range["e8:k8"].Text = "Банк : ";// + comboBox_bank_rasxod.Text;
            sheet.Range["e8:k8"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);

            sheet.Range["e9:k9"].Style.Font.IsBold = false;
            sheet.Range["e9:k9"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e9:k9"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e9:k9"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["e9:k9"].Style.Font.Size = 11;
            sheet.Range["e9:k9"].Merge(); // birlashtirish
            sheet.Range["e9:k9"].Text = "МФО :";// + comboBox_mfo_rasxod.Text;
            sheet.Range["e9:k9"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);

            sheet.Range["e10:k10"].Style.Font.IsBold = false;
            sheet.Range["e10:k10"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e10:k10"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e10:k10"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["e10:k10"].Style.Font.Size = 11;
            sheet.Range["e10:k10"].Merge(); // birlashtirish
            sheet.Range["e10:k10"].Text = "ИНН : ";// + comboBox_inn_rasxod.Text;
            sheet.Range["e10:k10"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);



            sheet.SetRowHeight(2, 21);
            sheet.SetRowHeight(3, 18);
            sheet.SetRowHeight(4, 20);
            sheet.SetRowHeight(5, 18);
            sheet.SetRowHeight(6, 18);
            sheet.SetRowHeight(7, 18);
            sheet.SetRowHeight(8, 18);
            sheet.SetRowHeight(9, 18);
            sheet.SetRowHeight(10, 18);

            sheet.Range["a11:a11"].Merge(); // birlashtirish
            sheet.Range["a11:a11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["a11:a11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["a11:a11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["a11:a11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["a11:a11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a11:a11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a11:a11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a11:a11"].Style.Font.Size = 11;
            sheet.Range["a11:a11"].Style.WrapText = true;
            sheet.Range["a11:a11"].Text = "№";

            sheet.Range["b11:b11"].Merge(); // birlashtirish
            sheet.Range["b11:b11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["b11:b11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["b11:b11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["b11:b11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["b11:b11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b11:b11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b11:b11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b11:b11"].Style.Font.Size = 11;
            sheet.Range["b11:b11"].Style.WrapText = true;
            sheet.Range["b11:b11"].Text = "Наименование товара (работ, услуг)";

            sheet.Range["c11:c11"].Merge(); // birlashtirish
            sheet.Range["c11:c11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["c11:c11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["c11:c11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["c11:c11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["c11:c11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c11:c11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c11:c11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c11:c11"].Style.Font.Size = 11;
            sheet.Range["c11:c11"].Style.WrapText = true;
            sheet.Range["c11:c11"].Text = "Едизм";

            sheet.Range["d11:d11"].Merge(); // birlashtirish
            sheet.Range["d11:d11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["d11:d11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["d11:d11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["d11:d11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["d11:d11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d11:d11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d11:d11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d11:d11"].Style.Font.Size = 11;
            sheet.Range["d11:d11"].Style.WrapText = true;
            sheet.Range["d11:d11"].Text = "Кол-во";

            sheet.Range["e11:e11"].Merge(); // birlashtirish
            sheet.Range["e11:e11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["e11:e11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["e11:e11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["e11:e11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["e11:e11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e11:e11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e11:e11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e11:e11"].Style.Font.Size = 11;
            sheet.Range["e11:e11"].Style.WrapText = true;
            sheet.Range["e11:e11"].Text = "Цена";

            sheet.Range["f11:f11"].Merge(); // birlashtirish
            sheet.Range["f11:f11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["f11:f11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["f11:f11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["f11:f11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["f11:f11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f11:f11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f11:f11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f11:f11"].Style.Font.Size = 11;
            sheet.Range["f11:f11"].Text = "Стомость поставка";
            sheet.Range["f11:f11"].Style.WrapText = true;

            sheet.Range["g11:h11"].Merge(); // birlashtirish
            sheet.Range["g11:h11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["g11:h11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["g11:h11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["g11:h11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["g11:h11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g11:h11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g11:h11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g11:h11"].Style.Font.Size = 11;
            sheet.Range["g11:h11"].Style.WrapText = true;
            sheet.Range["g11:h11"].Text = "Акцизный налог";

            sheet.Range["i11:j11"].Merge(); // birlashtirish
            sheet.Range["i11:j11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["i11:j11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["i11:j11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["i11:j11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["i11:j11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["i11:j11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["i11:j11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["i11:j11"].Style.Font.Size = 11;
            sheet.Range["i11:j11"].Style.WrapText = true;
            sheet.Range["i11:j11"].Text = "НДС";

            sheet.Range["k11:k11"].Merge(); // birlashtirish
            sheet.Range["k11:k11"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["k11:k11"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["k11:k11"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["k11:k11"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["k11:k11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["k11:k11"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["k11:k11"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["k11:k11"].Style.Font.Size = 11;
            sheet.Range["k11:k11"].Style.WrapText = true;
            sheet.Range["k11:k11"].Text = "Стоимость поставки с НДС";


            sheet.Range["a12:a12"].Merge(); // birlashtirish
            sheet.Range["a12:a12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["a12:a12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["a12:a12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["a12:a12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["a12:a12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a12:a12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a12:a12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a12:a12"].Style.Font.Size = 11;
            sheet.Range["a12:a12"].Style.WrapText = true;
            sheet.Range["a12:a12"].Text = "1";

            sheet.Range["b12:b12"].Merge(); // birlashtirish
            sheet.Range["b12:b12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["b12:b12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["b12:b12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["b12:b12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["b12:b12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b12:b12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b12:b12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b12:b12"].Style.Font.Size = 11;
            sheet.Range["b12:b12"].Style.WrapText = true;
            sheet.Range["b12:b12"].Text = "2";

            sheet.Range["c12:c12"].Merge(); // birlashtirish
            sheet.Range["c12:c12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["c12:c12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["c12:c12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["c12:c12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["c12:c12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c12:c12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c12:c12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c12:c12"].Style.Font.Size = 11;
            sheet.Range["c12:c12"].Style.WrapText = true;
            sheet.Range["c12:c12"].Text = "3";

            sheet.Range["d12:d12"].Merge(); // birlashtirish
            sheet.Range["d12:d12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["d12:d12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["d12:d12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["d12:d12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["d12:d12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d12:d12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d12:d12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d12:d12"].Style.Font.Size = 11;
            sheet.Range["d12:d12"].Style.WrapText = true;
            sheet.Range["d12:d12"].Text = "4";

            sheet.Range["e12:e12"].Merge(); // birlashtirish
            sheet.Range["e12:e12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["e12:e12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["e12:e12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["e12:e12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["e12:e12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e12:e12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e12:e12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e12:e12"].Style.Font.Size = 11;
            sheet.Range["e12:e12"].Style.WrapText = true;
            sheet.Range["e12:e12"].Text = "5";

            sheet.Range["f12:f12"].Merge(); // birlashtirish
            sheet.Range["f12:f12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["f12:f12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["f12:f12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["f12:f12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["f12:f12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f12:f12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f12:f12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f12:f12"].Style.Font.Size = 11;
            sheet.Range["f12:f12"].Text = "6";
            sheet.Range["f12:f12"].Style.WrapText = true;

            sheet.Range["g12:g12"].Merge(); // birlashtirish
            sheet.Range["g12:g12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["g12:g12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["g12:g12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["g12:g12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["g12:g12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g12:g12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g12:g12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g12:g12"].Style.Font.Size = 11;
            sheet.Range["g12:g12"].Style.WrapText = true;
            sheet.Range["g12:g12"].Text = "7";

            sheet.Range["h12:h12"].Merge(); // birlashtirish
            sheet.Range["h12:h12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["h12:h12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["h12:h12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["h12:h12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["h12:h12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h12:h12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h12:h12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["h12:h12"].Style.Font.Size = 11;
            sheet.Range["h12:h12"].Style.WrapText = true;
            sheet.Range["h12:h12"].Text = "8";

            sheet.Range["i12:i12"].Merge(); // birlashtirish
            sheet.Range["i12:i12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["i12:i12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["i12:i12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["i12:i12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["i12:i12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["i12:i12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["i12:i12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["i12:i12"].Style.Font.Size = 11;
            sheet.Range["i12:i12"].Style.WrapText = true;
            sheet.Range["i12:i12"].Text = "9";

            sheet.Range["j12:j12"].Merge(); // birlashtirish
            sheet.Range["j12:j12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["j12:j12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["j12:j12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["j12:j12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["j12:j12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["j12:j12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["j12:j12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["j12:j12"].Style.Font.Size = 11;
            sheet.Range["j12:j12"].Style.WrapText = true;
            sheet.Range["j12:j12"].Text = "10";

            sheet.Range["k12:k12"].Merge(); // birlashtirish
            sheet.Range["k12:k12"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["k12:k12"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["k12:k12"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["k12:k12"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["k12:k12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["k12:k12"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["k12:k12"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["k12:k12"].Style.Font.Size = 11;
            sheet.Range["k12:k12"].Style.WrapText = true;
            sheet.Range["k12:k12"].Text = "11";

            //////////////////////



            int i = 0;
            int myrow = 13;
            int j = 0;

            double all_summa = 0;

            var send2 = "SELECT * FROM products_jur7 where kod_doc='" + kod_num_textBox.Text + "' and user='" + string_for_otdels + "' and month='" + month_textBox.Text + "' and year='" + year_textBox.Text + "' ";

            sql2.myReader = sql2.return_MySqlCommand(send2).ExecuteReader();
            while (sql2.myReader.Read())
            {
                j = i;
                j = j + 1;

                sheet.Range["a" + myrow + ":a" + myrow].Merge();
                sheet.Range["a" + myrow + ":a" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["a" + myrow + ":a" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.Size = 10;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["a" + myrow + ":a" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["a" + myrow + ":a" + myrow].Text = (String)j.ToString();

                sheet.Range["b" + myrow + ":b" + myrow].Merge();
                sheet.Range["b" + myrow + ":b" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["b" + myrow + ":b" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["b" + myrow + ":b" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.Size = 10;
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["b" + myrow + ":b" + myrow].Text = sql2.myReader["naim_tov"] != DBNull.Value ? sql2.myReader.GetString("naim_tov") : "";
                sheet.Range["b" + myrow + ":b" + myrow].Style.WrapText = true;


                sheet.Range["c" + myrow + ":c" + myrow].Merge();
                sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 10;
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["c" + myrow + ":c" + myrow].Value = sql2.myReader["edin"] != DBNull.Value ? sql2.myReader.GetString("edin") : "";
                sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;

                // naim_tov,edin,kol,sena,summa
                sheet.Range["d" + myrow + ":d" + myrow].Merge();
                sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
                sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["d" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 10;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["d" + myrow + ":d" + myrow].Value = sql2.myReader["kol"] != DBNull.Value ? sql2.myReader.GetString("kol") : "";

                sheet.Range["e" + myrow + ":e" + myrow].Merge();
                sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 10;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["e" + myrow + ":e" + myrow].Value = sql2.myReader["sena"] != DBNull.Value ? sql2.myReader.GetString("sena") : "";

                all_summa += sql2.myReader["summa"] != DBNull.Value ? sql2.myReader.GetDouble("summa") : 0;

                sheet.Range["f" + myrow + ":f" + myrow].Merge();
                sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":f" + myrow].Value = sql2.myReader["summa"] != DBNull.Value ? sql2.myReader.GetString("summa") : "";

                sheet.Range["g" + myrow + ":k" + myrow].Merge();
                sheet.Range["g" + myrow + ":k" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["g" + myrow + ":k" + myrow].Style.WrapText = true;
                sheet.Range["g" + myrow + ":k" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["g" + myrow + ":k" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["g" + myrow + ":k" + myrow].Style.Font.Size = 10;
                sheet.Range["g" + myrow + ":k" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["g" + myrow + ":k" + myrow].Text = "Без акц.налог Без НДС";

                myrow = myrow + 1;
                i = i + 1;

            }
            sql2.myReader.Close();


            //sheet.Range["a" + myrow + ":k" + myrow].Merge(); // birlashtirish
            //sheet.Range["a" + myrow + ":k" + myrow].Style.Font.FontName = "Times New Roman";
            //sheet.Range["a" + myrow + ":k" + myrow].Style.VerticalAlignment = VerticalAlignType.Bottom;
            //sheet.Range["a" + myrow + ":k" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            //sheet.Range["a" + myrow + ":k" + myrow].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            //sheet.Range["a" + myrow + ":k" + myrow].Style.Font.Size = 11;
            //sheet.Range["a" + myrow + ":k" + myrow].Text = "Без акц.налог Без НДС   ";
            //sheet.SetRowHeight(myrow, 18);
            //myrow++;

            sheet.Range["b" + myrow + ":d" + myrow].Merge(); // birlashtirish
            //sheet.Range["b" + myrow + ":d" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":d" + myrow].Text = "Итого : ";

            sheet.Range["e" + myrow + ":f" + myrow].Merge(); // birlashtirish
            //sheet.Range["e" + myrow + ":f" + myrow].Style.Font.IsBold = true;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["e" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Font.Size = 11;
            sheet.Range["e" + myrow + ":f" + myrow].Value = all_summa.ToString();

            sheet.SetRowHeight(myrow, 18);
            myrow++;

            String[] arr = refresh_strings_to_mysql(all_summa.ToString()).Split('.');
            b[0] = Convert.ToInt32(arr[0]);
            b[1] = Convert.ToInt32(arr[1]);

            sheet.Range["b" + myrow + ":k" + myrow].Style.Font.IsBold = true;
            //sheet.Range["b" + myrow + ":k" + myrow].Style.Font.IsItalic = true;
            sheet.Range["b" + myrow + ":k" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":k" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":k" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":k" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":k" + myrow].Merge(); // birlashtirish
            sheet.Range["b" + myrow + ":k" + myrow].Text = "Прописью :" + number_russian.toWords(b[0]) + " сум " + b[1] + " тийин";
            sheet.Range["b" + myrow + ":k" + myrow].Style.Font.Underline = FontUnderlineType.Single;

            sheet.SetRowHeight(myrow, 18);
            myrow++;

            //sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Merge(); // birlashtirish
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Нач. ФЭО: ";
            //  sheet.Range["b" + myrow + ":b" + myrow].Style.Font.Underline = FontUnderlineType.Single;

            myrow++;
            //sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Merge(); // birlashtirish
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Нач.Отдел";

            //sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["e" + myrow + ":j" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["e" + myrow + ":j" + myrow].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["e" + myrow + ":j" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["e" + myrow + ":j" + myrow].Style.Font.Size = 11;
            sheet.Range["e" + myrow + ":j" + myrow].Merge(); // birlashtirish
            sheet.Range["e" + myrow + ":j" + myrow].Text = "Получил :_________________________________";

            sheet.SetRowHeight(myrow, 18);
            myrow++;

            //sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Merge(); // birlashtirish
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Гл.Спец";

            sheet.SetRowHeight(myrow, 18);
            myrow++;

            //sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Merge(); // birlashtirish
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Отпустил : _________________________________";
            sheet.SetRowHeight(myrow, 18);

            sheet.Range["d13:" + myrow + "f"].NumberFormat = "#,##0.00";

            string kod_num = kod_num_textBox.Text;

            string filePath = Environment.CurrentDirectory + "\\docs\\qrcode\\" + string_for_otdels + "-" + year_textBox.Text + "-" + month_textBox.Text + "-" + kod_num + "-" + "1" + ".png";
            FileInfo file = new FileInfo(filePath);
            if (file.Exists)
            {
                string picPath = Environment.CurrentDirectory + "\\docs\\qrcode\\" + string_for_otdels + "-" + year_textBox.Text + "-" + month_textBox.Text + "-" + kod_num + "-" + "1" + ".png";
                ExcelPicture picture = sheet.Pictures.Add(1, 8, picPath);
                picture.Width = 60;
                picture.Height = 60;

            }
            else
            {
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode((string_for_otdels + "-" + year_textBox.Text + "-" + month_textBox.Text + "-" + kod_num + "-" + "1"), QRCodeGenerator.ECCLevel.Q);
                QRCode qrCode = new QRCode(qrCodeData);

                Bitmap bitMap = qrCode.GetGraphic(20);

                bitMap.Save(Environment.CurrentDirectory + "\\docs\\qrcode\\" + string_for_otdels + "-" + year_textBox.Text + "-" + month_textBox.Text + "-" + kod_num + "-" + "1" + ".png");

                string picPath = Environment.CurrentDirectory + "\\docs\\qrcode\\" + string_for_otdels + "-" + year_textBox.Text + "-" + month_textBox.Text + "-" + kod_num + "-" + "1" + ".png";
                ExcelPicture picture = sheet.Pictures.Add(1, 8, picPath);
                picture.Width = 60;
                picture.Height = 60;


            }

            workbook.SaveToFile(Environment.CurrentDirectory + "\\docs\\Счет-фастура.xlsx", Spire.Xls.FileFormat.Version2007);
            System.Diagnostics.Process.Start(workbook.FileName);


            //  }
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Накладная_Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void qrcode_dataGridView_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            try
            {
                DialogResult dialogResult = MessageBox.Show("Вы действительно хотите удалить данные?", "Удаление", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    foreach (DataGridViewRow row in qrcode_dataGridView.SelectedRows)
                    {
                        if (row.Cells[0].Value != null && row.Cells[18].Value != null)
                        {

                            sql.return_MySqlCommand("delete from products_prixod_jur7 where id = " + row.Cells[0].Value + "").ExecuteNonQuery();
                            sql.return_MySqlCommand("delete from products_jur7 where id = " + row.Cells[18].Value + "").ExecuteNonQuery();
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

        int exist = 0;
        int product_id;
        int id_sklad_products;


        public void set_items_to_values_prixod()
        {
            try
            {

                this.qrcode_dataGridView.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.qrcode_dataGridView_CellValueChanged);

                string debet_01 = "";
                double debet_01_sum = 0;
                double debet_06_sum = 0;
                double debet_07_sum = 0;

                jur_order_qrcode_textBox.Text = "";
                num_qrcode_textBox.Text = "";
                primech_qrcode_textBox.Text = "";
                doveren_qrcode_textBox.Text = "";
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";

                var query = "SELECT * FROM products_jur7 where vid_doc='" + vid_doc + "' and year='" + year_textBox.Text + "' and month='" + month_textBox.Text + "' and kod_doc='" + kod_num_textBox.Text + "' and user='" + string_for_otdels + "' group by kod_doc";
                sql.myReader = sql.return_MySqlCommand(query).ExecuteReader();
                while (sql.myReader.Read())
                {
                    jur_order_qrcode_textBox.Text = (sql.myReader["jur_order"] != DBNull.Value ? sql.myReader.GetString("jur_order") : "");
                    num_qrcode_textBox.Text = (sql.myReader["num_doc"] != DBNull.Value ? sql.myReader.GetString("num_doc") : "");
                    data_qrcode_DateTimePicker.Value = (sql.myReader["date_doc"] != DBNull.Value ? sql.myReader.GetDateTime("date_doc") : DateTime.Now);
                    primech_qrcode_textBox.Text = (sql.myReader["primech"] != DBNull.Value ? sql.myReader.GetString("primech") : "");
                    doveren_qrcode_textBox.Text = (sql.myReader["doverennost"] != DBNull.Value ? sql.myReader.GetString("doverennost") : "");
                    textBox1.Text = (sql.myReader["ot_kogo"] != DBNull.Value ? sql.myReader.GetString("ot_kogo") : "");
                    textBox2.Text = (sql.myReader["ot_kogo_2"] != DBNull.Value ? sql.myReader.GetString("ot_kogo_2") : "");
                    textBox3.Text = (sql.myReader["komu_1"] != DBNull.Value ? sql.myReader.GetString("komu_1") : "");
                    textBox4.Text = (sql.myReader["komu_2"] != DBNull.Value ? sql.myReader.GetString("komu_2") : "");

                    kod_num_textBox.Text = (sql.myReader["kod_doc"] != DBNull.Value ? sql.myReader.GetString("kod_doc") : "");

                    year_textBox.Text = (sql.myReader["year"] != DBNull.Value ? sql.myReader.GetString("year") : "");
                    month_textBox.Text = (sql.myReader["month"] != DBNull.Value ? sql.myReader.GetString("month") : "");

                    vid_doc = (sql.myReader["vid_doc"] != DBNull.Value ? sql.myReader.GetString("vid_doc") : "");
                }
                sql.myReader.Close();




                qrcode_dataGridView.Rows.Clear();

                int id_product = 0;


                var select_ras = " SELECT * FROM products_jur7 where vid_doc='"+vid_doc+"' and year='" + year_textBox.Text + "' and month='" + month_textBox.Text + "' and kod_doc='" + kod_num_textBox.Text + "' and user='" + string_for_otdels + "' ";
                sql.myReader = sql.return_MySqlCommand(select_ras).ExecuteReader();
                while (sql.myReader.Read())
                {

                    //kod_tov,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,

                    int index = qrcode_dataGridView.Rows.Add();
                    qrcode_dataGridView.Rows[index].Cells[0].Value = (sql.myReader["id"] != DBNull.Value ? sql.myReader.GetString("id") : "");
                    qrcode_dataGridView.Rows[index].Cells[1].Value = (sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetString("product_id") : "");
                    qrcode_dataGridView.Rows[index].Cells[2].Value = (sql.myReader["gruppa"] != DBNull.Value ? sql.myReader.GetString("gruppa") : "");
                    qrcode_dataGridView.Rows[index].Cells[3].Value = (sql.myReader["naim_tov"] != DBNull.Value ? sql.myReader.GetString("naim_tov") : "");
                    qrcode_dataGridView.Rows[index].Cells[4].Value = (sql.myReader["edin"] != DBNull.Value ? sql.myReader.GetString("edin") : "");
                    qrcode_dataGridView.Rows[index].Cells[5].Value = (sql.myReader["inventar_num"] != DBNull.Value ? sql.myReader.GetString("inventar_num") : "");
                    qrcode_dataGridView.Rows[index].Cells[6].Value = (sql.myReader["seria_num"] != DBNull.Value ? sql.myReader.GetString("seria_num") : "");
                    string kols = sql.myReader["kol"] != DBNull.Value ? sql.myReader.GetString("kol") : "";

                    if (kols.Length <= 3)
                    {
                        qrcode_dataGridView.Rows[qrcode_dataGridView.Rows.Count - 2].Cells[7].Value = string.Format("{0:#0.00}", (sql.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("kol").Replace(".", ","))) : 0));
                    }
                    if (kols.Length > 3)
                    {
                        qrcode_dataGridView.Rows[qrcode_dataGridView.Rows.Count - 2].Cells[7].Value = string.Format("{0:#,###.00}", (sql.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("kol").Replace(".", ","))) : 0));
                    }

                    string sena = sql.myReader["sena"] != DBNull.Value ? sql.myReader.GetString("sena") : "";

                    if (sena.Length <= 3)
                    {
                        qrcode_dataGridView.Rows[qrcode_dataGridView.Rows.Count - 2].Cells[8].Value = string.Format("{0:#0.00}", (sql.myReader["sena"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("sena").Replace(".", ","))) : 0));
                    }
                    if (sena.Length > 3)
                    {
                        qrcode_dataGridView.Rows[qrcode_dataGridView.Rows.Count - 2].Cells[8].Value = string.Format("{0:#,###.00}", (sql.myReader["sena"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("sena").Replace(".", ","))) : 0));
                    }

                    string summa = sql.myReader["summa"] != DBNull.Value ? sql.myReader.GetString("summa") : "";

                    if (summa.Length <= 3)
                    {
                        qrcode_dataGridView.Rows[qrcode_dataGridView.Rows.Count - 2].Cells[9].Value = string.Format("{0:#0.00}", (sql.myReader["summa"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa").Replace(".", ","))) : 0));
                    }
                    if (summa.Length > 3)
                    {
                        qrcode_dataGridView.Rows[qrcode_dataGridView.Rows.Count - 2].Cells[9].Value = string.Format("{0:#,###.00}", (sql.myReader["summa"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa").Replace(".", ","))) : 0));
                    }

                    //deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,id_products

                    qrcode_dataGridView.Rows[index].Cells[10].Value = (sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch") : "");


                    debet_01 = (sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch") : "");
                    string first = debet_01.Substring(0, 2);

                    if (first == "01")
                    {
                        debet_01_sum += (qrcode_dataGridView.Rows[index].Cells[9].Value != null ? Double.Parse(qrcode_dataGridView.Rows[index].Cells[9].Value.ToString()) : 0);
                    }
                    else if (first == "06")
                    {
                        debet_06_sum += (qrcode_dataGridView.Rows[index].Cells[9].Value != null ? Double.Parse(qrcode_dataGridView.Rows[index].Cells[9].Value.ToString()) : 0);
                    }
                    else if (first == "07")
                    {
                        debet_07_sum += (qrcode_dataGridView.Rows[index].Cells[9].Value != null ? Double.Parse(qrcode_dataGridView.Rows[index].Cells[9].Value.ToString()) : 0);
                    }

                    qrcode_dataGridView.Rows[index].Cells[11].Value = (sql.myReader["deb_sch_2"] != DBNull.Value ? sql.myReader.GetString("deb_sch_2") : "");
                    qrcode_dataGridView.Rows[index].Cells[12].Value = (sql.myReader["kre_sch"] != DBNull.Value ? sql.myReader.GetString("kre_sch") : "");
                    qrcode_dataGridView.Rows[index].Cells[13].Value = (sql.myReader["kre_sch_2"] != DBNull.Value ? sql.myReader.GetString("kre_sch_2") : "");
                    qrcode_dataGridView.Rows[index].Cells[14].Value = (sql.myReader["provodka_iznos"] != DBNull.Value ? sql.myReader.GetString("provodka_iznos") : "");
                    qrcode_dataGridView.Rows[index].Cells[15].Value = (sql.myReader["provodka_iznos_2"] != DBNull.Value ? sql.myReader.GetString("provodka_iznos_2") : "");

                    string summa_iznos = sql.myReader["summa_iznos"] != DBNull.Value ? sql.myReader.GetString("summa_iznos") : "";

                    if (summa_iznos.Length <= 3)
                    {
                        qrcode_dataGridView.Rows[qrcode_dataGridView.Rows.Count - 2].Cells[16].Value = string.Format("{0:#0.00}", (sql.myReader["summa_iznos"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa_iznos").Replace(".", ","))) : 0));
                    }
                    if (summa_iznos.Length > 3)
                    {
                        qrcode_dataGridView.Rows[qrcode_dataGridView.Rows.Count - 2].Cells[16].Value = string.Format("{0:#,###.00}", (sql.myReader["summa_iznos"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa_iznos").Replace(".", ","))) : 0));
                    }

                    qrcode_dataGridView.Rows[index].Cells[17].Value = (sql.myReader["date_pr"] != DBNull.Value ? (DateTime.Parse(sql.myReader.GetString("date_pr")).ToString("dd.MM.yyyy")) : "");



                    var id_sklad_products = " SELECT id FROM products_prixod_jur7 where id_sklad_products = '" + (sql.myReader["id"] != DBNull.Value ? sql.myReader.GetString("id") : "") + "' ";
                    sql2.myReader = sql2.return_MySqlCommand(id_sklad_products).ExecuteReader();
                    while (sql2.myReader.Read())
                    {
                        id_product = (sql2.myReader["id"] != DBNull.Value ? sql2.myReader.GetInt32("id") : 0);
                        qrcode_dataGridView.Rows[index].Cells[18].Value = id_product;
                    }
                    sql2.myReader.Close();

                    var id_sklad_products2 = " SELECT id FROM products_rasxod_jur7 where id_sklad_products = '" + (sql.myReader["id"] != DBNull.Value ? sql.myReader.GetString("id") : "") + "' ";
                    sql2.myReader = sql2.return_MySqlCommand(id_sklad_products2).ExecuteReader();
                    while (sql2.myReader.Read())
                    {
                        id_product = (sql2.myReader["id"] != DBNull.Value ? sql2.myReader.GetInt32("id") : 0);
                        qrcode_dataGridView.Rows[index].Cells[18].Value = id_product;
                    }
                    sql2.myReader.Close();

                    var id_sklad_products3 = " SELECT id FROM products_vnut_per_jur7 where id_sklad_products = '" + (sql.myReader["id"] != DBNull.Value ? sql.myReader.GetString("id") : "") + "' ";
                    sql2.myReader = sql2.return_MySqlCommand(id_sklad_products3).ExecuteReader();
                    while (sql2.myReader.Read())
                    {
                        id_product = (sql2.myReader["id"] != DBNull.Value ? sql2.myReader.GetInt32("id") : 0);
                        qrcode_dataGridView.Rows[index].Cells[18].Value = id_product;
                    }
                    sql2.myReader.Close();


                    //sklad_dataGridView.Rows[index].Cells[3].Value = refresh_strings_to_mysql(sql.myReader["sena"] != DBNull.Value ? string.Format("{0:#0.00}", sql.myReader.GetDouble("sena")) : "0");
                    //qrcode_dataGridView.Rows[index].Cells[18].Value = (sql.myReader["id_products"] != DBNull.Value ? sql.myReader.GetString("id_products") : "");


                }
                sql.myReader.Close();


                if (debet_01_sum.ToString().Length <= 3)
                {
                    nol_bir_lbl.Text = string.Format("{0:#0.00}", debet_01_sum);
                }
                if (debet_01_sum.ToString().Length > 3)
                {
                    nol_bir_lbl.Text = string.Format("{0:#0,000.00}", debet_01_sum);
                }

                if (debet_06_sum.ToString().Length <= 3)
                {
                    nol_olti_lbl.Text = string.Format("{0:#0.00}", debet_06_sum);
                }
                if (debet_06_sum.ToString().Length > 3)
                {
                    nol_olti_lbl.Text = string.Format("{0:#0,000.00}", debet_06_sum);
                }
                if (debet_07_sum.ToString().Length <= 3)
                {
                    nol_7_lbl.Text = string.Format("{0:#0.00}", debet_07_sum);
                }
                if (debet_07_sum.ToString().Length > 3)
                {
                    nol_7_lbl.Text = string.Format("{0:#0,000.00}", debet_07_sum);
                }


                this.qrcode_dataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.qrcode_dataGridView_CellValueChanged);
                label_update_prixod();

            }
            catch (Exception ex)
            {
                sql.myReader.Close();
                MessageBox.Show("Хато маълумот киритилган (" + ex.Message + ")");
            }
        }


        private void qrcode_save_btn_Click(object sender, EventArgs e)
        {

            try
            {
                if (kod_num_textBox.Text != "")
                {
                    var ext = "select exists(SELECT * FROM products_jur7 where vid_doc='" + vid_doc + "' and year='" + year_textBox.Text + "' and month='" + month_textBox.Text + "' and kod_doc='" + kod_num_textBox.Text + "' and user='" + string_for_otdels + "' ) as ex";

                    sql.myReader = sql.return_MySqlCommand(ext).ExecuteReader();
                    while (sql.myReader.Read())
                    {
                        exist = sql.myReader.GetInt32("ex");

                    }
                    sql.myReader.Close();

                    if (exist == 0)
                    {

                        //for (int i = 0; i < qrcode_dataGridView.Rows.Count - 1; i++)
                        //{

                        //    if (vid_doc == "1")
                        //    {

                        //        var naim_tov = "insert into naim_tov_jur7 (naim,kod_gruppa,sena,month,year) values(" +
                        //                    "'" + (qrcode_dataGridView.Rows[i].Cells[3].Value != null ? qrcode_dataGridView.Rows[i].Cells[3].Value.ToString() : "") + "'," +
                        //                     "'" + (qrcode_dataGridView.Rows[i].Cells[2].Value != null ? qrcode_dataGridView.Rows[i].Cells[2].Value.ToString() : "") + "'," +
                        //                     "'" + (qrcode_dataGridView.Rows[i].Cells[8].Value != null ? qrcode_dataGridView.Rows[i].Cells[8].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                     "'" + (month_textBox.Text) + "'," +
                        //                     "'" + (year_textBox.Text) + "'" +
                        //                    ")";
                        //        sql.return_MySqlCommand(naim_tov).ExecuteNonQuery();


                        //        var query3 = "SELECT max(id) as product_id FROM naim_tov_jur7";
                        //        sql.myReader = sql.return_MySqlCommand(query3).ExecuteReader();
                        //        while (sql.myReader.Read())
                        //        {
                        //            product_id = sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetInt32("product_id") : 1;
                        //        }
                        //        sql.myReader.Close();


                        //        var insert_product = "insert into products_jur7 (user,year,month,vid_doc,kod_doc,product_id,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,komu_1,komu_2) values(" +
                        //                        "'" + string_for_otdels + "'," +
                        //                        "'" + (year_textBox.Text) + "'," +
                        //                        "'" + (month_textBox.Text) + "'," +
                        //                        "'" + ("1") + "'," +
                        //                        "'" + (kod_num_textBox.Text) + "'," +
                        //                         "'" + (product_id) + "'," +
                        //                        "'" + (qrcode_dataGridView.Rows[i].Cells[2].Value != null ? qrcode_dataGridView.Rows[i].Cells[2].Value.ToString() : "") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[3].Value != null ? qrcode_dataGridView.Rows[i].Cells[3].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[4].Value != null ? qrcode_dataGridView.Rows[i].Cells[4].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[5].Value != null ? qrcode_dataGridView.Rows[i].Cells[5].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[6].Value != null ? qrcode_dataGridView.Rows[i].Cells[6].Value.ToString() : "") + "'," +
                        //                        "'" + (qrcode_dataGridView.Rows[i].Cells[7].Value != null ? qrcode_dataGridView.Rows[i].Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                        "'" + (qrcode_dataGridView.Rows[i].Cells[8].Value != null ? qrcode_dataGridView.Rows[i].Cells[8].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                        "'" + (qrcode_dataGridView.Rows[i].Cells[9].Value != null ? qrcode_dataGridView.Rows[i].Cells[9].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                        "'" + (qrcode_dataGridView.Rows[i].Cells[10].Value != null ? qrcode_dataGridView.Rows[i].Cells[10].Value.ToString() : "") + "'," +
                        //                        "'" + (qrcode_dataGridView.Rows[i].Cells[11].Value != null ? qrcode_dataGridView.Rows[i].Cells[11].Value.ToString() : "") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[12].Value != null ? qrcode_dataGridView.Rows[i].Cells[12].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[13].Value != null ? qrcode_dataGridView.Rows[i].Cells[13].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[14].Value != null ? qrcode_dataGridView.Rows[i].Cells[14].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[15].Value != null ? qrcode_dataGridView.Rows[i].Cells[15].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[16].Value != null ? qrcode_dataGridView.Rows[i].Cells[16].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                        "'" + (qrcode_dataGridView.Rows[i].Cells[17].Value != null ? DateTime.Parse(qrcode_dataGridView.Rows[i].Cells[17].Value.ToString()).ToString("yyyy-MM-dd") : DateTime.Now.ToString("yyyy-MM-dd")) + "', " +
                        //                        "'" + (jur_order_qrcode_textBox.Text) + "', " +
                        //                        "'" + (num_qrcode_textBox.Text) + "', " +
                        //                        (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                        //                        "'" + (primech_qrcode_textBox.Text) + "', " +
                        //                        "'" + (doveren_qrcode_textBox.Text) + "', " +
                        //                        "'" + (textBox1.Text) + "', " +
                        //                        "'" + (textBox3.Text) + "', " +
                        //                        "'" + (textBox4.Text) + "' " +
                        //                        ")";
                        //        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();


                        //        var id_sklad = "SELECT max(id) as id_sklad FROM products_jur7";
                        //        sql.myReader = sql.return_MySqlCommand(id_sklad).ExecuteReader();
                        //        while (sql.myReader.Read())
                        //        {
                        //            id_sklad_products = sql.myReader["id_sklad"] != DBNull.Value ? sql.myReader.GetInt32("id_sklad") : 1;
                        //        }
                        //        sql.myReader.Close();


                        //        var insert_product_pri = "insert into products_prixod_jur7 (user,year,month,vid_doc,product_id,id_sklad_products,kod_doc,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,komu_1,komu_2) values(" +
                        //                         "'" + string_for_otdels + "'," +
                        //                         "'" + (year_textBox.Text) + "'," +
                        //                         "'" + (month_textBox.Text) + "'," +
                        //                         "'" + ("1") + "'," +
                        //                         "'" + (product_id) + "'," +
                        //                         "'" + (id_sklad_products) + "'," +
                        //                         "'" + (kod_num_textBox.Text) + "'," +
                        //                         //"'" + (naryad_num_prixod_int) + "'," +//kod_tov orniga
                        //                         //"'" + (qrcode_dataGridView.Rows[i].Cells[1].Value != null ? qrcode_dataGridView.Rows[i].Cells[1].Value.ToString() : "") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[2].Value != null ? qrcode_dataGridView.Rows[i].Cells[2].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[3].Value != null ? qrcode_dataGridView.Rows[i].Cells[3].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[4].Value != null ? qrcode_dataGridView.Rows[i].Cells[4].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[5].Value != null ? qrcode_dataGridView.Rows[i].Cells[5].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[6].Value != null ? qrcode_dataGridView.Rows[i].Cells[6].Value.ToString() : "") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[7].Value != null ? qrcode_dataGridView.Rows[i].Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[8].Value != null ? qrcode_dataGridView.Rows[i].Cells[8].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[9].Value != null ? qrcode_dataGridView.Rows[i].Cells[9].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[10].Value != null ? qrcode_dataGridView.Rows[i].Cells[10].Value.ToString() : "") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[11].Value != null ? qrcode_dataGridView.Rows[i].Cells[11].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[12].Value != null ? qrcode_dataGridView.Rows[i].Cells[12].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[13].Value != null ? qrcode_dataGridView.Rows[i].Cells[13].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[14].Value != null ? qrcode_dataGridView.Rows[i].Cells[14].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[15].Value != null ? qrcode_dataGridView.Rows[i].Cells[15].Value.ToString() : "") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[16].Value != null ? qrcode_dataGridView.Rows[i].Cells[16].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[17].Value != null ? DateTime.Parse(qrcode_dataGridView.Rows[i].Cells[17].Value.ToString()).ToString("yyyy-MM-dd") : DateTime.Now.ToString("yyyy-MM-dd")) + "', " +
                        //                         //"'" + (qrcode_dataGridView.Rows[i].Cells[17].Value != null ? qrcode_dataGridView.Rows[i].Cells[17].Value.ToString() : "") + "'," +
                        //                         //"" + (qrcode_dataGridView.Rows[i].Cells[17].Value != null ? "'" + DateTime.Parse(qrcode_dataGridView.Rows[i].Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                        //                         //"'" + (qrcode_dataGridView.Rows[i].Cells[18].Value != null ? qrcode_dataGridView.Rows[i].Cells[18].Value.ToString() : "") + "'," +
                        //                         "'" + (jur_order_qrcode_textBox.Text) + "', " +
                        //                         "'" + (num_qrcode_textBox.Text) + "', " +
                        //                         (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                        //                         "'" + (primech_qrcode_textBox.Text) + "', " +
                        //                         "'" + (doveren_qrcode_textBox.Text) + "', " +
                        //                         "'" + (textBox1.Text) + "', " +
                        //                         "'" + (textBox3.Text) + "', " +
                        //                         "'" + (textBox4.Text) + "' " +
                        //                         ")";
                        //        sql.return_MySqlCommand(insert_product_pri).ExecuteNonQuery();

                        //    }
                        //    else if (vid_doc == "2")
                        //    {
                        //        var insert_product = "insert into products_jur7 (user,year,month,vid_doc,kod_doc,product_id,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,ot_kogo_2,komu_1) values(" +
                        //                                                    "'" + string_for_otdels + "'," +
                        //                                                    "'" + (year_textBox.Text) + "'," +
                        //                                                    "'" + (month_textBox.Text) + "'," +
                        //                                                    "'" + ("2") + "'," +
                        //                                                    "'" + (kod_num_textBox.Text) + "'," +
                        //                                                     //"'" + (product_id) + "'," +
                        //                                                     "'" + (qrcode_dataGridView.Rows[i].Cells[1].Value != null ? qrcode_dataGridView.Rows[i].Cells[1].Value.ToString() : "") + "'," +
                        //                                                    "'" + (qrcode_dataGridView.Rows[i].Cells[2].Value != null ? qrcode_dataGridView.Rows[i].Cells[2].Value.ToString() : "") + "'," +
                        //                                                     "'" + (qrcode_dataGridView.Rows[i].Cells[3].Value != null ? qrcode_dataGridView.Rows[i].Cells[3].Value.ToString() : "") + "'," +
                        //                                                      "'" + (qrcode_dataGridView.Rows[i].Cells[4].Value != null ? qrcode_dataGridView.Rows[i].Cells[4].Value.ToString() : "") + "'," +
                        //                                                      "'" + (qrcode_dataGridView.Rows[i].Cells[5].Value != null ? qrcode_dataGridView.Rows[i].Cells[5].Value.ToString() : "") + "'," +
                        //                                                      "'" + (qrcode_dataGridView.Rows[i].Cells[6].Value != null ? qrcode_dataGridView.Rows[i].Cells[6].Value.ToString() : "") + "'," +
                        //                                                    "'" + refresh_strings_to_mysql(qrcode_dataGridView.Rows[i].Cells[7].Value != null ? qrcode_dataGridView.Rows[i].Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                                                    "'" + refresh_strings_to_mysql(qrcode_dataGridView.Rows[i].Cells[8].Value != null ? qrcode_dataGridView.Rows[i].Cells[8].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                                                    "'" + refresh_strings_to_mysql(qrcode_dataGridView.Rows[i].Cells[9].Value != null ? qrcode_dataGridView.Rows[i].Cells[9].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                                                    "'" + (qrcode_dataGridView.Rows[i].Cells[10].Value != null ? qrcode_dataGridView.Rows[i].Cells[10].Value.ToString() : "") + "'," +
                        //                                                    "'" + (qrcode_dataGridView.Rows[i].Cells[11].Value != null ? qrcode_dataGridView.Rows[i].Cells[11].Value.ToString() : "") + "'," +
                        //                                                     "'" + (qrcode_dataGridView.Rows[i].Cells[12].Value != null ? qrcode_dataGridView.Rows[i].Cells[12].Value.ToString() : "") + "'," +
                        //                                                      "'" + (qrcode_dataGridView.Rows[i].Cells[13].Value != null ? qrcode_dataGridView.Rows[i].Cells[13].Value.ToString() : "") + "'," +
                        //                                                      "'" + (qrcode_dataGridView.Rows[i].Cells[14].Value != null ? qrcode_dataGridView.Rows[i].Cells[14].Value.ToString() : "") + "'," +
                        //                                                      "'" + (qrcode_dataGridView.Rows[i].Cells[15].Value != null ? qrcode_dataGridView.Rows[i].Cells[15].Value.ToString() : "") + "'," +
                        //                                                    "'" + refresh_strings_to_mysql(qrcode_dataGridView.Rows[i].Cells[16].Value != null ? qrcode_dataGridView.Rows[i].Cells[16].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                                                   "" + (qrcode_dataGridView.Rows[i].Cells[17].Value != null ? "'" + DateTime.Parse(qrcode_dataGridView.Rows[i].Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                        //                                                    //"" + (qrcode_dataGridView.Rows[i].Cells[17].Value == null ? "'" + DateTime.Parse(qrcode_dataGridView.Rows[i].Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                        //                                                    "'" + (jur_order_qrcode_textBox.Text) + "', " +
                        //                                                    "'" + (num_qrcode_textBox.Text) + "', " +
                        //                                                    (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                        //                                                    "'" + (primech_qrcode_textBox.Text) + "', " +
                        //                                                    "'" + (doveren_qrcode_textBox.Text) + "', " +
                        //                                                    "'" + (textBox1.Text) + "', " +
                        //                                                    "'" + (textBox2.Text) + "', " +
                        //                                                    "'" + (textBox3.Text) + "' " +
                        //                                                    ")";
                        //        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();


                        //        var id_sklad = "SELECT max(id) as id_sklad FROM products_jur7";
                        //        sql.myReader = sql.return_MySqlCommand(id_sklad).ExecuteReader();
                        //        while (sql.myReader.Read())
                        //        {
                        //            id_sklad_products = sql.myReader["id_sklad"] != DBNull.Value ? sql.myReader.GetInt32("id_sklad") : 1;
                        //        }
                        //        sql.myReader.Close();

                        //        var insert_product_pri = "insert into products_rasxod_jur7 (user,year,month,vid_doc,id_sklad_products,kod_doc,product_id,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,ot_kogo_2,komu_1) values(" +
                        //                         "'" + string_for_otdels + "'," +
                        //                         "'" + (year_textBox.Text) + "'," +
                        //                         "'" + (month_textBox.Text) + "'," +
                        //                         "'" + ("2") + "'," +
                        //                         "'" + (id_sklad_products) + "'," +
                        //                         "'" + (kod_num_textBox.Text) + "'," +
                        //                         //"'" + (naryad_num_prixod_int) + "'," +//kod_tov orniga
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[1].Value != null ? qrcode_dataGridView.Rows[i].Cells[1].Value.ToString() : "") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[2].Value != null ? qrcode_dataGridView.Rows[i].Cells[2].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[3].Value != null ? qrcode_dataGridView.Rows[i].Cells[3].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[4].Value != null ? qrcode_dataGridView.Rows[i].Cells[4].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[5].Value != null ? qrcode_dataGridView.Rows[i].Cells[5].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[6].Value != null ? qrcode_dataGridView.Rows[i].Cells[6].Value.ToString() : "") + "'," +
                        //                         "'" + refresh_strings_to_mysql(qrcode_dataGridView.Rows[i].Cells[7].Value != null ? qrcode_dataGridView.Rows[i].Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                         "'" + refresh_strings_to_mysql(qrcode_dataGridView.Rows[i].Cells[8].Value != null ? qrcode_dataGridView.Rows[i].Cells[8].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                         "'" + refresh_strings_to_mysql(qrcode_dataGridView.Rows[i].Cells[9].Value != null ? qrcode_dataGridView.Rows[i].Cells[9].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[10].Value != null ? qrcode_dataGridView.Rows[i].Cells[10].Value.ToString() : "") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[11].Value != null ? qrcode_dataGridView.Rows[i].Cells[11].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[12].Value != null ? qrcode_dataGridView.Rows[i].Cells[12].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[13].Value != null ? qrcode_dataGridView.Rows[i].Cells[13].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[14].Value != null ? qrcode_dataGridView.Rows[i].Cells[14].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[15].Value != null ? qrcode_dataGridView.Rows[i].Cells[15].Value.ToString() : "") + "'," +
                        //                         "'" + refresh_strings_to_mysql(qrcode_dataGridView.Rows[i].Cells[16].Value != null ? qrcode_dataGridView.Rows[i].Cells[16].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                          "" + (qrcode_dataGridView.Rows[i].Cells[17].Value != null ? "'" + DateTime.Parse(qrcode_dataGridView.Rows[i].Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                        //                         //"'" + (qrcode_dataGridView.Rows[i].Cells[18].Value != null ? qrcode_dataGridView.Rows[i].Cells[18].Value.ToString() : "") + "'," +
                        //                         "'" + (jur_order_qrcode_textBox.Text) + "', " +
                        //                         "'" + (num_qrcode_textBox.Text) + "', " +
                        //                         (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                        //                         "'" + (primech_qrcode_textBox.Text) + "', " +
                        //                         "'" + (doveren_qrcode_textBox.Text) + "', " +
                        //                         "'" + (textBox1.Text) + "', " +
                        //                         "'" + (textBox2.Text) + "', " +
                        //                         "'" + (textBox3.Text) + "' " +
                        //                         ")";
                        //        sql.return_MySqlCommand(insert_product_pri).ExecuteNonQuery();
                        //    }
                        //    else if (vid_doc == "3")
                        //    {
                        //        var query3 = "SELECT max(id) as product_id FROM naim_tov_jur7";
                        //        sql.myReader = sql.return_MySqlCommand(query3).ExecuteReader();
                        //        while (sql.myReader.Read())
                        //        {
                        //            product_id = sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetInt32("product_id") : 1;
                        //        }
                        //        sql.myReader.Close();


                        //        var insert_product = "insert into products_jur7 (user,year,month,vid_doc,kod_doc,product_id,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,ot_kogo_2,komu_1,komu_2) values(" +
                        //                        "'" + string_for_otdels + "'," +
                        //                        "'" + (year_textBox.Text) + "'," +
                        //                        "'" + (month_textBox.Text) + "'," +
                        //                        "'" + ("3") + "'," +
                        //                        "'" + (kod_num_textBox.Text) + "'," +
                        //                         //"'" + (product_id) + "'," +
                        //                         //"'" + (product_id) + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[1].Value != null ? qrcode_dataGridView.Rows[i].Cells[1].Value.ToString() : "") + "'," +
                        //                        "'" + (qrcode_dataGridView.Rows[i].Cells[2].Value != null ? qrcode_dataGridView.Rows[i].Cells[2].Value.ToString() : "") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[3].Value != null ? qrcode_dataGridView.Rows[i].Cells[3].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[4].Value != null ? qrcode_dataGridView.Rows[i].Cells[4].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[5].Value != null ? qrcode_dataGridView.Rows[i].Cells[5].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[6].Value != null ? qrcode_dataGridView.Rows[i].Cells[6].Value.ToString() : "") + "'," +
                        //                        "'" + refresh_strings_to_mysql(qrcode_dataGridView.Rows[i].Cells[7].Value != null ? qrcode_dataGridView.Rows[i].Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                        "'" + refresh_strings_to_mysql(qrcode_dataGridView.Rows[i].Cells[8].Value != null ? qrcode_dataGridView.Rows[i].Cells[8].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                        "'" + refresh_strings_to_mysql(qrcode_dataGridView.Rows[i].Cells[9].Value != null ? qrcode_dataGridView.Rows[i].Cells[9].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                        "'" + (qrcode_dataGridView.Rows[i].Cells[10].Value != null ? qrcode_dataGridView.Rows[i].Cells[10].Value.ToString() : "") + "'," +
                        //                        "'" + (qrcode_dataGridView.Rows[i].Cells[11].Value != null ? qrcode_dataGridView.Rows[i].Cells[11].Value.ToString() : "") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[12].Value != null ? qrcode_dataGridView.Rows[i].Cells[12].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[13].Value != null ? qrcode_dataGridView.Rows[i].Cells[13].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[14].Value != null ? qrcode_dataGridView.Rows[i].Cells[14].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[15].Value != null ? qrcode_dataGridView.Rows[i].Cells[15].Value.ToString() : "") + "'," +
                        //                        "'" + refresh_strings_to_mysql(qrcode_dataGridView.Rows[i].Cells[16].Value != null ? qrcode_dataGridView.Rows[i].Cells[16].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                       "" + (qrcode_dataGridView.Rows[i].Cells[17].Value != null ? "'" + DateTime.Parse(qrcode_dataGridView.Rows[i].Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                        //                        //"" + (qrcode_dataGridView.Rows[i].Cells[17].Value == null ? "'" + DateTime.Parse(qrcode_dataGridView.Rows[i].Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                        //                        "'" + (jur_order_qrcode_textBox.Text) + "', " +
                        //                        "'" + (num_qrcode_textBox.Text) + "', " +
                        //                        (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                        //                        "'" + (primech_qrcode_textBox.Text) + "', " +
                        //                        "'" + (doveren_qrcode_textBox.Text) + "', " +
                        //                        "'" + (textBox1.Text) + "', " +
                        //                        "'" + (textBox2.Text) + "', " +
                        //                        "'" + (textBox3.Text) + "', " +
                        //                        "'" + (textBox4.Text) + "' " +
                        //                        ")";
                        //        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();


                        //        var id_sklad = "SELECT max(id) as id_sklad FROM products_jur7";
                        //        sql.myReader = sql.return_MySqlCommand(id_sklad).ExecuteReader();
                        //        while (sql.myReader.Read())
                        //        {
                        //            id_sklad_products = sql.myReader["id_sklad"] != DBNull.Value ? sql.myReader.GetInt32("id_sklad") : 1;
                        //        }
                        //        sql.myReader.Close();

                        //        var insert_product_pri = "insert into products_vnut_per_jur7 (user,year,month,vid_doc,id_sklad_products,kod_doc,product_id,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,ot_kogo_2,komu_1,komu_2) values(" +
                        //                         "'" + string_for_otdels + "'," +
                        //                         "'" + (year_textBox.Text) + "'," +
                        //                         "'" + (month_textBox.Text) + "'," +
                        //                         "'" + ("3") + "'," +
                        //                         "'" + (id_sklad_products) + "'," +
                        //                         "'" + (kod_num_textBox.Text) + "'," +
                        //                        //"'" + (naryad_num_prixod_int) + "'," +//kod_tov orniga
                        //                        "'" + (qrcode_dataGridView.Rows[i].Cells[1].Value != null ? qrcode_dataGridView.Rows[i].Cells[1].Value.ToString() : "") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[2].Value != null ? qrcode_dataGridView.Rows[i].Cells[2].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[3].Value != null ? qrcode_dataGridView.Rows[i].Cells[3].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[4].Value != null ? qrcode_dataGridView.Rows[i].Cells[4].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[5].Value != null ? qrcode_dataGridView.Rows[i].Cells[5].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[6].Value != null ? qrcode_dataGridView.Rows[i].Cells[6].Value.ToString() : "") + "'," +
                        //                         "'" + refresh_strings_to_mysql(qrcode_dataGridView.Rows[i].Cells[7].Value != null ? qrcode_dataGridView.Rows[i].Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                         "'" + refresh_strings_to_mysql(qrcode_dataGridView.Rows[i].Cells[8].Value != null ? qrcode_dataGridView.Rows[i].Cells[8].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                         "'" + refresh_strings_to_mysql(qrcode_dataGridView.Rows[i].Cells[9].Value != null ? qrcode_dataGridView.Rows[i].Cells[9].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[10].Value != null ? qrcode_dataGridView.Rows[i].Cells[10].Value.ToString() : "") + "'," +
                        //                         "'" + (qrcode_dataGridView.Rows[i].Cells[11].Value != null ? qrcode_dataGridView.Rows[i].Cells[11].Value.ToString() : "") + "'," +
                        //                          "'" + (qrcode_dataGridView.Rows[i].Cells[12].Value != null ? qrcode_dataGridView.Rows[i].Cells[12].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[13].Value != null ? qrcode_dataGridView.Rows[i].Cells[13].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[14].Value != null ? qrcode_dataGridView.Rows[i].Cells[14].Value.ToString() : "") + "'," +
                        //                           "'" + (qrcode_dataGridView.Rows[i].Cells[15].Value != null ? qrcode_dataGridView.Rows[i].Cells[15].Value.ToString() : "") + "'," +
                        //                         "'" + refresh_strings_to_mysql(qrcode_dataGridView.Rows[i].Cells[16].Value != null ? qrcode_dataGridView.Rows[i].Cells[16].Value.ToString().Replace(',', '.') : "0") + "'," +
                        //                         "" + (qrcode_dataGridView.Rows[i].Cells[17].Value != null ? "'" + DateTime.Parse(qrcode_dataGridView.Rows[i].Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                        //                        //"'" + (qrcode_dataGridView.Rows[i].Cells[18].Value != null ? qrcode_dataGridView.Rows[i].Cells[18].Value.ToString() : "") + "'," +
                        //                        "'" + (jur_order_qrcode_textBox.Text) + "', " +
                        //                        "'" + (num_qrcode_textBox.Text) + "', " +
                        //                        (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                        //                        "'" + (primech_qrcode_textBox.Text) + "', " +
                        //                        "'" + (doveren_qrcode_textBox.Text) + "', " +
                        //                        "'" + (textBox1.Text) + "', " +
                        //                        "'" + (textBox2.Text) + "', " +
                        //                        "'" + (textBox3.Text) + "', " +
                        //                        "'" + (textBox4.Text) + "' " +
                        //                         ")";
                        //        sql.return_MySqlCommand(insert_product_pri).ExecuteNonQuery();

                        //    }

                        //}


                        label_update_prixod();
                        //sql.return_MySqlCommand("insert into prixod_rasxod (name, prixod,rasxod) values ('" + ot_kogoComboBox.Text + "', '" + (nakladnaydataGridView.Rows[i].Cells[7].Value != null ? nakladnaydataGridView.Rows[i].Cells[7].Value.ToString().Replace(',', '.') : "0") + "','0');").ExecuteNonQuery();
                        MessageBox.Show("Добавлено ", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                    else
                    {
                        DialogResult dialogResult = MessageBox.Show("Обновить данные?", "Обновление", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                        {

                            foreach (DataGridViewRow row in qrcode_dataGridView.Rows)
                            {

                                if (vid_doc == "1")
                                {
                                    if (row.Cells[1].Value != null && row.Cells[0].Value != null && row.Cells[18].Value != null)
                                    {

                                        //kod_tov,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,
                                        //deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,id_products
                                        //jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,komu_1,komu_2

                                        var update_naim = "update naim_tov_jur7 set " +

                                        "naim = '" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                        "kod_gruppa = '" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                        "sena = '" + refresh_strings_to_mysql(row.Cells[8].Value != null ? row.Cells[8].Value.ToString().Replace(",", ".") : "0") + "'," +
                                        "year = '" + year_textBox.Text + "'," +
                                        "month = '" + month_textBox.Text + "'" +
                                       " where id = " + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "";
                                        sql.return_MySqlCommand(update_naim).ExecuteNonQuery();

                                        var update = "update products_jur7 set " +

                                            "user = '" + string_for_otdels + "'," +
                                            "year = '" + year_textBox.Text + "'," +
                                            "month = '" + month_textBox.Text + "'," +
                                            //"vid_doc ='" + "1" + "'," +
                                            "kod_doc ='" + kod_num_textBox.Text + "'," +
                                           "product_id = '" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                            "gruppa = '" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                            "naim_tov = '" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                            "edin = '" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                            "inventar_num = '" + (row.Cells[5].Value != null ? row.Cells[5].Value.ToString() : "") + "'," +
                                            "seria_num = '" + (row.Cells[6].Value != null ? row.Cells[6].Value.ToString() : "") + "'," +
                                            "kol = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'," +
                                            "sena = '" + refresh_strings_to_mysql(row.Cells[8].Value != null ? row.Cells[8].Value.ToString().Replace(",", ".") : "0") + "'," +
                                            "summa = '" + refresh_strings_to_mysql(row.Cells[9].Value != null ? row.Cells[9].Value.ToString().Replace(",", ".") : "0") + "'," +
                                            "deb_sch = '" + (row.Cells[10].Value != null ? row.Cells[10].Value.ToString() : "") + "'," +
                                            "deb_sch_2 = '" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                            "kre_sch = '" + (row.Cells[12].Value != null ? row.Cells[12].Value.ToString() : "") + "'," +
                                            "kre_sch_2 = '" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                            "provodka_iznos = '" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                            "provodka_iznos_2 = '" + (row.Cells[15].Value != null ? row.Cells[15].Value.ToString() : "") + "'," +
                                            "summa_iznos = '" + refresh_strings_to_mysql(row.Cells[16].Value != null ? row.Cells[16].Value.ToString().Replace(",", ".") : "0") + "'," +
                                            "date_pr = '" + (row.Cells[17].Value != null ? DateTime.Parse(row.Cells[17].Value.ToString()).ToString("yyyy-MM-dd") : "NULL") + "', " +
                                            //"date_pr = " + (row.Cells[17].Value != null ? "'" + DateTime.Parse(row.Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                                            // "id_products = '" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                            "jur_order = '" + jur_order_qrcode_textBox.Text + "'," +
                                            "num_doc = '" + num_qrcode_textBox.Text + "'," +
                                            "date_doc = '" + (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") : "NULL") + "'," +
                                            "primech = '" + primech_qrcode_textBox.Text + "'," +
                                            "doverennost = '" + doveren_qrcode_textBox.Text + "'," +
                                            "ot_kogo = '" + textBox1.Text + "'," +
                                            "komu_1 = '" + textBox3.Text + "'," +
                                            "komu_2 = '" + textBox4.Text + "'" +
                                            " where id = " + row.Cells[0].Value + "";
                                        sql.return_MySqlCommand(update).ExecuteNonQuery();

                                        var update2 = "update products_prixod_jur7 set " +

                                           "user = '" + string_for_otdels + "'," +
                                           "year = '" + year_textBox.Text + "'," +
                                           "month = '" + month_textBox.Text + "'," +
                                           "vid_doc ='" + "1" + "'," +
                                           "kod_doc ='" + kod_num_textBox.Text + "'," +
                                          "product_id = '" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                           "gruppa = '" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                           "naim_tov = '" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                           "edin = '" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                           "inventar_num = '" + (row.Cells[5].Value != null ? row.Cells[5].Value.ToString() : "") + "'," +
                                           "seria_num = '" + (row.Cells[6].Value != null ? row.Cells[6].Value.ToString() : "") + "'," +
                                           "kol = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'," +
                                           "sena = '" + refresh_strings_to_mysql(row.Cells[8].Value != null ? row.Cells[8].Value.ToString().Replace(",", ".") : "0") + "'," +
                                           "summa = '" + refresh_strings_to_mysql(row.Cells[9].Value != null ? row.Cells[9].Value.ToString().Replace(",", ".") : "0") + "'," +
                                           "deb_sch = '" + (row.Cells[10].Value != null ? row.Cells[10].Value.ToString() : "") + "'," +
                                           "deb_sch_2 = '" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                           "kre_sch = '" + (row.Cells[12].Value != null ? row.Cells[12].Value.ToString() : "") + "'," +
                                           "kre_sch_2 = '" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                           "provodka_iznos = '" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                           "provodka_iznos_2 = '" + (row.Cells[15].Value != null ? row.Cells[15].Value.ToString() : "") + "'," +
                                           "summa_iznos = '" + refresh_strings_to_mysql(row.Cells[16].Value != null ? row.Cells[16].Value.ToString().Replace(",", ".") : "0") + "'," +
                                           "date_pr = '" + (row.Cells[17].Value != null ? DateTime.Parse(row.Cells[17].Value.ToString()).ToString("yyyy-MM-dd") : "NULL") + "', " +
                                            //"id_products = '" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                            "jur_order = '" + jur_order_qrcode_textBox.Text + "'," +
                                            "num_doc = '" + num_qrcode_textBox.Text + "'," +
                                            "date_doc = '" + (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") : "NULL") + "'," +
                                            "primech = '" + primech_qrcode_textBox.Text + "'," +
                                            "doverennost = '" + doveren_qrcode_textBox.Text + "'," +
                                            "ot_kogo = '" + textBox1.Text + "'," +
                                            "komu_1 = '" + textBox3.Text + "'," +
                                            "komu_2 = '" + textBox4.Text + "'" +
                                           " where id = " + row.Cells[18].Value + "";
                                        sql.return_MySqlCommand(update2).ExecuteNonQuery();


                                    }
                                    if (row.Cells[0].Value == null && row.Cells[7].Value != null)
                                    {
                                        var naim_tov = "insert into naim_tov_jur7 (naim,kod_gruppa,sena,month,year) values(" +
                                               "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                "'" + (row.Cells[8].Value != null ? row.Cells[8].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                "'" + (month_textBox.Text) + "'," +
                                                "'" + (year_textBox.Text) + "'" +
                                               ")";
                                        sql.return_MySqlCommand(naim_tov).ExecuteNonQuery();


                                        var query3 = "SELECT max(id) as product_id FROM naim_tov_jur7";
                                        sql.myReader = sql.return_MySqlCommand(query3).ExecuteReader();
                                        while (sql.myReader.Read())
                                        {
                                            product_id = sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetInt32("product_id") : 1;
                                        }
                                        sql.myReader.Close();



                                        var insert_product = "insert into products_jur7 (user,year,month,vid_doc,kod_doc,product_id,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,komu_1,komu_2) values(" +
                                               "'" + string_for_otdels + "'," +
                                               "'" + (year_textBox.Text) + "'," +
                                               "'" + (month_textBox.Text) + "'," +
                                               "'" + ("1") + "'," +
                                               "'" + (kod_num_textBox.Text) + "'," +
                                                "'" + (product_id) + "'," +
                                               "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[5].Value != null ? row.Cells[5].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[6].Value != null ? row.Cells[6].Value.ToString() : "") + "'," +
                                               "'" + (row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                               "'" + (row.Cells[8].Value != null ? row.Cells[8].Value.ToString().Replace(',', '.') : "0") + "'," +
                                               "'" + (row.Cells[9].Value != null ? row.Cells[9].Value.ToString().Replace(',', '.') : "0") + "'," +
                                               "'" + (row.Cells[10].Value != null ? row.Cells[10].Value.ToString() : "") + "'," +
                                               "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                "'" + (row.Cells[12].Value != null ? row.Cells[12].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[15].Value != null ? row.Cells[15].Value.ToString() : "") + "'," +
                                               "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString().Replace(',', '.') : "0") + "'," +
                                               //"'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                               "'" + (row.Cells[17].Value != null ? DateTime.Parse(row.Cells[17].Value.ToString()).ToString("yyyy-MM-dd") : DateTime.Now.ToString("yyyy-MM-dd")) + "', " +
                                               //"'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                               "'" + (jur_order_qrcode_textBox.Text) + "', " +
                                               "'" + (num_qrcode_textBox.Text) + "', " +
                                               (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                                               "'" + (primech_qrcode_textBox.Text) + "', " +
                                               "'" + (doveren_qrcode_textBox.Text) + "', " +
                                               "'" + (textBox1.Text) + "', " +
                                               "'" + (textBox3.Text) + "', " +
                                               "'" + (textBox4.Text) + "' " +
                                               ")";
                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();

                                        var id_sklad = "SELECT max(id) as id_sklad FROM products_jur7";
                                        sql.myReader = sql.return_MySqlCommand(id_sklad).ExecuteReader();
                                        while (sql.myReader.Read())
                                        {
                                            id_sklad_products = sql.myReader["id_sklad"] != DBNull.Value ? sql.myReader.GetInt32("id_sklad") : 1;
                                        }
                                        sql.myReader.Close();

                                        var insert_product_pri = "insert into products_prixod_jur7 (user,year,month,vid_doc,kod_doc,product_id,id_sklad_products,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,komu_1,komu_2) values(" +
                                                         "'" + string_for_otdels + "'," +
                                                         "'" + (year_textBox.Text) + "'," +
                                                         "'" + (month_textBox.Text) + "'," +
                                                         "'" + ("1") + "'," +
                                                         "'" + (kod_num_textBox.Text) + "'," +
                                                         "'" + (product_id) + "'," +
                                                         "'" + (id_sklad_products) + "'," +
                                                         "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                          "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[5].Value != null ? row.Cells[5].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[6].Value != null ? row.Cells[6].Value.ToString() : "") + "'," +
                                                         "'" + (row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                         "'" + (row.Cells[8].Value != null ? row.Cells[8].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                         "'" + (row.Cells[9].Value != null ? row.Cells[9].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                         "'" + (row.Cells[10].Value != null ? row.Cells[10].Value.ToString() : "") + "'," +
                                                         "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                          "'" + (row.Cells[12].Value != null ? row.Cells[12].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[15].Value != null ? row.Cells[15].Value.ToString() : "") + "'," +
                                                         "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                         //"'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                                         "'" + (row.Cells[17].Value != null ? DateTime.Parse(row.Cells[17].Value.ToString()).ToString("yyyy-MM-dd") : DateTime.Now.ToString("yyyy-MM-dd")) + "', " +
                                                         //"'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                         "'" + (jur_order_qrcode_textBox.Text) + "', " +
                                                       "'" + (num_qrcode_textBox.Text) + "', " +
                                                       (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                                                       "'" + (primech_qrcode_textBox.Text) + "', " +
                                                       "'" + (doveren_qrcode_textBox.Text) + "', " +
                                                       "'" + (textBox1.Text) + "', " +
                                                       "'" + (textBox3.Text) + "', " +
                                                       "'" + (textBox4.Text) + "' " +
                                                         ")";
                                        sql.return_MySqlCommand(insert_product_pri).ExecuteNonQuery();

                                    }
                                }
                                else if (vid_doc == "2")
                                {
                                    if (row.Cells[1].Value != null && row.Cells[0].Value != null && row.Cells[18].Value != null)
                                    {

                                        //kod_tov,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,
                                        //deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,id_products
                                        //jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,komu_1,komu_2

                                        var update = "update products_jur7 set " +

                                            "user = '" + string_for_otdels + "'," +
                                            "year = '" + year_textBox.Text + "'," +
                                            "month = '" + month_textBox.Text + "'," +
                                            "vid_doc ='" + "2" + "'," +
                                            "kod_doc ='" + kod_num_textBox.Text + "'," +
                                           "product_id = '" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                            "gruppa = '" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                            "naim_tov = '" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                            "edin = '" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                            "inventar_num = '" + (row.Cells[5].Value != null ? row.Cells[5].Value.ToString() : "") + "'," +
                                            "seria_num = '" + (row.Cells[6].Value != null ? row.Cells[6].Value.ToString() : "") + "'," +
                                            "kol = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'," +
                                            "sena = '" + refresh_strings_to_mysql(row.Cells[8].Value != null ? row.Cells[8].Value.ToString().Replace(",", ".") : "0") + "'," +
                                            "summa = '" + refresh_strings_to_mysql(row.Cells[9].Value != null ? row.Cells[9].Value.ToString().Replace(",", ".") : "0") + "'," +
                                            "deb_sch = '" + (row.Cells[10].Value != null ? row.Cells[10].Value.ToString() : "") + "'," +
                                            "deb_sch_2 = '" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                            "kre_sch = '" + (row.Cells[12].Value != null ? row.Cells[12].Value.ToString() : "") + "'," +
                                            "kre_sch_2 = '" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                            "provodka_iznos = '" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                            "provodka_iznos_2 = '" + (row.Cells[15].Value != null ? row.Cells[15].Value.ToString() : "") + "'," +
                                            "summa_iznos = '" + refresh_strings_to_mysql(row.Cells[16].Value != null ? row.Cells[16].Value.ToString().Replace(",", ".") : "0") + "'," +
                                            "date_pr = " + (row.Cells[17].Value != null ? "'" + DateTime.Parse(row.Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                                            // "id_products = '" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                            "jur_order = '" + jur_order_qrcode_textBox.Text + "'," +
                                           "num_doc = '" + num_qrcode_textBox.Text + "'," +
                                           "date_doc = '" + (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") : "NULL") + "'," +
                                           "primech = '" + primech_qrcode_textBox.Text + "'," +
                                           "doverennost = '" + doveren_qrcode_textBox.Text + "'," +
                                           "ot_kogo = '" + textBox1.Text + "'," +
                                           "ot_kogo_2 = '" + textBox2.Text + "'," +
                                           "komu_1 = '" + textBox3.Text + "'" +
                                        " where id = " + row.Cells[0].Value + "";
                                        sql.return_MySqlCommand(update).ExecuteNonQuery();

                                        var update2 = "update products_rasxod_jur7 set " +

                                           "user = '" + string_for_otdels + "'," +
                                           "year = '" + year_textBox.Text + "'," +
                                           "month = '" + month_textBox.Text + "'," +
                                           "vid_doc ='" + "2" + "'," +
                                           "kod_doc ='" + kod_num_textBox.Text + "'," +
                                          "product_id = '" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                           "gruppa = '" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                           "naim_tov = '" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                           "edin = '" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                           "inventar_num = '" + (row.Cells[5].Value != null ? row.Cells[5].Value.ToString() : "") + "'," +
                                           "seria_num = '" + (row.Cells[6].Value != null ? row.Cells[6].Value.ToString() : "") + "'," +
                                           "kol = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'," +
                                           "sena = '" + refresh_strings_to_mysql(row.Cells[8].Value != null ? row.Cells[8].Value.ToString().Replace(",", ".") : "0") + "'," +
                                           "summa = '" + refresh_strings_to_mysql(row.Cells[9].Value != null ? row.Cells[9].Value.ToString().Replace(",", ".") : "0") + "'," +
                                           "deb_sch = '" + (row.Cells[10].Value != null ? row.Cells[10].Value.ToString() : "") + "'," +
                                           "deb_sch_2 = '" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                           "kre_sch = '" + (row.Cells[12].Value != null ? row.Cells[12].Value.ToString() : "") + "'," +
                                           "kre_sch_2 = '" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                           "provodka_iznos = '" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                           "provodka_iznos_2 = '" + (row.Cells[15].Value != null ? row.Cells[15].Value.ToString() : "") + "'," +
                                           "summa_iznos = '" + refresh_strings_to_mysql(row.Cells[16].Value != null ? row.Cells[16].Value.ToString().Replace(",", ".") : "0") + "'," +
                                          "date_pr = " + (row.Cells[17].Value != null ? "'" + DateTime.Parse(row.Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                                           //"id_products = '" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                           "jur_order = '" + jur_order_qrcode_textBox.Text + "'," +
                                           "num_doc = '" + num_qrcode_textBox.Text + "'," +
                                           "date_doc = '" + (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") : "NULL") + "'," +
                                           "primech = '" + primech_qrcode_textBox.Text + "'," +
                                           "doverennost = '" + doveren_qrcode_textBox.Text + "'," +
                                           "ot_kogo = '" + textBox1.Text + "'," +
                                           "ot_kogo_2 = '" + textBox2.Text + "'," +
                                           "komu_1 = '" + textBox3.Text + "'" +
                                           " where id = " + row.Cells[18].Value + "";
                                        sql.return_MySqlCommand(update2).ExecuteNonQuery();


                                    }
                                    if (row.Cells[0].Value == null && row.Cells[7].Value != null)
                                    {


                                        var insert_product = "insert into products_jur7 (user,year,month,vid_doc,kod_doc,product_id,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,ot_kogo_2,komu_1) values(" +
                                               "'" + string_for_otdels + "'," +
                                               "'" + (year_textBox.Text) + "'," +
                                               "'" + (month_textBox.Text) + "'," +
                                               "'" + ("2") + "'," +
                                               "'" + (kod_num_textBox.Text) + "'," +
                                               //"'" + (product_id) + "'," +//kod_tov orniga
                                               "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                               "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[5].Value != null ? row.Cells[5].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[6].Value != null ? row.Cells[6].Value.ToString() : "") + "'," +
                                               "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                               "'" + refresh_strings_to_mysql(row.Cells[8].Value != null ? row.Cells[8].Value.ToString().Replace(',', '.') : "0") + "'," +
                                               "'" + refresh_strings_to_mysql(row.Cells[9].Value != null ? row.Cells[9].Value.ToString().Replace(',', '.') : "0") + "'," +
                                               "'" + (row.Cells[10].Value != null ? row.Cells[10].Value.ToString() : "") + "'," +
                                               "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                "'" + (row.Cells[12].Value != null ? row.Cells[12].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[15].Value != null ? row.Cells[15].Value.ToString() : "") + "'," +
                                               "'" + refresh_strings_to_mysql(row.Cells[16].Value != null ? row.Cells[16].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                "" + (row.Cells[17].Value != null ? "'" + DateTime.Parse(row.Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                                                 //"'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                 "'" + (jur_order_qrcode_textBox.Text) + "', " +
                                                         "'" + (num_qrcode_textBox.Text) + "', " +
                                                         (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                                                         "'" + (primech_qrcode_textBox.Text) + "', " +
                                                         "'" + (doveren_qrcode_textBox.Text) + "', " +
                                                         "'" + (textBox1.Text) + "', " +
                                                         "'" + (textBox2.Text) + "', " +
                                                         "'" + (textBox3.Text) + "' " +
                                               ")";
                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();

                                        var id_sklad = "SELECT max(id) as id_sklad FROM products_jur7";
                                        sql.myReader = sql.return_MySqlCommand(id_sklad).ExecuteReader();
                                        while (sql.myReader.Read())
                                        {
                                            id_sklad_products = sql.myReader["id_sklad"] != DBNull.Value ? sql.myReader.GetInt32("id_sklad") : 1;
                                        }
                                        sql.myReader.Close();

                                        var insert_product_pri = "insert into products_rasxod_jur7 (user,year,month,vid_doc,id_sklad_products,kod_doc,product_id,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,ot_kogo_2,komu_1) values(" +
                                                         "'" + string_for_otdels + "'," +
                                                         "'" + (year_textBox.Text) + "'," +
                                                         "'" + (month_textBox.Text) + "'," +
                                                         "'" + ("2") + "'," +
                                                         "'" + (id_sklad_products) + "'," +
                                                         "'" + (kod_num_textBox.Text) + "'," +
                                                         //"'" + (product_id) + "'," +//kod_tov orniga
                                                         "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                         "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                          "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[5].Value != null ? row.Cells[5].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[6].Value != null ? row.Cells[6].Value.ToString() : "") + "'," +
                                                         "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                         "'" + refresh_strings_to_mysql(row.Cells[8].Value != null ? row.Cells[8].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                         "'" + refresh_strings_to_mysql(row.Cells[9].Value != null ? row.Cells[9].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                         "'" + (row.Cells[10].Value != null ? row.Cells[10].Value.ToString() : "") + "'," +
                                                         "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                          "'" + (row.Cells[12].Value != null ? row.Cells[12].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[15].Value != null ? row.Cells[15].Value.ToString() : "") + "'," +
                                                         "'" + refresh_strings_to_mysql(row.Cells[16].Value != null ? row.Cells[16].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                          "" + (row.Cells[17].Value != null ? "'" + DateTime.Parse(row.Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                                                         //"'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                         "'" + (jur_order_qrcode_textBox.Text) + "', " +
                                                         "'" + (num_qrcode_textBox.Text) + "', " +
                                                         (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                                                         "'" + (primech_qrcode_textBox.Text) + "', " +
                                                         "'" + (doveren_qrcode_textBox.Text) + "', " +
                                                         "'" + (textBox1.Text) + "', " +
                                                         "'" + (textBox2.Text) + "', " +
                                                         "'" + (textBox3.Text) + "' " +
                                                         ")";
                                        sql.return_MySqlCommand(insert_product_pri).ExecuteNonQuery();

                                    }
                                }
                                else if (vid_doc == "3")
                                {
                                    if (row.Cells[0].Value != null && row.Cells[18].Value != null)
                                    {

                                        var update = "update products_jur7 set " +

                                            "user = '" + string_for_otdels + "'," +
                                            "year = '" + year_textBox.Text + "'," +
                                            "month = '" + month_textBox.Text + "'," +
                                            "vid_doc ='" + "3" + "'," +
                                            "kod_doc ='" + kod_num_textBox.Text + "'," +
                                            //"product_id = '" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                            "gruppa = '" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                            "naim_tov = '" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                            "edin = '" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                            "inventar_num = '" + (row.Cells[5].Value != null ? row.Cells[5].Value.ToString() : "") + "'," +
                                            "seria_num = '" + (row.Cells[6].Value != null ? row.Cells[6].Value.ToString() : "") + "'," +
                                            "kol = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'," +
                                            "sena = '" + refresh_strings_to_mysql(row.Cells[8].Value != null ? row.Cells[8].Value.ToString().Replace(",", ".") : "0") + "'," +
                                            "summa = '" + refresh_strings_to_mysql(row.Cells[9].Value != null ? row.Cells[9].Value.ToString().Replace(",", ".") : "0") + "'," +
                                            "deb_sch = '" + (row.Cells[10].Value != null ? row.Cells[10].Value.ToString() : "") + "'," +
                                            "deb_sch_2 = '" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                            "kre_sch = '" + (row.Cells[12].Value != null ? row.Cells[12].Value.ToString() : "") + "'," +
                                            "kre_sch_2 = '" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                            "provodka_iznos = '" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                            "provodka_iznos_2 = '" + (row.Cells[15].Value != null ? row.Cells[15].Value.ToString() : "") + "'," +
                                            "summa_iznos = '" + refresh_strings_to_mysql(row.Cells[16].Value != null ? row.Cells[16].Value.ToString().Replace(",", ".") : "0") + "'," +
                                            "date_pr = " + (row.Cells[17].Value != null ? "'" + DateTime.Parse(row.Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                                            // "id_products = '" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                            "jur_order = '" + jur_order_qrcode_textBox.Text + "'," +
                                           "num_doc = '" + num_qrcode_textBox.Text + "'," +
                                           "date_doc = '" + (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") : "NULL") + "'," +
                                           "primech = '" + primech_qrcode_textBox.Text + "'," +
                                           "doverennost = '" + doveren_qrcode_textBox.Text + "'," +
                                           "ot_kogo = '" + textBox1.Text + "'," +
                                           "ot_kogo_2 = '" + textBox2.Text + "'," +
                                           "komu_1 = '" + textBox3.Text + "'," +
                                           "komu_2 = '" + textBox4.Text + "'" +
                                            " where id = " + row.Cells[0].Value + "";
                                        sql.return_MySqlCommand(update).ExecuteNonQuery();

                                        var update2 = "update products_vnut_per_jur7 set " +

                                          "user = '" + string_for_otdels + "'," +
                                            "year = '" + year_textBox.Text + "'," +
                                            "month = '" + month_textBox.Text + "'," +
                                            "vid_doc ='" + "3" + "'," +
                                            "kod_doc ='" + kod_num_textBox.Text + "'," +
                                           //"product_id = '" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                           "gruppa = '" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                           "naim_tov = '" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                           "edin = '" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                           "inventar_num = '" + (row.Cells[5].Value != null ? row.Cells[5].Value.ToString() : "") + "'," +
                                           "seria_num = '" + (row.Cells[6].Value != null ? row.Cells[6].Value.ToString() : "") + "'," +
                                           "kol = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'," +
                                           "sena = '" + refresh_strings_to_mysql(row.Cells[8].Value != null ? row.Cells[8].Value.ToString().Replace(",", ".") : "0") + "'," +
                                           "summa = '" + refresh_strings_to_mysql(row.Cells[9].Value != null ? row.Cells[9].Value.ToString().Replace(",", ".") : "0") + "'," +
                                           "deb_sch = '" + (row.Cells[10].Value != null ? row.Cells[10].Value.ToString() : "") + "'," +
                                           "deb_sch_2 = '" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                           "kre_sch = '" + (row.Cells[12].Value != null ? row.Cells[12].Value.ToString() : "") + "'," +
                                           "kre_sch_2 = '" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                           "provodka_iznos = '" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                           "provodka_iznos_2 = '" + (row.Cells[15].Value != null ? row.Cells[15].Value.ToString() : "") + "'," +
                                           "summa_iznos = '" + refresh_strings_to_mysql(row.Cells[16].Value != null ? row.Cells[16].Value.ToString().Replace(",", ".") : "0") + "'," +
                                            "date_pr = " + (row.Cells[17].Value != null ? "'" + DateTime.Parse(row.Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                                           //"id_products = '" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                           "jur_order = '" + jur_order_qrcode_textBox.Text + "'," +
                                           "num_doc = '" + num_qrcode_textBox.Text + "'," +
                                           "date_doc = '" + (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") : "NULL") + "'," +
                                           "primech = '" + primech_qrcode_textBox.Text + "'," +
                                           "doverennost = '" + doveren_qrcode_textBox.Text + "'," +
                                           "ot_kogo = '" + textBox1.Text + "'," +
                                           "ot_kogo_2 = '" + textBox2.Text + "'," +
                                           "komu_1 = '" + textBox3.Text + "'," +
                                           "komu_2 = '" + textBox4.Text + "'" +
                                           " where id = " + row.Cells[18].Value + "";
                                        sql.return_MySqlCommand(update2).ExecuteNonQuery();


                                    }
                                    if (row.Cells[0].Value == null && row.Cells[7].Value != null)
                                    {


                                        var insert_product = "insert into products_jur7 (user,year,month,vid_doc,kod_doc,product_id,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,ot_kogo_2,komu_1,komu_2) values(" +
                                               "'" + string_for_otdels + "'," +
                                               "'" + (year_textBox.Text) + "'," +
                                               "'" + (month_textBox.Text) + "'," +
                                               "'" + ("3") + "'," +
                                               "'" + (kod_num_textBox.Text) + "'," +
                                               //"'" + (product_id) + "'," +//kod_tov orniga
                                               //"'" + (product_id) + "'," +
                                               "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                               "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[5].Value != null ? row.Cells[5].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[6].Value != null ? row.Cells[6].Value.ToString() : "") + "'," +
                                               "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                               "'" + refresh_strings_to_mysql(row.Cells[8].Value != null ? row.Cells[8].Value.ToString().Replace(',', '.') : "0") + "'," +
                                               "'" + refresh_strings_to_mysql(row.Cells[9].Value != null ? row.Cells[9].Value.ToString().Replace(',', '.') : "0") + "'," +
                                               "'" + (row.Cells[10].Value != null ? row.Cells[10].Value.ToString() : "") + "'," +
                                               "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                "'" + (row.Cells[12].Value != null ? row.Cells[12].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                 "'" + (row.Cells[15].Value != null ? row.Cells[15].Value.ToString() : "") + "'," +
                                               "'" + refresh_strings_to_mysql(row.Cells[16].Value != null ? row.Cells[16].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                "" + (row.Cells[17].Value != null ? "'" + DateTime.Parse(row.Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                                               //"'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                               "'" + (jur_order_qrcode_textBox.Text) + "', " +
                                                         "'" + (num_qrcode_textBox.Text) + "', " +
                                                         (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                                                         "'" + (primech_qrcode_textBox.Text) + "', " +
                                                         "'" + (doveren_qrcode_textBox.Text) + "', " +
                                                         "'" + (textBox1.Text) + "', " +
                                                         "'" + (textBox2.Text) + "', " +
                                                         "'" + (textBox3.Text) + "', " +
                                                         "'" + (textBox4.Text) + "' " +
                                               ")";
                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();

                                        var id_sklad = "SELECT max(id) as id_sklad FROM products_jur7";
                                        sql.myReader = sql.return_MySqlCommand(id_sklad).ExecuteReader();
                                        while (sql.myReader.Read())
                                        {
                                            id_sklad_products = sql.myReader["id_sklad"] != DBNull.Value ? sql.myReader.GetInt32("id_sklad") : 1;
                                        }
                                        sql.myReader.Close();

                                        var insert_product_pri = "insert into products_vnut_per_jur7 (user,year,month,vid_doc,id_sklad_products,kod_doc,product_id,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,ot_kogo_2,komu_1,komu_2) values(" +
                                                         "'" + string_for_otdels + "'," +
                                                         "'" + (year_textBox.Text) + "'," +
                                                         "'" + (month_textBox.Text) + "'," +
                                                         "'" + ("3") + "'," +
                                                         "'" + (id_sklad_products) + "'," +
                                                         "'" + (kod_num_textBox.Text) + "'," +
                                                         //"'" + (product_id) + "'," +//kod_tov orniga
                                                         "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                         "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                          "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[5].Value != null ? row.Cells[5].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[6].Value != null ? row.Cells[6].Value.ToString() : "") + "'," +
                                                         "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                         "'" + refresh_strings_to_mysql(row.Cells[8].Value != null ? row.Cells[8].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                         "'" + refresh_strings_to_mysql(row.Cells[9].Value != null ? row.Cells[9].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                         "'" + (row.Cells[10].Value != null ? row.Cells[10].Value.ToString() : "") + "'," +
                                                         "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                          "'" + (row.Cells[12].Value != null ? row.Cells[12].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                           "'" + (row.Cells[15].Value != null ? row.Cells[15].Value.ToString() : "") + "'," +
                                                         "'" + refresh_strings_to_mysql(row.Cells[16].Value != null ? row.Cells[16].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                          "" + (row.Cells[17].Value != null ? "'" + DateTime.Parse(row.Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                                                         //"'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                         "'" + (jur_order_qrcode_textBox.Text) + "', " +
                                                         "'" + (num_qrcode_textBox.Text) + "', " +
                                                         (data_qrcode_DateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(data_qrcode_DateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                                                         "'" + (primech_qrcode_textBox.Text) + "', " +
                                                         "'" + (doveren_qrcode_textBox.Text) + "', " +
                                                         "'" + (textBox1.Text) + "', " +
                                                         "'" + (textBox2.Text) + "', " +
                                                         "'" + (textBox3.Text) + "', " +
                                                         "'" + (textBox4.Text) + "' " +
                                                         ")";
                                        sql.return_MySqlCommand(insert_product_pri).ExecuteNonQuery();

                                    }
                                }


                            }
                            MessageBox.Show("Обновлено");
                        }
                        else if (dialogResult == DialogResult.No)
                        {
                            //MessageBox.Show("ma'lumot to'liq emas");
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("save_button_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void qrcode_dataGridView_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            Console.WriteLine("autocomplite");

            if (qrcode_dataGridView.CurrentCell.ColumnIndex == 3)
            {
                TextBox auto_text = e.Control as TextBox;

                if (auto_text != null)
                {
                    auto_text.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                    auto_text.AutoCompleteSource = AutoCompleteSource.CustomSource;
                    // AutoCompleteStringCollection sc = new AutoCompleteStringCollection();

                    auto_text.AutoCompleteCustomSource = column;
                }
            }
            //else if (prixod_dataGridView.CurrentCell.ColumnIndex == 1)
            //{
            //    TextBox auto_text = e.Control as TextBox;

            //    if (auto_text != null)
            //    {
            //        auto_text.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            //        auto_text.AutoCompleteSource = AutoCompleteSource.CustomSource;
            //        // AutoCompleteStringCollection sc = new AutoCompleteStringCollection();

            //        auto_text.AutoCompleteCustomSource = column_kart_num;
            //    }
            //}
            else
            {
                TextBox auto_text = e.Control as TextBox;
                if (auto_text != null)
                {
                    auto_text.AutoCompleteMode = AutoCompleteMode.None;
                }
            }
        }

        AutoCompleteStringCollection column = new AutoCompleteStringCollection();
        //AutoCompleteStringCollection column_kart_num = new AutoCompleteStringCollection();


        public void add_items()
        {
            try
            {
                var auto = " SELECT distinct naim FROM naim_tov_jur7 ";
                sql.myReader = sql.return_MySqlCommand(auto).ExecuteReader();
                while (sql.myReader.Read())
                {
                    column.Add(sql.myReader.GetString("naim"));
                }
                sql.myReader.Close();
                //sql.myReader = sql.return_MySqlCommand("select kart_num from sklad where otdel = '" + otdel_name + "' ").ExecuteReader();
                //while (sql.myReader.Read())
                //{
                //    column_kart_num.Add(sql.myReader.GetString("kart_num"));
                //}
                //sql.myReader.Close();
            }
            catch (Exception ex)
            {
                sql.myReader.Close();
                MessageBox.Show("add_items " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

    }
}
