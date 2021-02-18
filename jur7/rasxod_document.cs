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
    public partial class rasxod_document : Form
    {
        Connect sql = new Connect();
        Connect sql2 = new Connect();
        Connect sql3 = new Connect();

        public string string_for_otdels;
        public string month_global;
        public string year_global;
        public string iznos;

        Number_To_Words_russian number_russian = new Number_To_Words_russian();


        public rasxod_document(string string_for_otdels, string year_global, string month_global, string iznos)
        {
            InitializeComponent();
            sql.Connection();
            sql2.Connection();
            sql3.Connection();

            this.string_for_otdels = string_for_otdels;
            this.year_global = year_global;
            this.month_global = month_global;
            this.iznos = iznos;

            run_main();

            komu_combo();
        }

        //DataTable multi_col = new DataTable();
        public void komu_combo()
        {
            //komu_ras_comboBox.Items.Clear();
            //sql.myReader = sql.return_MySqlCommand("SELECT distinct podraz FROM podrazdeleniya where type='Кому'").ExecuteReader();

            //while (sql.myReader.Read())
            //{
            //    komu_ras_comboBox.Items.Add(sql.myReader.GetString("podraz"));
            //}
            //sql.myReader.Close();

            //txt_edin_prixod.Items.Clear();
            //txt_edin_prixod.Items.Add("шт");
            //txt_edin_prixod.Items.Add("кт");
            //txt_edin_prixod.Items.Add("метр");
            //txt_edin_prixod.Items.Add("м");
            //txt_edin_prixod.Items.Add("м2");
            //txt_edin_prixod.Items.Add("комп");
            //txt_edin_prixod.Items.Add("гол");
            //txt_edin_prixod.Items.Add("уп");
            //txt_edin_prixod.Items.Add("пара");
            //txt_edin_prixod.Items.Add("литр");
            //txt_edin_prixod.Items.Add("пач");
            //txt_edin_prixod.Items.Add("экз");
            //txt_edin_prixod.Items.Add("пар");


            ot_kogo_ras_ComboBox.Items.Clear();
            sql.myReader = sql.return_MySqlCommand("SELECT distinct podraz_naim FROM podraz_jur7 ").ExecuteReader();

            while (sql.myReader.Read())
            {
                ot_kogo_ras_ComboBox.Items.Add(sql.myReader.GetString("podraz_naim"));
            }
            sql.myReader.Close();

            //DataGridViewMultiColumnComboColumn newColumn = (DataGridViewMultiColumnComboColumn)rasxod_dataGridView.Columns[2];

            //sql.mydataAdapter = new MySqlDataAdapter();
            //multi_col.Clear();
            //sql.mydataAdapter.SelectCommand = this.sql.return_MySqlCommand(" SELECT kod_gruppa,schet,naim FROM gruppa order by naim asc");
            //sql.mydataAdapter.Fill(multi_col);

            //newColumn.DataSource = multi_col;

            //newColumn.DropDownWidth = 600;
            //newColumn.Width = 100;

            //newColumn.DataPropertyName = "kod_gruppa";
            //newColumn.DisplayMember = "kod_gruppa";
            //newColumn.ValueMember = "kod_gruppa";
            //newColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
        }

        int naryad_num_prixod_int;
        int scetcik_num_prixod_int;
        DataTable rasxod_datatable = new DataTable();
        public void run_main()
        {
            try
            {
                //txt_edin_prixod.Items.Clear();
                //sql.myReader = sql.return_MySqlCommand("SELECT distinct edin FROM products where edin is not null").ExecuteReader();

                //while (sql.myReader.Read())
                //{
                //    txt_edin_prixod.Items.Add(sql.myReader.GetString("edin"));
                //}
                //sql.myReader.Close();

                var select_akt = "select max(kod_doc) as max_kod_doc from products_rasxod_jur7 where vid_doc='2' and year='" + year_global + "' and month='" + month_global + "' and user='" + string_for_otdels + "' ORDER BY kod_doc desc";
                sql.myReader = sql.return_MySqlCommand(select_akt).ExecuteReader();
                while (sql.myReader.Read())
                {
                    naryad_num_prixod_int = sql.myReader["max_kod_doc"] != DBNull.Value ? sql.myReader.GetInt32("max_kod_doc") : 1;
                }
                sql.myReader.Close();

                set_items_to_values_prixod();
                var query3 = "select max(kod_doc) as max_kod_doc from products_rasxod_jur7 where vid_doc='2' and year='" + year_global + "' and month='" + month_global + "' and user='" + string_for_otdels + "' ORDER BY kod_doc desc";
                sql.myReader = sql.return_MySqlCommand(query3).ExecuteReader();
                scetcik_num_prixod_int = 0;
                while (sql.myReader.Read())
                {
                    scetcik_num_prixod_int = sql.myReader["max_kod_doc"] != DBNull.Value ? sql.myReader.GetInt32("max_kod_doc") : 1;
                }
                sql.myReader.Close();



                scetcik_num_prixod_int++;
                label30.Text = Convert.ToString(naryad_num_prixod_int) + " от " + Convert.ToString(scetcik_num_prixod_int);

                kod_num_textBox.Text = Convert.ToString(naryad_num_prixod_int);


                //sql.mydataAdapter = new MySqlDataAdapter();
                //sql.mydataAdapter.SelectCommand = sql.return_MySqlCommand("SELECT id,kod_gruppa,naim,schet FROM gruppa;");
                //DataSet dtbl = new DataSet();
                //sql.mydataAdapter.Fill(dtbl);
                //for (int i = 0; i < dtbl.Tables[0].Rows.Count; i++)
                //{
                //    //txt_gruppa_rasxod.Items.Add(dtbl.Tables[0].Rows[i][1] + "| " + dtbl.Tables[0].Rows[i][2] + "| " + dtbl.Tables[0].Rows[i][3]);
                //    //komu_ras_comboBox.Columns[dtbl.Tables[0].Rows[i][1]].Width = 300;
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show("run_treeview " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        public void set_items_to_values_prixod()
        {
            try
            {
                this.rasxod_dataGridView.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.rasxod_dataGridView_CellValueChanged);

                string debet_01 = "";
                double debet_01_sum = 0;
                double debet_06_sum = 0;
                double debet_07_sum = 0;

                jur_order_ras_textBox.Text = "";
                num_ras_textBox.Text = "";
                primech_ras_textBox.Text = "";
                doveren_ras_textBox.Text = "";
                ot_kogo_ras_ComboBox.Text = "";
                ot_kogo_ras_comboBox2.Text = "";
                komu_ras_comboBox.Text = "";

                var query = "SELECT id,jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,ot_kogo_2,komu_1,year,month FROM products_rasxod_jur7 where vid_doc='2' and year='" + year_global + "' and month='" + month_global + "' and kod_doc='" + naryad_num_prixod_int + "' and user='" + string_for_otdels + "' order by kod_doc";
                sql.myReader = sql.return_MySqlCommand(query).ExecuteReader();
                while (sql.myReader.Read())
                {
                    jur_order_ras_textBox.Text = (sql.myReader["jur_order"] != DBNull.Value ? sql.myReader.GetString("jur_order") : "");
                    num_ras_textBox.Text = (sql.myReader["num_doc"] != DBNull.Value ? sql.myReader.GetString("num_doc") : "");
                    data_ras_DateTimePicker.Value = (sql.myReader["date_doc"] != DBNull.Value ? sql.myReader.GetDateTime("date_doc") : DateTime.Now);
                    primech_ras_textBox.Text = (sql.myReader["primech"] != DBNull.Value ? sql.myReader.GetString("primech") : "");
                    doveren_ras_textBox.Text = (sql.myReader["doverennost"] != DBNull.Value ? sql.myReader.GetString("doverennost") : "");
                    ot_kogo_ras_ComboBox.Text = (sql.myReader["ot_kogo"] != DBNull.Value ? sql.myReader.GetString("ot_kogo") : "");
                    ot_kogo_ras_comboBox2.Text = (sql.myReader["ot_kogo_2"] != DBNull.Value ? sql.myReader.GetString("ot_kogo_2") : "");
                    komu_ras_comboBox.Text = (sql.myReader["komu_1"] != DBNull.Value ? sql.myReader.GetString("komu_1") : "");
                }
                sql.myReader.Close();

                rasxod_dataGridView.Rows.Clear();
                var select_ras = "SELECT * FROM products_rasxod_jur7 where vid_doc='2' and year='" + year_global + "' and month='" + month_global + "' and kod_doc='" + naryad_num_prixod_int + "' and user='" + string_for_otdels + "' ";
                sql.myReader = sql.return_MySqlCommand(select_ras).ExecuteReader();
                while (sql.myReader.Read())
                {

                    //kod_tov,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,

                    int index = rasxod_dataGridView.Rows.Add();
                    rasxod_dataGridView.Rows[index].Cells[0].Value = (sql.myReader["id"] != DBNull.Value ? sql.myReader.GetString("id") : "");
                    rasxod_dataGridView.Rows[index].Cells[1].Value = (sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetString("product_id") : "");
                    rasxod_dataGridView.Rows[index].Cells[2].Value = (sql.myReader["gruppa"] != DBNull.Value ? sql.myReader.GetString("gruppa") : "");
                    rasxod_dataGridView.Rows[index].Cells[3].Value = (sql.myReader["naim_tov"] != DBNull.Value ? sql.myReader.GetString("naim_tov") : "");
                    rasxod_dataGridView.Rows[index].Cells[4].Value = (sql.myReader["edin"] != DBNull.Value ? sql.myReader.GetString("edin") : "");
                    rasxod_dataGridView.Rows[index].Cells[5].Value = (sql.myReader["inventar_num"] != DBNull.Value ? sql.myReader.GetString("inventar_num") : "");
                    rasxod_dataGridView.Rows[index].Cells[6].Value = (sql.myReader["seria_num"] != DBNull.Value ? sql.myReader.GetString("seria_num") : "");
                    string kols = sql.myReader["kol"] != DBNull.Value ? sql.myReader.GetString("kol") : "";

                    if (kols.Length <= 3)
                    {
                        rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[7].Value = string.Format("{0:#0.00}", (sql.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("kol").Replace(".", ","))) : 0));
                    }
                    if (kols.Length > 3)
                    {
                        rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[7].Value = string.Format("{0:#,###.00}", (sql.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("kol").Replace(".", ","))) : 0));
                    }

                    string sena = sql.myReader["sena"] != DBNull.Value ? sql.myReader.GetString("sena") : "";

                    if (sena.Length <= 3)
                    {
                        rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[8].Value = string.Format("{0:#0.00}", (sql.myReader["sena"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("sena").Replace(".", ","))) : 0));
                    }
                    if (sena.Length > 3)
                    {
                        rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[8].Value = string.Format("{0:#,###.00}", (sql.myReader["sena"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("sena").Replace(".", ","))) : 0));
                    }

                    string summa = sql.myReader["summa"] != DBNull.Value ? sql.myReader.GetString("summa") : "";

                    if (summa.Length <= 3)
                    {
                        rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[9].Value = string.Format("{0:#0.00}", (sql.myReader["summa"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa").Replace(".", ","))) : 0));
                    }
                    if (summa.Length > 3)
                    {
                        rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[9].Value = string.Format("{0:#,###.00}", (sql.myReader["summa"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa").Replace(".", ","))) : 0));
                    }

                    //deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,id_products

                    rasxod_dataGridView.Rows[index].Cells[10].Value = (sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch") : "");


                    debet_01 = (sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch") : "");
                    string first = debet_01.Substring(0, 2);

                    if (first == "01")
                    {
                        debet_01_sum += (rasxod_dataGridView.Rows[index].Cells[9].Value != null ? Double.Parse(rasxod_dataGridView.Rows[index].Cells[9].Value.ToString()) : 0);
                    }
                    else if (first == "06")
                    {
                        debet_06_sum += (rasxod_dataGridView.Rows[index].Cells[9].Value != null ? Double.Parse(rasxod_dataGridView.Rows[index].Cells[9].Value.ToString()) : 0);
                    }
                    else if (first == "07")
                    {
                        debet_07_sum += (rasxod_dataGridView.Rows[index].Cells[9].Value != null ? Double.Parse(rasxod_dataGridView.Rows[index].Cells[9].Value.ToString()) : 0);
                    }

                    rasxod_dataGridView.Rows[index].Cells[11].Value = (sql.myReader["deb_sch_2"] != DBNull.Value ? sql.myReader.GetString("deb_sch_2") : "");
                    rasxod_dataGridView.Rows[index].Cells[12].Value = (sql.myReader["kre_sch"] != DBNull.Value ? sql.myReader.GetString("kre_sch") : "");
                    rasxod_dataGridView.Rows[index].Cells[13].Value = (sql.myReader["kre_sch_2"] != DBNull.Value ? sql.myReader.GetString("kre_sch_2") : "");
                    rasxod_dataGridView.Rows[index].Cells[14].Value = (sql.myReader["provodka_iznos"] != DBNull.Value ? sql.myReader.GetString("provodka_iznos") : "");
                    rasxod_dataGridView.Rows[index].Cells[15].Value = (sql.myReader["provodka_iznos_2"] != DBNull.Value ? sql.myReader.GetString("provodka_iznos_2") : "");

                    string summa_iznos = sql.myReader["summa_iznos"] != DBNull.Value ? sql.myReader.GetString("summa_iznos") : "";

                    if (summa_iznos.Length <= 3)
                    {
                        rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[16].Value = string.Format("{0:#0.00}", (sql.myReader["summa_iznos"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa_iznos").Replace(".", ","))) : 0));
                    }
                    if (summa_iznos.Length > 3)
                    {
                        rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[16].Value = string.Format("{0:#,###.00}", (sql.myReader["summa_iznos"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa_iznos").Replace(".", ","))) : 0));
                    }

                    rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[17].Value = (sql.myReader["date_pr"] != DBNull.Value ? (DateTime.Parse(sql.myReader.GetString("date_pr")).ToString("dd.MM.yyyy")) : null);
                    rasxod_dataGridView.Rows[index].Cells[18].Value = (sql.myReader["id_sklad_products"] != DBNull.Value ? sql.myReader.GetString("id_sklad_products") : "");
                    //sklad_dataGridView.Rows[index].Cells[3].Value = refresh_strings_to_mysql(sql.myReader["sena"] != DBNull.Value ? string.Format("{0:#0.00}", sql.myReader.GetDouble("sena")) : "0");
                    //rasxod_dataGridView.Rows[index].Cells[18].Value = (sql.myReader["id_products"] != DBNull.Value ? sql.myReader.GetString("id_products") : "");


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

                label30.Text = Convert.ToString(naryad_num_prixod_int) + " от " + Convert.ToString(scetcik_num_prixod_int);
                kod_num_textBox.Text = naryad_num_prixod_int.ToString();
                this.rasxod_dataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.rasxod_dataGridView_CellValueChanged);
                label_update_prixod();
            }
            catch (Exception ex)
            {
                //sql.myReader.Close();
                MessageBox.Show("Хато маълумот киритилган (" + ex.Message + ")");
            }
        }


        public void label_update_prixod()
        {
            double summa = 0;
            double iznos = 0;


            foreach (DataGridViewRow row in rasxod_dataGridView.Rows)
            {
                summa = summa + (row.Cells[9].Value != null ? Double.Parse(row.Cells[9].Value.ToString()) : 0);

                iznos = iznos + (row.Cells[16].Value != null ? Double.Parse(row.Cells[16].Value.ToString()) : 0);

            }
            if (summa.ToString().Length <= 3)
            {
                rasxod_obshiy_summa_label.Text = string.Format("{0:#0.00}", summa);
            }
            if (summa.ToString().Length > 3)
            {
                rasxod_obshiy_summa_label.Text = string.Format("{0:#0,000.00}", summa);
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


        private void rasxod_dataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                //DataGridViewRow row = rasxod_dataGridView.CurrentRow;
                //Gruppa gruppa = new Gruppa();
                //if (e.ColumnIndex==2)
                //{
                //    if (gruppa.ShowDialog() == DialogResult.OK)
                //    {
                //        row.Cells[2].Value = gruppa.kod_gruppa;
                //        row.Cells[3].Value = gruppa.naim;
                //        row.Cells[10].Value = gruppa.schet;
                //        row.Cells[11].Value = gruppa.debet;
                //        row.Cells[12].Value = gruppa.kredit;

                //    }
                //}

                string ot_kogo_1 = ot_kogo_ras_ComboBox.Text;
                string ot_kogo_2 = ot_kogo_ras_comboBox2.Text;
                string pereosenka_visible = "0";
                //if (ot_kogo_ras_ComboBox.Text != "" && ot_kogo_ras_comboBox2.Text != "")
                //{
                if (e.ColumnIndex == 3)
                {

                    this.rasxod_dataGridView.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.rasxod_dataGridView_CellValueChanged);
                    DataGridViewRow row_ost = rasxod_dataGridView.CurrentRow;
                    Sklad ostatok = new Sklad(string_for_otdels, year_global, month_global, ot_kogo_1, ot_kogo_2, pereosenka_visible);

                    ostatok.WindowState = FormWindowState.Maximized;


                    if (ostatok.ShowDialog() == DialogResult.OK)
                    {
                        this.rasxod_dataGridView.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.rasxod_dataGridView_CellValueChanged);
                        //int index = 0;
                        DataTable table = ostatok.table;


                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            int index = rasxod_dataGridView.Rows.Add();
                            //rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Add()].Cells[0].Value = table.Rows[i][0];
                            rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[1].Value = table.Rows[i][1];
                            rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[2].Value = table.Rows[i][11];
                            rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[3].Value = table.Rows[i][3];
                            rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[4].Value = table.Rows[i][4];
                            rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[5].Value = table.Rows[i][14];
                            rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[6].Value = table.Rows[i][13];//string.Format("{0:#0.00}", table.Rows[i][7]);
                            rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[7].Value = string.Format("{0:#0.00}", table.Rows[i][9]);
                            rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[8].Value = string.Format("{0:#0.00}", table.Rows[i][6]);
                            rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[9].Value = string.Format("{0:#0.00}", table.Rows[i][10]);
                            rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[10].Value = table.Rows[i][2];
                            rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[11].Value = table.Rows[i][16];
                            rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[12].Value = table.Rows[i][17];
                            rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[13].Value = table.Rows[i][18];
                            rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[16].Value = table.Rows[i][8];

                            rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[17].Value = table.Rows[i][15];
                            //rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[17].Value = table.Rows[i][15];
                            //index++;
                            Console.WriteLine(rasxod_dataGridView.Rows[rasxod_dataGridView.Rows.Count - 2].Cells[17].Value);
                        }
                        this.rasxod_dataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.rasxod_dataGridView_CellValueChanged);
                    }
                    this.rasxod_dataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.rasxod_dataGridView_CellValueChanged);
                }



                label_update_prixod();
            }
            catch (Exception ex)
            {
                MessageBox.Show("rasxod_dataGridView_CellDoubleClick " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        private void rasxod_dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            try
            {


                if (rasxod_dataGridView.SelectedCells.Count > 0)
                {
                    DataGridViewRow dgvRow = rasxod_dataGridView.CurrentRow;



                    if (e.ColumnIndex == 7)
                    {
                        // Console.WriteLine(dgvRow.Cells[7].Value);
                        dgvRow.Cells[9].Value = string.Format("{0:#0.00}", (dgvRow.Cells[7].Value != null ? (Convert.ToDouble(dgvRow.Cells[7].Value.ToString().Replace(".", ","))) : 0) *
                                                                           (dgvRow.Cells[8].Value != null ? (Convert.ToDouble(dgvRow.Cells[8].Value.ToString().Replace(".", ","))) : 0)
                                                                           );
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
                MessageBox.Show("rasxod_dataGridView_CellValueChanged " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }

        }

        private void rasxod_document_Load(object sender, EventArgs e)
        {
            this.rasxod_dataGridView.RowsDefaultCellStyle.BackColor = Color.White;
            this.rasxod_dataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(233, 233, 234);

            rasxod_dataGridView.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            rasxod_dataGridView.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            rasxod_dataGridView.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            rasxod_dataGridView.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            rasxod_dataGridView.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            rasxod_dataGridView.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            rasxod_dataGridView.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            rasxod_dataGridView.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            rasxod_dataGridView.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            rasxod_dataGridView.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            rasxod_dataGridView.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            rasxod_dataGridView.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            rasxod_dataGridView.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            rasxod_dataGridView.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            rasxod_dataGridView.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            rasxod_dataGridView.Columns[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        private void komu_ras_comboBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void label24_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void ot_kogo_ras_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ot_kogo_ras_comboBox2.Text = "";
                ot_kogo_ras_comboBox2.Items.Clear();
                var select = "SELECT * FROM podraz_jur7 where podraz_naim='" + ot_kogo_ras_ComboBox.Text + "'";
                sql3.myReader = sql3.return_MySqlCommand(select).ExecuteReader();
                while (sql3.myReader.Read())
                {
                    ot_kogo_ras_comboBox2.Items.Add(sql3.myReader["fio"] != DBNull.Value ? sql3.myReader.GetString("fio") : "");
                }
                sql3.myReader.Close();

            }
            catch (Exception ex)
            {

                MessageBox.Show("ot_kogo_ras_comboBox2_SelectedIndexChanged " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            naryad_num_prixod_int = scetcik_num_prixod_int;
            set_items_to_values_prixod();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            if (naryad_num_prixod_int < scetcik_num_prixod_int + 1)
            {
                naryad_num_prixod_int++;
                set_items_to_values_prixod();
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (naryad_num_prixod_int > 1)
            {
                naryad_num_prixod_int--;
                set_items_to_values_prixod();
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            naryad_num_prixod_int = 1;
            set_items_to_values_prixod();
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


        int exist = 0;
        int id_sklad_products;
        private void ras_save_btn_Click(object sender, EventArgs e)
        {
            try
            {
                if (jur_order_ras_textBox.Text != "" && num_ras_textBox.Text != "" && ot_kogo_ras_ComboBox.Text != "" && ot_kogo_ras_comboBox2.Text != "")
                {

                    sql.myReader = sql.return_MySqlCommand("select exists(SELECT * FROM products_rasxod_jur7 where vid_doc='2' and year='" + year_global + "' and month='" + month_global + "' and kod_doc='" + naryad_num_prixod_int + "' and user='" + string_for_otdels + "' ) as ex").ExecuteReader();
                    while (sql.myReader.Read())
                    {
                        exist = sql.myReader.GetInt32("ex");
                    }
                    sql.myReader.Close();

                    if (exist == 0)
                    {
                        //var insert = "insert into document (jur_order,num,data_reg,primech,doverennost,fio_ot,fio_ot_2,fio_komu,vid_doc,year,month,user,kod_doc) values(" +
                        //                           "'" + (jur_order_ras_textBox.Text) + "'," +
                        //                           "'" + (num_ras_textBox.Text) + "'," +
                        //                           (data_ras_DateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(data_ras_DateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                        //                          "'" + (primech_ras_textBox.Text) + "'," +
                        //                          "'" + (doveren_ras_textBox.Text) + "'," +
                        //                          "'" + (ot_kogo_ras_ComboBox.Text) + "'," +
                        //                          "'" + (ot_kogo_ras_comboBox2.Text) + "'," +
                        //                          "'" + (komu_ras_comboBox.Text) + "'," +
                        //                          "'" + ('2') + "', " +
                        //                           "'" + (year_global) + "', " +
                        //                            "'" + (month_global) + "', " +
                        //                             "'" + (string_for_otdels) + "', " +
                        //                            "'" + (kod_num_textBox.Text) + "' " +
                        //                           ");";
                        //sql.return_MySqlCommand(insert).ExecuteNonQuery();

                        for (int i = 0; i < rasxod_dataGridView.Rows.Count - 1; i++)
                        {


                            var insert_product = "insert into products_jur7 (user,year,month,vid_doc,kod_doc,product_id,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,ot_kogo_2,komu_1) values(" +
                                            "'" + string_for_otdels + "'," +
                                            "'" + (year_global) + "'," +
                                            "'" + (month_global) + "'," +
                                            "'" + ("2") + "'," +
                                            "'" + (kod_num_textBox.Text) + "'," +
                                             //"'" + (product_id) + "'," +
                                             "'" + (rasxod_dataGridView.Rows[i].Cells[1].Value != null ? rasxod_dataGridView.Rows[i].Cells[1].Value.ToString() : "") + "'," +
                                            "'" + (rasxod_dataGridView.Rows[i].Cells[2].Value != null ? rasxod_dataGridView.Rows[i].Cells[2].Value.ToString() : "") + "'," +
                                             "'" + (rasxod_dataGridView.Rows[i].Cells[3].Value != null ? rasxod_dataGridView.Rows[i].Cells[3].Value.ToString() : "") + "'," +
                                              "'" + (rasxod_dataGridView.Rows[i].Cells[4].Value != null ? rasxod_dataGridView.Rows[i].Cells[4].Value.ToString() : "") + "'," +
                                              "'" + (rasxod_dataGridView.Rows[i].Cells[5].Value != null ? rasxod_dataGridView.Rows[i].Cells[5].Value.ToString() : "") + "'," +
                                              "'" + (rasxod_dataGridView.Rows[i].Cells[6].Value != null ? rasxod_dataGridView.Rows[i].Cells[6].Value.ToString() : "") + "'," +
                                            "'" + refresh_strings_to_mysql(rasxod_dataGridView.Rows[i].Cells[7].Value != null ? rasxod_dataGridView.Rows[i].Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                            "'" + refresh_strings_to_mysql(rasxod_dataGridView.Rows[i].Cells[8].Value != null ? rasxod_dataGridView.Rows[i].Cells[8].Value.ToString().Replace(',', '.') : "0") + "'," +
                                            "'" + refresh_strings_to_mysql(rasxod_dataGridView.Rows[i].Cells[9].Value != null ? rasxod_dataGridView.Rows[i].Cells[9].Value.ToString().Replace(',', '.') : "0") + "'," +
                                            "'" + (rasxod_dataGridView.Rows[i].Cells[10].Value != null ? rasxod_dataGridView.Rows[i].Cells[10].Value.ToString() : "") + "'," +
                                            "'" + (rasxod_dataGridView.Rows[i].Cells[11].Value != null ? rasxod_dataGridView.Rows[i].Cells[11].Value.ToString() : "") + "'," +
                                             "'" + (rasxod_dataGridView.Rows[i].Cells[12].Value != null ? rasxod_dataGridView.Rows[i].Cells[12].Value.ToString() : "") + "'," +
                                              "'" + (rasxod_dataGridView.Rows[i].Cells[13].Value != null ? rasxod_dataGridView.Rows[i].Cells[13].Value.ToString() : "") + "'," +
                                              "'" + (rasxod_dataGridView.Rows[i].Cells[14].Value != null ? rasxod_dataGridView.Rows[i].Cells[14].Value.ToString() : "") + "'," +
                                              "'" + (rasxod_dataGridView.Rows[i].Cells[15].Value != null ? rasxod_dataGridView.Rows[i].Cells[15].Value.ToString() : "") + "'," +
                                            "'" + refresh_strings_to_mysql(rasxod_dataGridView.Rows[i].Cells[16].Value != null ? rasxod_dataGridView.Rows[i].Cells[16].Value.ToString().Replace(',', '.') : "0") + "'," +
                                           "" + (rasxod_dataGridView.Rows[i].Cells[17].Value != null ? "'" + DateTime.Parse(rasxod_dataGridView.Rows[i].Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                                            //"" + (rasxod_dataGridView.Rows[i].Cells[17].Value == null ? "'" + DateTime.Parse(rasxod_dataGridView.Rows[i].Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                                            "'" + (jur_order_ras_textBox.Text) + "', " +
                                            "'" + (num_ras_textBox.Text) + "', " +
                                            (data_ras_DateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(data_ras_DateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                                            "'" + (primech_ras_textBox.Text) + "', " +
                                            "'" + (doveren_ras_textBox.Text) + "', " +
                                            "'" + (ot_kogo_ras_ComboBox.Text) + "', " +
                                            "'" + (ot_kogo_ras_comboBox2.Text) + "', " +
                                            "'" + (komu_ras_comboBox.Text) + "' " +
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
                                             "'" + (year_global) + "'," +
                                             "'" + (month_global) + "'," +
                                             "'" + ("2") + "'," +
                                             "'" + (id_sklad_products) + "'," +
                                             "'" + (kod_num_textBox.Text) + "'," +
                                             //"'" + (naryad_num_prixod_int) + "'," +//kod_tov orniga
                                             "'" + (rasxod_dataGridView.Rows[i].Cells[1].Value != null ? rasxod_dataGridView.Rows[i].Cells[1].Value.ToString() : "") + "'," +
                                             "'" + (rasxod_dataGridView.Rows[i].Cells[2].Value != null ? rasxod_dataGridView.Rows[i].Cells[2].Value.ToString() : "") + "'," +
                                              "'" + (rasxod_dataGridView.Rows[i].Cells[3].Value != null ? rasxod_dataGridView.Rows[i].Cells[3].Value.ToString() : "") + "'," +
                                               "'" + (rasxod_dataGridView.Rows[i].Cells[4].Value != null ? rasxod_dataGridView.Rows[i].Cells[4].Value.ToString() : "") + "'," +
                                               "'" + (rasxod_dataGridView.Rows[i].Cells[5].Value != null ? rasxod_dataGridView.Rows[i].Cells[5].Value.ToString() : "") + "'," +
                                               "'" + (rasxod_dataGridView.Rows[i].Cells[6].Value != null ? rasxod_dataGridView.Rows[i].Cells[6].Value.ToString() : "") + "'," +
                                             "'" + refresh_strings_to_mysql(rasxod_dataGridView.Rows[i].Cells[7].Value != null ? rasxod_dataGridView.Rows[i].Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                             "'" + refresh_strings_to_mysql(rasxod_dataGridView.Rows[i].Cells[8].Value != null ? rasxod_dataGridView.Rows[i].Cells[8].Value.ToString().Replace(',', '.') : "0") + "'," +
                                             "'" + refresh_strings_to_mysql(rasxod_dataGridView.Rows[i].Cells[9].Value != null ? rasxod_dataGridView.Rows[i].Cells[9].Value.ToString().Replace(',', '.') : "0") + "'," +
                                             "'" + (rasxod_dataGridView.Rows[i].Cells[10].Value != null ? rasxod_dataGridView.Rows[i].Cells[10].Value.ToString() : "") + "'," +
                                             "'" + (rasxod_dataGridView.Rows[i].Cells[11].Value != null ? rasxod_dataGridView.Rows[i].Cells[11].Value.ToString() : "") + "'," +
                                              "'" + (rasxod_dataGridView.Rows[i].Cells[12].Value != null ? rasxod_dataGridView.Rows[i].Cells[12].Value.ToString() : "") + "'," +
                                               "'" + (rasxod_dataGridView.Rows[i].Cells[13].Value != null ? rasxod_dataGridView.Rows[i].Cells[13].Value.ToString() : "") + "'," +
                                               "'" + (rasxod_dataGridView.Rows[i].Cells[14].Value != null ? rasxod_dataGridView.Rows[i].Cells[14].Value.ToString() : "") + "'," +
                                               "'" + (rasxod_dataGridView.Rows[i].Cells[15].Value != null ? rasxod_dataGridView.Rows[i].Cells[15].Value.ToString() : "") + "'," +
                                             "'" + refresh_strings_to_mysql(rasxod_dataGridView.Rows[i].Cells[16].Value != null ? rasxod_dataGridView.Rows[i].Cells[16].Value.ToString().Replace(',', '.') : "0") + "'," +
                                              "" + (rasxod_dataGridView.Rows[i].Cells[17].Value != null ? "'" + DateTime.Parse(rasxod_dataGridView.Rows[i].Cells[17].Value.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL") + ", " +
                                             //"'" + (rasxod_dataGridView.Rows[i].Cells[18].Value != null ? rasxod_dataGridView.Rows[i].Cells[18].Value.ToString() : "") + "'," +
                                             "'" + (jur_order_ras_textBox.Text) + "', " +
                                             "'" + (num_ras_textBox.Text) + "', " +
                                             (data_ras_DateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(data_ras_DateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                                             "'" + (primech_ras_textBox.Text) + "', " +
                                             "'" + (doveren_ras_textBox.Text) + "', " +
                                             "'" + (ot_kogo_ras_ComboBox.Text) + "', " +
                                             "'" + (ot_kogo_ras_comboBox2.Text) + "', " +
                                             "'" + (komu_ras_comboBox.Text) + "' " +
                                             ")";
                            sql.return_MySqlCommand(insert_product_pri).ExecuteNonQuery();

                        }
                        scetcik_num_prixod_int++;
                        label_update_prixod();
                        //sql.return_MySqlCommand("insert into prixod_rasxod (name, prixod,rasxod) values ('" + ot_kogoComboBox.Text + "', '" + (nakladnaydataGridView.Rows[i].Cells[7].Value != null ? nakladnaydataGridView.Rows[i].Cells[7].Value.ToString().Replace(',', '.') : "0") + "','0');").ExecuteNonQuery();
                        MessageBox.Show("Добавлено ", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        set_items_to_values_prixod();
                    }
                    else
                    {
                        DialogResult dialogResult = MessageBox.Show("Обновить данные?", "Обновление", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                        {
                            //var update_document = "update document set " +

                            //    "jur_order = '" + jur_order_ras_textBox.Text + "'," +
                            //    "data_reg = '" + (data_ras_DateTimePicker.Text != DBNull.Value.ToString() ? DateTime.Parse(data_ras_DateTimePicker.Text).ToString("yyyy-MM-dd") : "NULL") + "'," +
                            //    "num = '" + num_ras_textBox.Text + "'," +
                            //    "primech = '" + primech_ras_textBox.Text + "'," +
                            //    "doverennost = '" + doveren_ras_textBox.Text + "'," +
                            //    "fio_ot = '" + ot_kogo_ras_ComboBox.Text + "'," +
                            //    "fio_ot_2 = '" + ot_kogo_ras_comboBox2.Text + "'," +
                            //    "fio_komu = '" + komu_ras_comboBox.Text + "'," +
                            //    "year = '" + year_global + "'," +
                            //    "month = '" + month_global + "'," +
                            //      "user = '" + string_for_otdels + "'," +
                            //    "kod_doc = '" + kod_num_textBox.Text + "'" +

                            //   " where kod_doc = " + kod_num_textBox.Text + "";
                            //sql.return_MySqlCommand(update_document).ExecuteNonQuery();


                            foreach (DataGridViewRow row in rasxod_dataGridView.Rows)
                            {

                                if (row.Cells[1].Value != null && row.Cells[0].Value != null && row.Cells[18].Value != null)
                                {

                                    //kod_tov,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,
                                    //deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,id_products
                                    //jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,komu_1,komu_2

                                    var update = "update products_jur7 set " +

                                        "user = '" + string_for_otdels + "'," +
                                        "year = '" + year_global + "'," +
                                        "month = '" + month_global + "'," +
                                        "vid_doc ='" + "2" + "'," +
                                        "kod_doc ='" + naryad_num_prixod_int + "'," +
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
                                        "jur_order = '" + jur_order_ras_textBox.Text + "'," +
                                        "num_doc = '" + num_ras_textBox.Text + "'," +
                                        "date_doc = '" + (data_ras_DateTimePicker.Text != DBNull.Value.ToString() ? DateTime.Parse(data_ras_DateTimePicker.Text).ToString("yyyy-MM-dd") : "NULL") + "'," +
                                        "primech = '" + primech_ras_textBox.Text + "'," +
                                        "doverennost = '" + doveren_ras_textBox.Text + "'," +
                                        "ot_kogo = '" + ot_kogo_ras_ComboBox.Text + "'," +
                                        "ot_kogo_2 = '" + ot_kogo_ras_comboBox2.Text + "'," +
                                        "komu_1 = '" + komu_ras_comboBox.Text + "'" +
                                        " where id = " + row.Cells[18].Value + "";
                                    sql.return_MySqlCommand(update).ExecuteNonQuery();

                                    var update2 = "update products_rasxod_jur7 set " +

                                       "user = '" + string_for_otdels + "'," +
                                       "year = '" + year_global + "'," +
                                       "month = '" + month_global + "'," +
                                       "vid_doc ='" + "2" + "'," +
                                       "kod_doc ='" + naryad_num_prixod_int + "'," +
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
                                       "jur_order = '" + jur_order_ras_textBox.Text + "'," +
                                       "num_doc = '" + num_ras_textBox.Text + "'," +
                                       "date_doc = '" + (data_ras_DateTimePicker.Text != DBNull.Value.ToString() ? DateTime.Parse(data_ras_DateTimePicker.Text).ToString("yyyy-MM-dd") : "NULL") + "'," +
                                       "primech = '" + primech_ras_textBox.Text + "'," +
                                       "doverennost = '" + doveren_ras_textBox.Text + "'," +
                                       "ot_kogo = '" + ot_kogo_ras_ComboBox.Text + "'," +
                                       "ot_kogo_2 = '" + ot_kogo_ras_comboBox2.Text + "'," +
                                       "komu_1 = '" + komu_ras_comboBox.Text + "'" +
                                       " where id = " + row.Cells[0].Value + "";
                                    sql.return_MySqlCommand(update2).ExecuteNonQuery();


                                }
                                if (row.Cells[0].Value == null && row.Cells[7].Value != null)
                                {
                                    //var naim_tov = "insert into naim_tov (naim,kod_gruppa) values(" +
                                    //       "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                    //        "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'" +
                                    //       ")";
                                    //sql.return_MySqlCommand(naim_tov).ExecuteNonQuery();


                                    //var query3 = "SELECT max(id) as product_id FROM naim_tov";
                                    //sql.myReader = sql.return_MySqlCommand(query3).ExecuteReader();
                                    //while (sql.myReader.Read())
                                    //{
                                    //    product_id = sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetInt32("product_id") : 1;
                                    //}
                                    //sql.myReader.Close();

                                    var insert_product = "insert into products_jur7 (user,year,month,vid_doc,kod_doc,product_id,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,jur_order,num_doc,date_doc,primech,doverennost,ot_kogo,ot_kogo_2,komu_1) values(" +
                                           "'" + string_for_otdels + "'," +
                                           "'" + (year_global) + "'," +
                                           "'" + (month_global) + "'," +
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
                                           "'" + (jur_order_ras_textBox.Text) + "', " +
                                           "'" + (num_ras_textBox.Text) + "', " +
                                           (data_ras_DateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(data_ras_DateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                                           "'" + (primech_ras_textBox.Text) + "', " +
                                           "'" + (doveren_ras_textBox.Text) + "', " +
                                           "'" + (ot_kogo_ras_ComboBox.Text) + "', " +
                                           "'" + (ot_kogo_ras_comboBox2.Text) + "', " +
                                           "'" + (komu_ras_comboBox.Text) + "' " +
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
                                                     "'" + (year_global) + "'," +
                                                     "'" + (month_global) + "'," +
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
                                                     "'" + (jur_order_ras_textBox.Text) + "', " +
                                                     "'" + (num_ras_textBox.Text) + "', " +
                                                     (data_ras_DateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(data_ras_DateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                                                     "'" + (primech_ras_textBox.Text) + "', " +
                                                     "'" + (doveren_ras_textBox.Text) + "', " +
                                                     "'" + (ot_kogo_ras_ComboBox.Text) + "', " +
                                                     "'" + (ot_kogo_ras_comboBox2.Text) + "', " +
                                                     "'" + (komu_ras_comboBox.Text) + "' " +
                                                     ")";
                                    sql.return_MySqlCommand(insert_product_pri).ExecuteNonQuery();

                                }


                            }


                            MessageBox.Show("Обновлено");
                            set_items_to_values_prixod();
                        }
                        else if (dialogResult == DialogResult.No)
                        {
                            //MessageBox.Show("ma'lumot to'liq emas");
                        }
                        run_main();
                    }
                }
                else
                {
                    MessageBox.Show("Не хватает значений для сохранения , mumkin emas!!!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("save_button_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void rasxod_dataGridView_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {

        }

        string fio_gl_bugalter = "";
        string fio_bugalter = "";
        string inspektor = "";

        int[] b = new int[2];
        private void ras_document_btn_Click(object sender, EventArgs e)
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
            sheet.Range["b1:b1"].ColumnWidth = 23.14;
            sheet.Range["c1:c1"].ColumnWidth = 5.71;
            sheet.Range["d1:d1"].ColumnWidth = 4;
            sheet.Range["e1:e1"].ColumnWidth = 6.57;
            sheet.Range["f1:f1"].ColumnWidth = 6;
            sheet.Range["g1:g1"].ColumnWidth = 9;
            sheet.Range["h1:h1"].ColumnWidth = 12;
            sheet.Range["i1:i1"].ColumnWidth = 10;
            sheet.Range["j1:j1"].ColumnWidth = 4;
            sheet.Range["k1:k1"].ColumnWidth = 4;
            sheet.Range["l1:l1"].ColumnWidth = 4;
            sheet.Range["m1:m1"].ColumnWidth = 4;

            sql.myReader.Close();
            string name_organ = "";
            var name_org = "SELECT * FROM spravochnik_main_jur7 where user_jur7='" + string_for_otdels + "'";

            sql.myReader = sql.return_MySqlCommand(name_org).ExecuteReader();
            while (sql.myReader.Read())
            {
                name_organ = (sql.myReader["naim_org"] != DBNull.Value ? sql.myReader.GetString("naim_org") : "");
            }
            sql.myReader.Close();

            sheet.Range["a1:m1"].Style.Font.IsBold = true;
            sheet.Range["a1:m1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a1:m1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a1:m1"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a1:m1"].Style.Font.Size = 12;
            sheet.Range["a1:m1"].Merge(); // birlashtirish
            sheet.Range["a1:m1"].Text = name_organ;
            sheet.Range["a1:m1"].Style.Font.Color = Color.DarkBlue;
            sheet.SetRowHeight(1, 21);



            sheet.Range["a2:m2"].Style.Font.IsBold = true;
            //sheet.Range["a2:m2"].Style.Font.IsItalic = true;
            sheet.Range["a2:m2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a2:m2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a2:m2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a2:m2"].Style.Font.Size = 12;
            sheet.Range["a2:m2"].Merge(); // birlashtirish
            sheet.Range["a2:m2"].Text = "ТРЕБОВАНИЕ на отпуск материалов № " + num_ras_textBox.Text + " от " + data_ras_DateTimePicker.Value.ToString("dd.MM.yyyy");
            sheet.Range["a2:m2"].Style.WrapText = true;
            sheet.SetRowHeight(2, 21);

            sheet.Range["b3:b3"].Style.Font.IsBold = true;
            //sheet.Range["b3:j3"].Style.Font.IsItalic = true;
            sheet.Range["b3:b3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b3:b3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b3:b3"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b3:b3"].Style.Font.Size = 11;
            sheet.Range["b3:b3"].Merge(); // birlashtirish
            sheet.Range["b3:b3"].Text = "Выдатъ(откуда): ";

            sheet.Range["c3:l3"].Style.Font.IsBold = true;
            //sheet.Range["b3:j3"].Style.Font.IsItalic = true;
            sheet.Range["c3:l3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c3:l3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c3:l3"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["c3:l3"].Style.Font.Size = 11;
            sheet.Range["c3:l3"].Merge(); // birlashtirish
            sheet.Range["c3:l3"].Text = ot_kogo_ras_comboBox2.Text + " " + ot_kogo_ras_ComboBox.Text;
            sheet.Range["c3:l3"].Style.WrapText = true;
            sheet.Range["c3:l3"].Style.WrapText = true;
            sheet.Range["c3:l3"].BorderAround(LineStyleType.Thin);
            sheet.SetRowHeight(3, 18);

            sheet.Range["b4:b4"].Style.Font.IsBold = true;
            //sheet.Range["bb:j4"].Style.Font.IsItalic = true;
            sheet.Range["b4:b4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b4:b4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b4:b4"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b4:b4"].Style.Font.Size = 11;
            sheet.Range["b4:b4"].Merge(); // birlashtirish
            sheet.Range["b4:b4"].Text = "Основание: ";
            sheet.Range["b4:b4"].Style.WrapText = true;

            sheet.Range["c4:l4"].Style.Font.IsBold = true;
            //sheet.Range["bb:j4"].Style.Font.IsItalic = true;
            sheet.Range["c4:l4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c4:l4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c4:l4"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["c4:l4"].Style.Font.Size = 11;
            sheet.Range["c4:l4"].Merge(); // birlashtirish
            sheet.Range["c4:l4"].Text = primech_ras_textBox.Text;
            sheet.Range["c4:l4"].Style.WrapText = true;
            sheet.Range["c4:l4"].BorderAround(LineStyleType.Thin);

            sheet.SetRowHeight(4, 18);

            sheet.Range["b5:b5"].Style.Font.IsBold = true;
            //sheet.Range["b5:j5"].Style.Font.IsItalic = true;
            sheet.Range["b5:b5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b5:b5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b5:b5"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b5:b5"].Style.Font.Size = 11;
            sheet.Range["b5:b5"].Merge(); // birlashtirish
            sheet.Range["b5:b5"].Text = "Кому: ";
            sheet.Range["b5:b5"].Style.WrapText = true;


            sheet.Range["c5:l5"].Style.Font.IsBold = true;
            //sheet.Range["b5:j5"].Style.Font.IsItalic = true;
            sheet.Range["c5:l5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c5:l5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c5:l5"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["c5:l5"].Style.Font.Size = 11;
            sheet.Range["c5:l5"].Merge(); // birlashtirish
            sheet.Range["c5:l5"].Text = komu_ras_comboBox.Text;
            sheet.Range["c5:l5"].Style.WrapText = true;
            sheet.Range["c5:l5"].BorderAround(LineStyleType.Thin);

            sheet.SetRowHeight(5, 20);

            sheet.Range["a6:a6"].Style.Font.IsBold = true;
            sheet.Range["a6:a6"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a6:a6"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a6:a6"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a6:a6"].Style.Font.Size = 11;
            sheet.Range["a6:a6"].Merge(); // birlashtirish
            sheet.Range["a6:a6"].Text = "№";
            sheet.Range["a6:a6"].Style.WrapText = true;
            sheet.Range["a6:a6"].BorderAround(LineStyleType.Thin);

            sheet.Range["b6:b6"].Style.Font.IsBold = true;
            sheet.Range["b6:b6"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b6:b6"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b6:b6"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b6:b6"].Style.Font.Size = 11;
            sheet.Range["b6:b6"].Merge(); // birlashtirish
            sheet.Range["b6:b6"].Text = "Наименование";
            sheet.Range["b6:b6"].Style.WrapText = true;
            sheet.Range["b6:b6"].BorderAround(LineStyleType.Thin);

            sheet.Range["c6:c6"].Style.Font.IsBold = true;
            sheet.Range["c6:c6"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c6:c6"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c6:c6"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c6:c6"].Style.Font.Size = 9;
            sheet.Range["c6:c6"].Merge(); // birlashtirish
            sheet.Range["c6:c6"].Text = "Инв. №";
            sheet.Range["c6:c6"].BorderAround(LineStyleType.Thin);
            sheet.Range["c6:c6"].Style.WrapText = true;

            sheet.Range["d6:d6"].Style.Font.IsBold = true;
            sheet.Range["d6:d6"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d6:d6"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d6:d6"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d6:d6"].Style.Font.Size = 11;
            sheet.Range["d6:d6"].Merge(); // birlashtirish
            sheet.Range["d6:d6"].Text = "Ед.";
            sheet.Range["d6:d6"].Style.WrapText = true;
            sheet.Range["d6:d6"].BorderAround(LineStyleType.Thin);

            sheet.Range["e6:e6"].Style.Font.IsBold = true;
            sheet.Range["e6:e6"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e6:e6"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e6:e6"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e6:e6"].Style.Font.Size = 11;
            sheet.Range["e6:e6"].Merge(); // birlashtirish
            sheet.Range["e6:e6"].Text = "треб.";
            sheet.Range["e6:e6"].Style.WrapText = true;
            sheet.Range["e6:e6"].BorderAround(LineStyleType.Thin);

            sheet.Range["f6:f6"].Style.Font.IsBold = true;
            sheet.Range["f6:f6"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f6:f6"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f6:f6"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f6:f6"].Style.Font.Size = 10;
            sheet.Range["f6:f6"].Merge(); // birlashtirish
            sheet.Range["f6:f6"].Text = "отпущ.";
            sheet.Range["f6:f6"].Style.WrapText = true;
            sheet.Range["f6:f6"].BorderAround(LineStyleType.Thin);

            sheet.Range["g6:g6"].Style.Font.IsBold = true;
            sheet.Range["g6:g6"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g6:g6"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g6:g6"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g6:g6"].Style.Font.Size = 11;
            sheet.Range["g6:g6"].Merge(); // birlashtirish
            sheet.Range["g6:g6"].Text = "Цена";
            sheet.Range["g6:g6"].Style.WrapText = true;
            sheet.Range["g6:g6"].BorderAround(LineStyleType.Thin);

            sheet.Range["h6:h6"].Style.Font.IsBold = true;
            sheet.Range["h6:h6"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h6:h6"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h6:h6"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["h6:h6"].Style.Font.Size = 11;
            sheet.Range["h6:h6"].Merge(); // birlashtirish
            sheet.Range["h6:h6"].Text = "Сумма";
            sheet.Range["h6:h6"].Style.WrapText = true;
            sheet.Range["h6:h6"].BorderAround(LineStyleType.Thin);

            sheet.Range["i6:i6"].Style.Font.IsBold = true;
            sheet.Range["i6:i6"].Style.Font.FontName = "Times New Roman";
            sheet.Range["i6:i6"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["i6:i6"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["i6:i6"].Style.Font.Size = 11;
            sheet.Range["i6:i6"].Merge(); // birlashtirish
            sheet.Range["i6:i6"].Text = "Сум.изн";
            sheet.Range["i6:i6"].Style.WrapText = true;
            sheet.Range["i6:i6"].BorderAround(LineStyleType.Thin);


            sheet.Range["j6:k6"].Style.Font.IsBold = true;
            sheet.Range["j6:k6"].Style.Font.FontName = "Times New Roman";
            sheet.Range["j6:k6"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["j6:k6"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["j6:k6"].Style.Font.Size = 11;
            sheet.Range["j6:k6"].Merge(); // birlashtirish
            sheet.Range["j6:k6"].Text = "Дебет";
            sheet.Range["j6:k6"].Style.WrapText = true;
            sheet.Range["j6:k6"].BorderAround(LineStyleType.Thin);


            sheet.Range["l6:m6"].Style.Font.IsBold = true;
            sheet.Range["l6:m6"].Style.Font.FontName = "Times New Roman";
            sheet.Range["l6:m6"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["l6:m6"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["l6:m6"].Style.Font.Size = 11;
            sheet.Range["l6:m6"].Merge(); // birlashtirish
            sheet.Range["l6:m6"].Text = "Кредит";
            sheet.Range["l6:m6"].Style.WrapText = true;
            sheet.Range["l6:m6"].BorderAround(LineStyleType.Thin);


            int i = 0;
            int myrow = 7;
            int j = 0;

            double all_kol_count = 0;

            var top = " SELECT * FROM products_rasxod_jur7 where kod_doc='" + kod_num_textBox.Text + "' and user='" + string_for_otdels + "' and month='" + month_global + "' and year='" + year_global + "'  ";
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
                sheet.Range["c" + myrow + ":c" + myrow].Text = sql.myReader["inventar_num"] != DBNull.Value ? sql.myReader.GetString("inventar_num").ToString() : "";
                sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;

                sheet.Range["d" + myrow + ":d" + myrow].Merge();
                sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
                sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["d" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 10;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["d" + myrow + ":d" + myrow].Text = sql.myReader["edin"] != DBNull.Value ? sql.myReader.GetString("edin").ToString() : "";

                all_kol_count = all_kol_count + (sql.myReader["kol"] != DBNull.Value ? sql.myReader.GetDouble("kol") : 0);

                sheet.Range["e" + myrow + ":e" + myrow].Merge();
                sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 10;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["e" + myrow + ":e" + myrow].Value = sql.myReader["kol"] != DBNull.Value ? sql.myReader.GetString("kol").ToString() : "";



                sheet.Range["f" + myrow + ":f" + myrow].Merge();
                sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":f" + myrow].Value = "";// sql.myReader["kol"] != DBNull.Value ? sql.myReader.GetString("kol").ToString() : "";

                sheet.Range["g" + myrow + ":g" + myrow].Merge();
                sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
                sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["g" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 10;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["g" + myrow + ":g" + myrow].Value = sql.myReader["sena"] != DBNull.Value ? sql.myReader.GetString("sena").ToString() : "";

                sheet.Range["h" + myrow + ":h" + myrow].Merge();
                sheet.Range["h" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["h" + myrow + ":h" + myrow].Style.WrapText = true;
                sheet.Range["h" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["h" + myrow + ":h" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.Size = 10;
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["h" + myrow + ":h" + myrow].Value = sql.myReader["summa"] != DBNull.Value ? sql.myReader.GetString("summa").ToString() : "";

                sheet.Range["i" + myrow + ":i" + myrow].Merge();
                sheet.Range["i" + myrow + ":i" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["i" + myrow + ":i" + myrow].Style.WrapText = true;
                sheet.Range["i" + myrow + ":i" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["i" + myrow + ":i" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["i" + myrow + ":i" + myrow].Style.Font.Size = 10;
                sheet.Range["i" + myrow + ":i" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["i" + myrow + ":i" + myrow].Value = sql.myReader["summa_iznos"] != DBNull.Value ? sql.myReader.GetString("summa_iznos").ToString() : "";

                sheet.Range["j" + myrow + ":j" + myrow].Merge();
                sheet.Range["j" + myrow + ":j" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["j" + myrow + ":j" + myrow].Style.WrapText = true;
                sheet.Range["j" + myrow + ":j" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["j" + myrow + ":j" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["j" + myrow + ":j" + myrow].Style.Font.Size = 10;
                sheet.Range["j" + myrow + ":j" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["j" + myrow + ":j" + myrow].Text = sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch").ToString() : "";

                sheet.Range["k" + myrow + ":k" + myrow].Merge();
                sheet.Range["k" + myrow + ":k" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["k" + myrow + ":k" + myrow].Style.WrapText = true;
                sheet.Range["k" + myrow + ":k" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["k" + myrow + ":k" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["k" + myrow + ":k" + myrow].Style.Font.Size = 10;
                sheet.Range["k" + myrow + ":k" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["k" + myrow + ":k" + myrow].Text = sql.myReader["deb_sch_2"] != DBNull.Value ? sql.myReader.GetString("deb_sch_2").ToString() : "";

                sheet.Range["l" + myrow + ":l" + myrow].Merge();
                sheet.Range["l" + myrow + ":l" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["l" + myrow + ":l" + myrow].Style.WrapText = true;
                sheet.Range["l" + myrow + ":l" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["l" + myrow + ":l" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["l" + myrow + ":l" + myrow].Style.Font.Size = 10;
                sheet.Range["l" + myrow + ":l" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["l" + myrow + ":l" + myrow].Text = sql.myReader["kre_sch"] != DBNull.Value ? sql.myReader.GetString("kre_sch").ToString() : "";

                sheet.Range["m" + myrow + ":m" + myrow].Merge();
                sheet.Range["m" + myrow + ":m" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["m" + myrow + ":m" + myrow].Style.WrapText = true;
                sheet.Range["m" + myrow + ":m" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["m" + myrow + ":m" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["m" + myrow + ":m" + myrow].Style.Font.Size = 10;
                sheet.Range["m" + myrow + ":m" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["m" + myrow + ":m" + myrow].Text = sql.myReader["kre_sch_2"] != DBNull.Value ? sql.myReader.GetString("kre_sch_2").ToString() : "";

                myrow = myrow + 1;
                i = i + 1;

            }
            sql.myReader.Close();

            String[] arr = refresh_strings_to_mysql(rasxod_obshiy_summa_label.Text).Split('.');
            b[0] = Convert.ToInt32(arr[0]);
            b[1] = Convert.ToInt32(arr[1]);

            sheet.Range["b" + myrow + ":g" + myrow].Merge();
            sheet.Range["b" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":g" + myrow].Style.WrapText = true;
            //sheet.Range["b" + myrow + ":g" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["b" + myrow + ":g" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":g" + myrow].Text = "Обший коп:" + i + " шт строка " + all_kol_count + " товар";
            //sheet.Range["b" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);

            sheet.Range["h" + myrow + ":h" + myrow].Merge();
            sheet.Range["h" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["h" + myrow + ":h" + myrow].Style.WrapText = true;
            sheet.Range["h" + myrow + ":h" + myrow].Style.Font.IsBold = true;
            sheet.Range["h" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h" + myrow + ":h" + myrow].Style.Font.Size = 11;
            sheet.Range["h" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["h" + myrow + ":h" + myrow].Text = "Всего:";
            sheet.Range["h" + myrow + ":h" + myrow].BorderAround(LineStyleType.Thin);

            sheet.Range["i" + myrow + ":j" + myrow].Merge();
            sheet.Range["i" + myrow + ":j" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["i" + myrow + ":j" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["i" + myrow + ":j" + myrow].Style.Font.Size = 11;
            sheet.Range["i" + myrow + ":j" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["i" + myrow + ":j" + myrow].Value = rasxod_obshiy_summa_label.Text;
            sheet.Range["i" + myrow + ":j" + myrow].Style.WrapText = true;
            sheet.Range["i" + myrow + ":j" + myrow].Style.Font.IsBold = true;
            sheet.Range["i" + myrow + ":j" + myrow].BorderAround(LineStyleType.Thin);

            sheet.Range["k" + myrow + ":m" + myrow].Merge();
            sheet.Range["k" + myrow + ":m" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["k" + myrow + ":m" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["k" + myrow + ":m" + myrow].Style.Font.Size = 11;
            sheet.Range["k" + myrow + ":m" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["k" + myrow + ":m" + myrow].Value = iznos_sum_lbl.Text;
            sheet.Range["k" + myrow + ":m" + myrow].Style.WrapText = true;
            sheet.Range["k" + myrow + ":m" + myrow].Style.Font.IsBold = true;
            sheet.Range["k" + myrow + ":m" + myrow].BorderAround(LineStyleType.Thin);
            sheet.SetRowHeight(myrow, 18);
            myrow++;

            sheet.Range["b" + myrow + ":m" + myrow].Merge();
            sheet.Range["b" + myrow + ":m" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":m" + myrow].Style.WrapText = true;
            //sheet.Range["b" + myrow + ":m" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":m" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":m" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":m" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":m" + myrow].Text = "Обшая сумма :" + number_russian.toWords(b[0]) + " сум " + b[1] + " тийин"; ;
            //sheet.Range["b" + myrow + ":m" + myrow].BorderAround(LineStyleType.Thin);
            sheet.SetRowHeight(myrow, 18);
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

            sheet.Range["g" + myrow + ":g" + myrow].Merge();
            sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
            sheet.Range["g" + myrow + ":g" + myrow].Style.Font.IsBold = true;
            sheet.Range["g" + myrow + ":g" + myrow].Style.Font.IsItalic = true;
            sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 10;
            sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["g" + myrow + ":g" + myrow].Text = "Деб.сч.";
            sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Color = Color.DarkBlue;

            sheet.Range["h" + myrow + ":h" + myrow].Merge();
            sheet.Range["h" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["h" + myrow + ":h" + myrow].Style.WrapText = true;
            sheet.Range["h" + myrow + ":h" + myrow].Style.Font.IsBold = true;
            sheet.Range["h" + myrow + ":h" + myrow].Style.Font.IsItalic = true;
            sheet.Range["h" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h" + myrow + ":h" + myrow].Style.Font.Size = 10;
            sheet.Range["h" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["h" + myrow + ":h" + myrow].Text = "Кре.сч.";
            sheet.Range["h" + myrow + ":h" + myrow].Style.Font.Color = Color.DarkBlue;

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

            var schet = "SELECT deb_sch,deb_sch_2,kre_sch,sum(summa) as summa FROM products_prixod_jur7 where kod_doc='" + kod_num_textBox.Text + "' and user='" + string_for_otdels + "' and month='" + month_global + "' and year='" + year_global + "' group by deb_sch";

            sql.myReader = sql.return_MySqlCommand(schet).ExecuteReader();
            while (sql.myReader.Read())
            {
                sheet.Range["f" + myrow + ":f" + myrow].Merge();
                sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.IsBold = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.IsItalic = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":f" + myrow].Text = sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch").ToString() : "";

                sheet.Range["g" + myrow + ":g" + myrow].Merge();
                sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.IsBold = true;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.IsItalic = true;
                sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 10;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["g" + myrow + ":g" + myrow].Text = sql.myReader["deb_sch_2"] != DBNull.Value ? sql.myReader.GetString("deb_sch_2").ToString() : "";

                sheet.Range["h" + myrow + ":h" + myrow].Merge();
                sheet.Range["h" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["h" + myrow + ":h" + myrow].Style.WrapText = true;
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.IsBold = true;
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.IsItalic = true;
                sheet.Range["h" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.Size = 10;
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["h" + myrow + ":h" + myrow].Text = sql.myReader["kre_sch"] != DBNull.Value ? sql.myReader.GetString("kre_sch").ToString() : "";

                sheet.Range["j" + myrow + ":m" + myrow].Merge();
                sheet.Range["j" + myrow + ":m" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["j" + myrow + ":m" + myrow].Style.WrapText = true;
                sheet.Range["j" + myrow + ":m" + myrow].Style.Font.IsBold = true;
                sheet.Range["j" + myrow + ":m" + myrow].Style.Font.IsItalic = true;
                sheet.Range["j" + myrow + ":m" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["j" + myrow + ":m" + myrow].Style.Font.Size = 10;
                sheet.Range["j" + myrow + ":m" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["j" + myrow + ":m" + myrow].Value = sql.myReader["summa"] != DBNull.Value ? sql.myReader.GetString("summa").ToString() : "";

                sheet.SetRowHeight(myrow, 18);

                myrow++;
            }
            sql.myReader.Close();

            myrow++;

            sheet.Range["b" + myrow + ":d" + myrow].Merge();
            sheet.Range["b" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":d" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Top;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":d" + myrow].Text = "Отпустил:______________________ ";

            sheet.Range["h" + myrow + ":l" + myrow].Merge();
            sheet.Range["h" + myrow + ":l" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["h" + myrow + ":l" + myrow].Style.WrapText = true;
            sheet.Range["h" + myrow + ":l" + myrow].Style.Font.IsBold = true;
            sheet.Range["h" + myrow + ":l" + myrow].Style.VerticalAlignment = VerticalAlignType.Top;
            sheet.Range["h" + myrow + ":l" + myrow].Style.Font.Size = 11;
            sheet.Range["h" + myrow + ":l" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["h" + myrow + ":l" + myrow].Text = "Получил:_______________________ ";
            sheet.SetRowHeight(myrow, 18);
            sheet.Range["d5:" + myrow + "m"].NumberFormat = "#,##0.00";


            workbook.SaveToFile(Environment.CurrentDirectory + "\\docs\\Документ.xlsx", Spire.Xls.FileFormat.Version2007);
            System.Diagnostics.Process.Start(workbook.FileName);



            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Документ_excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void ras_izv_btn_Click(object sender, EventArgs e)
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

            sql.myReader.Close();
            string name_organ = "";
            var name_org = "SELECT * FROM spravochnik_main_jur7 where user_jur7='" + string_for_otdels + "'";

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
            sheet.Range["b3:k3"].Text = "о безвозмездной передаче основныих средств № " + num_ras_textBox.Text + " от " + data_ras_DateTimePicker.Value.ToString("dd.MM.yyyy");
            sheet.Range["b3:k3"].Style.WrapText = true;
            sheet.SetRowHeight(3, 18);

            //sheet.Range["a4:k4"].Style.Font.IsBold = true;
            //sheet.Range["a4:k4"].Style.Font.IsItalic = true;
            sheet.Range["a4:k4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a4:k4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a4:k4"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a4:k4"].Style.Font.Size = 11;
            sheet.Range["a4:k4"].Merge(); // birlashtirish
            sheet.Range["a4:k4"].Text = "Кому: " + komu_ras_comboBox.Text;
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
            sheet.Range["e5:k5"].Text = "Получателъ: " + komu_ras_comboBox.Text;
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
            sheet.Range["a6:k6"].Text = "Основание на передачу(распоряжение № и дата): " + primech_ras_textBox.Text;
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
            sheet.Range["a7:c7"].Text = "№ Доверенност: " + doveren_ras_textBox.Text;
            sheet.Range["a7:c7"].Style.WrapText = true;
            sheet.Range["a7:c7"].Style.Font.Underline = FontUnderlineType.SingleAccounting;

            //sheet.Range["d7:f7"].Style.Font.IsBold = true;
            //sheet.Range["d7:f7"].Style.Font.IsItalic = true;
            sheet.Range["d7:g7"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d7:g7"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d7:g7"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["d7:g7"].Style.Font.Size = 11;
            sheet.Range["d7:g7"].Merge(); // birlashtirish
            sheet.Range["d7:g7"].Text = "№ Требование: " + num_ras_textBox.Text;
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

            var top = " SELECT * FROM products_rasxod_jur7 where kod_doc='" + kod_num_textBox.Text + "' and user='" + string_for_otdels + "' and month='" + month_global + "' and year='" + year_global + "'  ";
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
            sheet.Range["f" + myrow + ":g" + myrow].Value = rasxod_obshiy_summa_label.Text;
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
            sheet.Range["a" + myrow + ":h" + myrow].Text = komu_ras_comboBox.Text;

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

        private void ras_spisanie_btn_Click(object sender, EventArgs e)
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
            sheet.Range["b1:b1"].ColumnWidth = 28.86;
            sheet.Range["c1:c1"].ColumnWidth = 4.86;
            sheet.Range["d1:d1"].ColumnWidth = 7;
            sheet.Range["e1:e1"].ColumnWidth = 10;
            sheet.Range["f1:f1"].ColumnWidth = 12;
            sheet.Range["g1:g1"].ColumnWidth = 9;
            sheet.Range["h1:h1"].ColumnWidth = 5;
            sheet.Range["i1:i1"].ColumnWidth = 6;
            sheet.Range["j1:j1"].ColumnWidth = 5;
            sheet.Range["k1:k1"].ColumnWidth = 6;


            string name_organ = "";
            var name_org = "SELECT * FROM spravochnik_main_jur7 where user_jur7='" + string_for_otdels + "'";

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

            sheet.Range["a5:k5"].Style.Font.IsBold = true;
            sheet.Range["a5:k5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a5:k5"].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["a5:k5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a5:k5"].Style.Font.Size = 12;
            sheet.Range["a5:k5"].Merge(); // birlashtirish
            sheet.Range["a5:k5"].Text = "Акт списания № " + num_ras_textBox.Text + " от " + data_ras_DateTimePicker.Value.ToString("dd.MM.yyyy");
            sheet.Range["a5:k5"].Style.WrapText = true;
            //sheet.Range["a5:j5"].BorderAround(LineStyleType.Thin);

            sheet.SetRowHeight(5, 24);

            sheet.Range["b6:k6"].Style.Font.IsBold = true;
            sheet.Range["b6:k6"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b6:k6"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b6:k6"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b6:k6"].Style.Font.Size = 11;
            sheet.Range["b6:k6"].Merge(); // birlashtirish
            sheet.Range["b6:k6"].Text = "От кого :" + ot_kogo_ras_comboBox2.Text + " " + ot_kogo_ras_comboBox2.Text + " " + primech_ras_textBox.Text;
            sheet.Range["b6:k6"].Style.WrapText = true;
            //sheet.Range["a6:j6"].BorderAround(LineStyleType.Thin);

            sheet.SetRowHeight(6, 18);

            sheet.SetRowHeight(7, 5);
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

            sheet.Range["g8:g8"].Style.Font.IsBold = true;
            sheet.Range["g8:g8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g8:g8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g8:g8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g8:g8"].Style.Font.Size = 11;
            sheet.Range["g8:g8"].Merge(); // birlashtirish
            sheet.Range["g8:g8"].Text = "Сум.изн.";
            sheet.Range["g8:g8"].Style.WrapText = true;
            sheet.Range["g8:g8"].BorderAround(LineStyleType.Thin);

            sheet.Range["h8:i8"].Style.Font.IsBold = true;
            sheet.Range["h8:i8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h8:i8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h8:i8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["h8:i8"].Style.Font.Size = 11;
            sheet.Range["h8:i8"].Merge(); // birlashtirish
            sheet.Range["h8:i8"].Text = "Дебет";
            sheet.Range["h8:i8"].Style.WrapText = true;
            sheet.Range["h8:i8"].BorderAround(LineStyleType.Thin);

            sheet.Range["j8:k8"].Style.Font.IsBold = true;
            sheet.Range["j8:k8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["j8:k8"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["j8:k8"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["j8:k8"].Style.Font.Size = 11;
            sheet.Range["j8:k8"].Merge(); // birlashtirish
            sheet.Range["j8:k8"].Text = "Кредит";
            sheet.Range["j8:k8"].Style.WrapText = true;
            sheet.Range["j8:k8"].BorderAround(LineStyleType.Thin);



            int i = 0;
            int myrow = 9;
            int j = 0;

            var top = " SELECT * FROM products_rasxod_jur7 where kod_doc='" + kod_num_textBox.Text + "' and user='" + string_for_otdels + "' and month='" + month_global + "' and year='" + year_global + "'  ";
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
                sheet.Range["g" + myrow + ":g" + myrow].Value = sql.myReader["summa_iznos"] != DBNull.Value ? sql.myReader.GetString("summa_iznos").ToString() : "";

                sheet.Range["h" + myrow + ":h" + myrow].Merge();
                sheet.Range["h" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["h" + myrow + ":h" + myrow].Style.WrapText = true;
                sheet.Range["h" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["h" + myrow + ":h" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.Size = 10;
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["h" + myrow + ":h" + myrow].Text = sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch").ToString() : "";

                sheet.Range["i" + myrow + ":i" + myrow].Merge();
                sheet.Range["i" + myrow + ":i" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["i" + myrow + ":i" + myrow].Style.WrapText = true;
                sheet.Range["i" + myrow + ":i" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["i" + myrow + ":i" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["i" + myrow + ":i" + myrow].Style.Font.Size = 10;
                sheet.Range["i" + myrow + ":i" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["i" + myrow + ":i" + myrow].Text = sql.myReader["deb_sch_2"] != DBNull.Value ? sql.myReader.GetString("deb_sch_2").ToString() : "";

                sheet.Range["j" + myrow + ":j" + myrow].Merge();
                sheet.Range["j" + myrow + ":j" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["j" + myrow + ":j" + myrow].Style.WrapText = true;
                sheet.Range["j" + myrow + ":j" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["j" + myrow + ":j" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["j" + myrow + ":j" + myrow].Style.Font.Size = 10;
                sheet.Range["j" + myrow + ":j" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["j" + myrow + ":j" + myrow].Text = sql.myReader["kre_sch"] != DBNull.Value ? sql.myReader.GetString("kre_sch").ToString() : "";

                sheet.Range["k" + myrow + ":k" + myrow].Merge();
                sheet.Range["k" + myrow + ":k" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["k" + myrow + ":k" + myrow].Style.WrapText = true;
                sheet.Range["k" + myrow + ":k" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["k" + myrow + ":k" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["k" + myrow + ":k" + myrow].Style.Font.Size = 10;
                sheet.Range["k" + myrow + ":k" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["k" + myrow + ":k" + myrow].Text = sql.myReader["kre_sch_2"] != DBNull.Value ? sql.myReader.GetString("kre_sch_2").ToString() : "";


                myrow = myrow + 1;
                i = i + 1;

            }
            sql.myReader.Close();



            sheet.Range["c" + myrow + ":d" + myrow].Merge();
            sheet.Range["c" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["c" + myrow + ":d" + myrow].Style.WrapText = true;
            sheet.Range["c" + myrow + ":d" + myrow].Style.Font.IsBold = true;
            sheet.Range["c" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c" + myrow + ":d" + myrow].Style.Font.Size = 11;
            sheet.Range["c" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["c" + myrow + ":d" + myrow].Text = "Всего:";
            sheet.Range["c" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);

            sheet.Range["e" + myrow + ":f" + myrow].Merge();
            sheet.Range["e" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["e" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Font.Size = 10;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["e" + myrow + ":f" + myrow].Value = rasxod_obshiy_summa_label.Text;
            sheet.Range["e" + myrow + ":f" + myrow].Style.WrapText = true;
            sheet.Range["e" + myrow + ":f" + myrow].Style.Font.IsBold = true;
            sheet.Range["e" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);

            sheet.Range["g" + myrow + ":h" + myrow].Merge();
            sheet.Range["g" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["g" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g" + myrow + ":h" + myrow].Style.Font.Size = 10;
            sheet.Range["g" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["g" + myrow + ":h" + myrow].Value = iznos_sum_lbl.Text;
            sheet.Range["g" + myrow + ":h" + myrow].Style.WrapText = true;
            sheet.Range["g" + myrow + ":h" + myrow].Style.Font.IsBold = true;
            sheet.Range["g" + myrow + ":h" + myrow].BorderAround(LineStyleType.Thin);
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

            sheet.Range["d" + myrow + ":d" + myrow].Merge();
            sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
            sheet.Range["d" + myrow + ":d" + myrow].Style.Font.IsBold = true;
            sheet.Range["d" + myrow + ":d" + myrow].Style.Font.IsItalic = true;
            sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 10;
            sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["d" + myrow + ":d" + myrow].Text = "Деб.сч.";

            sheet.Range["e" + myrow + ":e" + myrow].Merge();
            sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.IsBold = true;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.IsItalic = true;
            sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 10;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["e" + myrow + ":e" + myrow].Text = "Деб.сч.";

            sheet.Range["f" + myrow + ":f" + myrow].Merge();
            sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.IsBold = true;
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.IsItalic = true;
            sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["f" + myrow + ":f" + myrow].Text = "Кре.сч.";

            sheet.Range["g" + myrow + ":g" + myrow].Merge();
            sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
            sheet.Range["g" + myrow + ":g" + myrow].Style.Font.IsBold = true;
            sheet.Range["g" + myrow + ":g" + myrow].Style.Font.IsItalic = true;
            sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 10;
            sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["g" + myrow + ":g" + myrow].Text = "Кре.сч.";

            sheet.Range["i" + myrow + ":k" + myrow].Merge();
            sheet.Range["i" + myrow + ":k" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["i" + myrow + ":k" + myrow].Style.WrapText = true;
            sheet.Range["i" + myrow + ":k" + myrow].Style.Font.IsBold = true;
            sheet.Range["i" + myrow + ":k" + myrow].Style.Font.IsItalic = true;
            sheet.Range["i" + myrow + ":k" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["i" + myrow + ":k" + myrow].Style.Font.Size = 10;
            sheet.Range["i" + myrow + ":k" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["i" + myrow + ":k" + myrow].Text = "Сумма";

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

            var schet = "SELECT deb_sch,deb_sch_2,kre_sch,kre_sch_2,sum(summa) as summa FROM products_prixod_jur7 where kod_doc='" + kod_num_textBox.Text + "' and user='" + string_for_otdels + "' and month='" + month_global + "' and year='" + year_global + "' group by deb_sch";

            sql.myReader = sql.return_MySqlCommand(schet).ExecuteReader();
            while (sql.myReader.Read())
            {

                sheet.Range["d" + myrow + ":d" + myrow].Merge();
                sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.IsBold = true;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.IsItalic = true;
                sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 10;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["d" + myrow + ":d" + myrow].Text = sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch").ToString() : "";

                sheet.Range["e" + myrow + ":e" + myrow].Merge();
                sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.IsBold = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.IsItalic = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 10;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["e" + myrow + ":e" + myrow].Text = sql.myReader["deb_sch_2"] != DBNull.Value ? sql.myReader.GetString("deb_sch_2").ToString() : "";

                sheet.Range["f" + myrow + ":f" + myrow].Merge();
                sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.IsBold = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.IsItalic = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":f" + myrow].Text = sql.myReader["kre_sch"] != DBNull.Value ? sql.myReader.GetString("kre_sch").ToString() : "";

                sheet.Range["g" + myrow + ":g" + myrow].Merge();
                sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.IsBold = true;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.IsItalic = true;
                sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 10;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["g" + myrow + ":g" + myrow].Text = sql.myReader["kre_sch_2"] != DBNull.Value ? sql.myReader.GetString("kre_sch_2").ToString() : "";

                sheet.Range["i" + myrow + ":k" + myrow].Merge();
                sheet.Range["i" + myrow + ":k" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["i" + myrow + ":k" + myrow].Style.WrapText = true;
                sheet.Range["i" + myrow + ":k" + myrow].Style.Font.IsBold = true;
                sheet.Range["i" + myrow + ":k" + myrow].Style.Font.IsItalic = true;
                sheet.Range["i" + myrow + ":k" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["i" + myrow + ":k" + myrow].Style.Font.Size = 10;
                sheet.Range["i" + myrow + ":k" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["i" + myrow + ":k" + myrow].Value = sql.myReader["summa"] != DBNull.Value ? sql.myReader.GetString("summa").ToString() : "";



                sheet.SetRowHeight(myrow, 18);

                myrow++;
            }
            sql.myReader.Close();


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
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Получил:_____________________";

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

        private void ras_spisok_btn_Click(object sender, EventArgs e)
        {

        }

        private void ras_naklad_btn_Click(object sender, EventArgs e)
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

            sheet.Range["a2:e2"].Style.Font.IsBold = false;
            sheet.Range["a2:e2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a2:e2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a2:e2"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["a2:e2"].Style.Font.Size = 12;
            sheet.Range["a2:e2"].Merge(); // birlashtirish
            sheet.Range["a2:e2"].Text = "№ " + num_ras_textBox.Text + "от " + data_ras_DateTimePicker.Value.ToString("dd.MM.yyyy") + "       ";


            sheet.Range["b3:e3"].Style.Font.IsBold = false;
            sheet.Range["b3:e3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b3:e3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b3:e3"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["b3:e3"].Style.Font.Size = 12;
            sheet.Range["b3:e3"].Merge(); // birlashtirish
            sheet.Range["b3:e3"].Text = "к товаро-отгрузочным документам № ";

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

            sql2.myReader = sql2.return_MySqlCommand(send).ExecuteReader();
            while (sql2.myReader.Read())
            {
                sheet.Range["a5:d5"].Style.Font.IsBold = false;
                sheet.Range["a5:d5"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a5:d5"].Style.VerticalAlignment = VerticalAlignType.Bottom;
                sheet.Range["a5:d5"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a5:d5"].Style.Font.Size = 11;
                sheet.Range["a5:d5"].Merge(); // birlashtirish
                sheet.Range["a5:d5"].Text = "Поставщик :" + (sql2.myReader["naim_org"] != DBNull.Value ? sql2.myReader.GetString("naim_org") : "");
                sheet.Range["a5:d5"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);

                //naim_org,adres,ras_s,bank,inn,okxn
                sheet.Range["a6:d6"].Style.Font.IsBold = false;
                sheet.Range["a6:d6"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a6:d6"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a6:d6"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a6:d6"].Style.Font.Size = 11;
                sheet.Range["a6:d6"].Merge(); // birlashtirish
                sheet.Range["a6:d6"].Text = "Адрес :" + (sql2.myReader["adres"] != DBNull.Value ? sql2.myReader.GetString("adres") : "");
                sheet.Range["a6:d6"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);



                sheet.Range["a7:d7"].Style.Font.IsBold = false;
                sheet.Range["a7:d7"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a7:d7"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a7:d7"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a7:d7"].Style.Font.Size = 11;
                sheet.Range["a7:d7"].Merge(); // birlashtirish
                sheet.Range["a7:d7"].Text = "Р/с :" + (sql2.myReader["ras_s"] != DBNull.Value ? sql2.myReader.GetString("ras_s") : "");
                sheet.Range["a7:d7"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);



                sheet.Range["a8:d8"].Style.Font.IsBold = false;
                sheet.Range["a8:d8"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a8:d8"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a8:d8"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a8:d8"].Style.Font.Size = 11;
                sheet.Range["a8:d8"].Merge(); // birlashtirish
                sheet.Range["a8:d8"].Text = "Банк : " + (sql2.myReader["bank"] != DBNull.Value ? sql2.myReader.GetString("bank") : "");
                sheet.Range["a8:d8"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);



                sheet.Range["a9:d9"].Style.Font.IsBold = false;
                sheet.Range["a9:d9"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a9:d9"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a9:d9"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a9:d9"].Style.Font.Size = 11;
                sheet.Range["a9:d9"].Merge(); // birlashtirish
                sheet.Range["a9:d9"].Text = "МФО :" + (sql2.myReader["mfo"] != DBNull.Value ? sql2.myReader.GetString("mfo") : "");
                sheet.Range["a9:d9"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);

                sheet.Range["a10:d10"].Style.Font.IsBold = false;
                sheet.Range["a10:d10"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a10:d10"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a10:d10"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a10:d10"].Style.Font.Size = 11;
                sheet.Range["a10:d10"].Merge(); // birlashtirish
                sheet.Range["a10:d10"].Text = "ИНН : " + (sql2.myReader["inn"] != DBNull.Value ? sql2.myReader.GetString("inn") : "") + "  " + "ОКЭТ :" + (sql2.myReader["okxn"] != DBNull.Value ? sql2.myReader.GetString("okxn") : "");
                sheet.Range["a10:d10"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);
            }
            sql2.myReader.Close();


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
            sql.myReader.Close();
            var send2 = "SELECT * FROM products_rasxod_jur7 where kod_doc='" + kod_num_textBox.Text + "' and user='" + string_for_otdels + "' and month='" + month_global + "' and year='" + year_global + "' ";

            sql.myReader = sql.return_MySqlCommand(send2).ExecuteReader();
            while (sql.myReader.Read())
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
                sheet.Range["b" + myrow + ":b" + myrow].Text = sql.myReader["naim_tov"] != DBNull.Value ? sql.myReader.GetString("naim_tov") : "";
                sheet.Range["b" + myrow + ":b" + myrow].Style.WrapText = true;


                sheet.Range["c" + myrow + ":c" + myrow].Merge();
                sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 10;
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["c" + myrow + ":c" + myrow].Value = sql.myReader["edin"] != DBNull.Value ? sql.myReader.GetString("edin") : "";
                sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;

                // naim_tov,edin,kol,sena,summa
                sheet.Range["d" + myrow + ":d" + myrow].Merge();
                sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
                sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["d" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 10;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["d" + myrow + ":d" + myrow].Value = sql.myReader["kol"] != DBNull.Value ? sql.myReader.GetString("kol") : "";

                sheet.Range["e" + myrow + ":e" + myrow].Merge();
                sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 10;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["e" + myrow + ":e" + myrow].Value = sql.myReader["sena"] != DBNull.Value ? sql.myReader.GetString("sena") : "";

                all_summa += sql.myReader["summa"] != DBNull.Value ? sql.myReader.GetDouble("summa") : 0;

                sheet.Range["f" + myrow + ":f" + myrow].Merge();
                sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":f" + myrow].Value = sql.myReader["summa"] != DBNull.Value ? sql.myReader.GetString("summa") : "";

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
            sql.myReader.Close();


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

            string filePath = Environment.CurrentDirectory + "\\docs\\qrcode\\" + string_for_otdels + "-" + year_global + "-" + month_global + "-" + kod_num + "-" + "2" + ".png";
            FileInfo file = new FileInfo(filePath);
            if (file.Exists)
            {
                string picPath = Environment.CurrentDirectory + "\\docs\\qrcode\\" + string_for_otdels + "-" + year_global + "-" + month_global + "-" + kod_num + "-" + "2" + ".png";
                ExcelPicture picture = sheet.Pictures.Add(1, 8, picPath);
                picture.Width = 60;
                picture.Height = 60;

            }
            else
            {
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode((string_for_otdels + "-" + year_global + "-" + month_global + "-" + kod_num + "-" + "2"), QRCodeGenerator.ECCLevel.Q);
                QRCode qrCode = new QRCode(qrCodeData);

                Bitmap bitMap = qrCode.GetGraphic(20);

                bitMap.Save(Environment.CurrentDirectory + "\\docs\\qrcode\\" + string_for_otdels + "-" + year_global + "-" + month_global + "-" + kod_num + "-" + "2" + ".png");

                string picPath = Environment.CurrentDirectory + "\\docs\\qrcode\\" + string_for_otdels + "-" + year_global + "-" + month_global + "-" + kod_num + "-" + "2" + ".png";
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

        private void ras_schet_fac_btn_Click(object sender, EventArgs e)
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

            sheet.Range["a2:e2"].Style.Font.IsBold = false;
            sheet.Range["a2:e2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a2:e2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a2:e2"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["a2:e2"].Style.Font.Size = 12;
            sheet.Range["a2:e2"].Merge(); // birlashtirish
            sheet.Range["a2:e2"].Text = "№ " + num_ras_textBox.Text + "от " + data_ras_DateTimePicker.Value.ToString("dd.MM.yyyy") + "       ";


            sheet.Range["b3:e3"].Style.Font.IsBold = false;
            sheet.Range["b3:e3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b3:e3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b3:e3"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["b3:e3"].Style.Font.Size = 12;
            sheet.Range["b3:e3"].Merge(); // birlashtirish
            sheet.Range["b3:e3"].Text = "к товаро-отгрузочным документам № ";

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

            sql2.myReader = sql2.return_MySqlCommand(send).ExecuteReader();
            while (sql2.myReader.Read())
            {
                sheet.Range["a5:d5"].Style.Font.IsBold = false;
                sheet.Range["a5:d5"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a5:d5"].Style.VerticalAlignment = VerticalAlignType.Bottom;
                sheet.Range["a5:d5"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a5:d5"].Style.Font.Size = 11;
                sheet.Range["a5:d5"].Merge(); // birlashtirish
                sheet.Range["a5:d5"].Text = "Поставщик :" + (sql2.myReader["naim_org"] != DBNull.Value ? sql2.myReader.GetString("naim_org") : "");
                sheet.Range["a5:d5"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);

                //naim_org,adres,ras_s,bank,inn,okxn
                sheet.Range["a6:d6"].Style.Font.IsBold = false;
                sheet.Range["a6:d6"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a6:d6"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a6:d6"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a6:d6"].Style.Font.Size = 11;
                sheet.Range["a6:d6"].Merge(); // birlashtirish
                sheet.Range["a6:d6"].Text = "Адрес :" + (sql2.myReader["adres"] != DBNull.Value ? sql2.myReader.GetString("adres") : "");
                sheet.Range["a6:d6"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);



                sheet.Range["a7:d7"].Style.Font.IsBold = false;
                sheet.Range["a7:d7"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a7:d7"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a7:d7"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a7:d7"].Style.Font.Size = 11;
                sheet.Range["a7:d7"].Merge(); // birlashtirish
                sheet.Range["a7:d7"].Text = "Р/с :" + (sql2.myReader["ras_s"] != DBNull.Value ? sql2.myReader.GetString("ras_s") : "");
                sheet.Range["a7:d7"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);



                sheet.Range["a8:d8"].Style.Font.IsBold = false;
                sheet.Range["a8:d8"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a8:d8"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a8:d8"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a8:d8"].Style.Font.Size = 11;
                sheet.Range["a8:d8"].Merge(); // birlashtirish
                sheet.Range["a8:d8"].Text = "Банк : " + (sql2.myReader["bank"] != DBNull.Value ? sql2.myReader.GetString("bank") : "");
                sheet.Range["a8:d8"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);



                sheet.Range["a9:d9"].Style.Font.IsBold = false;
                sheet.Range["a9:d9"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a9:d9"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a9:d9"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a9:d9"].Style.Font.Size = 11;
                sheet.Range["a9:d9"].Merge(); // birlashtirish
                sheet.Range["a9:d9"].Text = "МФО :" + (sql2.myReader["mfo"] != DBNull.Value ? sql2.myReader.GetString("mfo") : "");
                sheet.Range["a9:d9"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);

                sheet.Range["a10:d10"].Style.Font.IsBold = false;
                sheet.Range["a10:d10"].Style.Font.FontName = "Times New Roman";
                sheet.Range["a10:d10"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a10:d10"].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["a10:d10"].Style.Font.Size = 11;
                sheet.Range["a10:d10"].Merge(); // birlashtirish
                sheet.Range["a10:d10"].Text = "ИНН : " + (sql2.myReader["inn"] != DBNull.Value ? sql2.myReader.GetString("inn") : "") + "  " + "ОКЭТ :" + (sql2.myReader["okxn"] != DBNull.Value ? sql2.myReader.GetString("okxn") : "");
                sheet.Range["a10:d10"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
                // sheet.Range["B6:B6"].BorderAround(LineStyleType.Thin);
            }
            sql2.myReader.Close();


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
            sql.myReader.Close();
            var send2 = "SELECT * FROM products_rasxod_jur7 where kod_doc='" + kod_num_textBox.Text + "' and user='" + string_for_otdels + "' and month='" + month_global + "' and year='" + year_global + "' ";

            sql.myReader = sql.return_MySqlCommand(send2).ExecuteReader();
            while (sql.myReader.Read())
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
                sheet.Range["b" + myrow + ":b" + myrow].Text = sql.myReader["naim_tov"] != DBNull.Value ? sql.myReader.GetString("naim_tov") : "";
                sheet.Range["b" + myrow + ":b" + myrow].Style.WrapText = true;


                sheet.Range["c" + myrow + ":c" + myrow].Merge();
                sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 10;
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["c" + myrow + ":c" + myrow].Value = sql.myReader["edin"] != DBNull.Value ? sql.myReader.GetString("edin") : "";
                sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;

                // naim_tov,edin,kol,sena,summa
                sheet.Range["d" + myrow + ":d" + myrow].Merge();
                sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
                sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["d" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 10;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["d" + myrow + ":d" + myrow].Value = sql.myReader["kol"] != DBNull.Value ? sql.myReader.GetString("kol") : "";

                sheet.Range["e" + myrow + ":e" + myrow].Merge();
                sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 10;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["e" + myrow + ":e" + myrow].Value = sql.myReader["sena"] != DBNull.Value ? sql.myReader.GetString("sena") : "";

                all_summa += sql.myReader["summa"] != DBNull.Value ? sql.myReader.GetDouble("summa") : 0;

                sheet.Range["f" + myrow + ":f" + myrow].Merge();
                sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":f" + myrow].Value = sql.myReader["summa"] != DBNull.Value ? sql.myReader.GetString("summa") : "";

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
            sql.myReader.Close();


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

            string filePath = Environment.CurrentDirectory + "\\docs\\qrcode\\" + string_for_otdels + "-" + year_global + "-" + month_global + "-" + kod_num + "-" + "2" + ".png";
            FileInfo file = new FileInfo(filePath);
            if (file.Exists)
            {
                string picPath = Environment.CurrentDirectory + "\\docs\\qrcode\\" + string_for_otdels + "-" + year_global + "-" + month_global + "-" + kod_num + "-" + "2" + ".png";
                ExcelPicture picture = sheet.Pictures.Add(1, 8, picPath);
                picture.Width = 60;
                picture.Height = 60;

            }
            else
            {
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode((string_for_otdels + "-" + year_global + "-" + month_global + "-" + kod_num + "-" + "2"), QRCodeGenerator.ECCLevel.Q);
                QRCode qrCode = new QRCode(qrCodeData);

                Bitmap bitMap = qrCode.GetGraphic(20);

                bitMap.Save(Environment.CurrentDirectory + "\\docs\\qrcode\\" + string_for_otdels + "-" + year_global + "-" + month_global + "-" + kod_num + "-" + "2" + ".png");

                string picPath = Environment.CurrentDirectory + "\\docs\\qrcode\\" + string_for_otdels + "-" + year_global + "-" + month_global + "-" + kod_num + "-" + "2" + ".png";
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

        private void ras_pereosenka_btn_Click(object sender, EventArgs e)
        {
            //try
            //{
            // month_global = month_textBox.Text;
            // year_global = year_textBox.Text;

            //pereotsenka pereotsenka = new pereotsenka();
            //pereotsenka.Show();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("prixod_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                podraz podraz = new podraz(string_for_otdels, year_global, month_global);
                podraz.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("prixod_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void rasxod_dataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.Exception.Message == "DataGridViewComboBoxCell value is not valid.")
            {
                object value = rasxod_dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (!((DataGridViewComboBoxColumn)rasxod_dataGridView.Columns[e.ColumnIndex]).Items.Contains(value))
                {
                    ((DataGridViewComboBoxColumn)rasxod_dataGridView.Columns[e.ColumnIndex]).Items.Add(value);
                    e.ThrowException = false;
                }
            }
        }

        private void rasxod_dataGridView_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            try
            {
                DialogResult dialogResult = MessageBox.Show("Вы действительно хотите удалить данные?", "Удаление", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    foreach (DataGridViewRow row in rasxod_dataGridView.SelectedRows)
                    {
                        if (row.Cells[0].Value != null && row.Cells[18].Value != null)
                        {

                            sql.return_MySqlCommand("delete from products_rasxod_jur7 where id = " + row.Cells[0].Value + "").ExecuteNonQuery();
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
    }
}
