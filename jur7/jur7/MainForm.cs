using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Globalization;
using System.Threading;
using Tulpep.NotificationWindow;

namespace jur7
{
    public partial class MainForm : Form
    {
        public static Connect sql = new Connect();
        public static Connect sql1 = new Connect();
        public static Connect sql2 = new Connect();
        public static Connect sql3 = new Connect();


        public string string_for_otdels;
        public string iznos;
        public MainForm(string string_for_otdels)
        {
            InitializeComponent();

            sql.Connection();
            sql1.Connection();
            sql2.Connection();
            sql3.Connection();

            this.string_for_otdels = string_for_otdels;
            run_main();
        }

        public void run_main()
        {
            //try
            //{
                var querty = " SELECT * FROM spravochnik_main where user_jur7='" + string_for_otdels + "' ";
                sql.myReader = sql.return_MySqlCommand(querty).ExecuteReader();
                while (sql.myReader.Read())
                {
                    //naim_org,adres,telefon,ras_s,bank,mfo,gorod,inn,fio_gl_bugalter,fio_bugalter,inspektor,okxn,inven,schet,schet1,iznos

                    naim_org_textBox.Text = (sql.myReader["naim_org"] != DBNull.Value ? sql.myReader.GetString("naim_org") : "");
                    adres_textBox.Text = (sql.myReader["adres"] != DBNull.Value ? sql.myReader.GetString("adres") : "");
                    tel_num_textBox.Text = (sql.myReader["telefon"] != DBNull.Value ? sql.myReader.GetString("telefon") : "");
                    ras_s_textBox.Text = (sql.myReader["ras_s"] != DBNull.Value ? sql.myReader.GetString("ras_s") : "");
                    bank_textBox.Text = (sql.myReader["bank"] != DBNull.Value ? sql.myReader.GetString("bank") : "");
                    mfo_textBox.Text = (sql.myReader["mfo"] != DBNull.Value ? sql.myReader.GetString("mfo") : "");
                    gorod_textBox.Text = (sql.myReader["gorod"] != DBNull.Value ? sql.myReader.GetString("gorod") : "");
                    inn_textBox.Text = (sql.myReader["inn"] != DBNull.Value ? sql.myReader.GetString("inn") : "");
                    gl_bux_textBox.Text = (sql.myReader["fio_gl_bugalter"] != DBNull.Value ? sql.myReader.GetString("fio_gl_bugalter") : "");
                    buxga_textBox.Text = (sql.myReader["fio_bugalter"] != DBNull.Value ? sql.myReader.GetString("fio_bugalter") : "");
                    insp_textBox.Text = (sql.myReader["inspektor"] != DBNull.Value ? sql.myReader.GetString("inspektor") : "");
                    okxn_textBox.Text = (sql.myReader["okxn"] != DBNull.Value ? sql.myReader.GetString("okxn") : "");
                    inven_textBox.Text = (sql.myReader["inven"] != DBNull.Value ? sql.myReader.GetString("inven") : "");
                    schet_textBox.Text = (sql.myReader["schet"] != DBNull.Value ? sql.myReader.GetString("schet") : "");
                    schet_2_textBox.Text = (sql.myReader["schet1"] != DBNull.Value ? sql.myReader.GetString("schet1") : "");
                    month_textBox.Text = (sql.myReader["month"] != DBNull.Value ? sql.myReader.GetString("month") : "");
                    year_textBox.Text = (sql.myReader["year"] != DBNull.Value ? sql.myReader.GetString("year") : "");

                    
                    if (sql.myReader.GetBoolean("iznos") == true)
                    {
                        iznos_checkBox.CheckState = CheckState.Checked;
                        iznos = "1";
                }
                else
                {
                    iznos = "0";
                }

                }
                sql.myReader.Close();

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("prixod_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }


        PopupNotifier popup = new PopupNotifier();
        public void run_alert(string fio)
        {
            popup.BodyColor = Color.FromArgb(116, 209, 106);
            // popup.BorderColor = Color.White;
            popup.ContentHoverColor = Color.Black;
            popup.TitleColor = Color.White;
            popup.ContentColor = Color.White;

            popup.TitleText = "Успешно";
            // popup.ContentText = fio;

            popup.TitleFont = new Font("Times New Roman", 12f);
            popup.Popup();
        }

        string month_global = "";
        string year_global = "";
        private void prixod_btn_Click(object sender, EventArgs e)
        {
            //try
            //{
                month_global = month_textBox.Text;
                year_global = year_textBox.Text;

                prixod_document prixod_document = new prixod_document(string_for_otdels, year_global, month_global,iznos);
                prixod_document.WindowState = FormWindowState.Maximized;
                prixod_document.Show();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("prixod_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void rasxod_btn_Click(object sender, EventArgs e)
        {
            //try
            //{
                month_global = month_textBox.Text;
                year_global = year_textBox.Text;

                rasxod_document rasxod_document = new rasxod_document(string_for_otdels, year_global, month_global);
                rasxod_document.WindowState = FormWindowState.Maximized;
                rasxod_document.Show();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("rasxod_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void vnut_per_btn_Click(object sender, EventArgs e)
        {
            try
            {
                month_global = month_textBox.Text;
                year_global = year_textBox.Text;

                vnut_per_document vnut_per = new vnut_per_document(string_for_otdels, year_global, month_global);
                vnut_per.WindowState = FormWindowState.Maximized;
                vnut_per.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("vnut_per_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void formirovanie_btn_Click(object sender, EventArgs e)
        {
            try
            {
                month_global = month_textBox.Text;
                year_global = year_textBox.Text;
                formirovanie_jur7 formirovanie_jur7 = new formirovanie_jur7(string_for_otdels, year_global, month_global);
                //formirovanie_jur7.WindowState = FormWindowState.Maximized;
                formirovanie_jur7.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("formirovanie_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void mesya_ost_btn_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    month_ost month_ost = new month_ost();
            //    //month_ost.WindowState = FormWindowState.Maximized;
            //    month_ost.Show();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("mesya_ost_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void oborotka_iznos_btn_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    oborotka_iznos oborotka_iznos = new oborotka_iznos();
            //    //oborotka_iznos.WindowState = FormWindowState.Maximized;
            //    oborotka_iznos.Show();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("oborotka_iznos_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void saldo_btn_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    saldo saldo = new saldo();
            //    //saldo.WindowState = FormWindowState.Maximized;
            //    saldo.Show();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("saldo_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void iznos_po_sub_btn_Click(object sender, EventArgs e)
        {

        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            main_form_dateTimePicker.Value = DateTime.Now;
            prixod_btn.Focus();
        }

        private void obnavit_toolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (naim_org_textBox.Text != "")
                {

                    string iznos = "";

                    if (iznos_checkBox.Checked)
                    {
                        iznos = "1";
                    }
                    else
                    {
                        iznos = "0";
                    }

                    try
                    {
                        var query = "select * from spravochnik_main where user_jur7 = '" + string_for_otdels + "' ";
                        sql.myReader = sql.return_MySqlCommand(query).ExecuteReader();
                        if (sql.myReader.Read())
                        {
                            var update_sklad = "update spravochnik_main set " +
                                            "user_jur7 = '" + string_for_otdels + "'," +
                                            "naim_org = '" + naim_org_textBox.Text + "'," +
                                            "adres = '" + adres_textBox.Text + "'," +
                                            "telefon = '" + tel_num_textBox.Text + "'," +
                                            "ras_s = '" + ras_s_textBox.Text + "'," +
                                            "bank = '" + bank_textBox.Text + "'," +
                                            "mfo = '" + mfo_textBox.Text + "'," +
                                            "gorod = '" + gorod_textBox.Text + "'," +
                                            "inn = '" + inn_textBox.Text + "'," +
                                            "fio_gl_bugalter = '" + gl_bux_textBox.Text + "'," +
                                            "fio_bugalter = '" + buxga_textBox.Text + "'," +
                                            "inspektor = '" + insp_textBox.Text + "'," +
                                            "okxn = '" + okxn_textBox.Text + "'," +
                                            "inven = '" + inven_textBox.Text + "'," +
                                            "schet = '" + schet_textBox.Text + "'," +
                                            "schet1 = '" + schet_2_textBox.Text + "'," +
                                            "month = '" + month_textBox.Text + "'," +
                                            "year = '" + year_textBox.Text + "'," +
                                            "iznos = '" + iznos + "'" +

                                           " where user_jur7 = '" + string_for_otdels + "'";
                            sql1.return_MySqlCommand(update_sklad).ExecuteNonQuery();
                            run_alert("");

                        }
                        else
                        {


                            sql1.return_MySqlCommand("insert into spravochnik_main (user_jur7,naim_org,adres,telefon,ras_s,bank,mfo,gorod,inn,fio_gl_bugalter,fio_bugalter,inspektor,okxn,inven,schet,schet1,month,year,iznos) values(" +
                                                     "'" + (string_for_otdels) + "'," +
                                                     "'" + (naim_org_textBox.Text) + "'," +
                                                      "'" + (adres_textBox.Text) + "'," +
                                                    //(main_form_dateTimePicker.Text != DBNull.Value.ToString() ? "'" + DateTime.Parse(main_form_dateTimePicker.Text).ToString("yyyy-MM-dd") + "'" : "NULL") + "," +
                                                    "'" + (tel_num_textBox.Text) + "'," +
                                                    "'" + (ras_s_textBox.Text) + "'," +
                                                    "'" + (bank_textBox.Text) + "'," +
                                                    "'" + (mfo_textBox.Text) + "'," +
                                                    "'" + (gorod_textBox.Text) + "'," +
                                                    "'" + (inn_textBox.Text) + "'," +
                                                    "'" + (gl_bux_textBox.Text) + "'," +
                                                    "'" + (buxga_textBox.Text) + "'," +
                                                    "'" + (insp_textBox.Text) + "'," +
                                                    "'" + (okxn_textBox.Text) + "'," +
                                                    "'" + (inven_textBox.Text) + "'," +
                                                    "'" + (schet_textBox.Text) + "'," +
                                                    "'" + (schet_2_textBox.Text) + "'," +
                                                    "'" + (month_textBox.Text) + "'," +
                                                    "'" + (year_textBox.Text) + "'," +
                                                    "'" + (iznos) + "'" +
                                          "); ").ExecuteNonQuery();
                            run_alert("");
                        }
                        sql.myReader.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("obnavit_toolStripButton_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("prixod_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void plus_btn_Click(object sender, EventArgs e)
        {
            if (Int32.Parse(month_textBox.Text.ToString()) > 0 && Int32.Parse(month_textBox.Text.ToString()) < 12)
            {
                month_textBox.TextChanged -= month_textBox_TextChanged;


                int num = 0;
                Int32.TryParse(month_textBox.Text.ToString(), out num);

                if (num >= 1)
                {
                    month_textBox.Text = (num + 1).ToString();
                    if (num == 12)
                    {
                        month_textBox.Text = 1.ToString();
                    }
                }

                //insert_dannie_function();
                month_textBox.TextChanged += month_textBox_TextChanged;

            }
            else if (Int32.Parse(month_textBox.Text.ToString()) > 0 && Int32.Parse(month_textBox.Text.ToString()) == 12)
            {
                year_textBox.TextChanged -= year_textBox_TextChanged;

                month_textBox.Text = 1.ToString();
                int year = 0;
                Int32.TryParse(year_textBox.Text.ToString(), out year);

                year_textBox.Text = (year + 1).ToString();

                year_textBox.TextChanged -= year_textBox_TextChanged;

            }
        }

        private void month_textBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void year_textBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void minus_btn_Click(object sender, EventArgs e)
        {
            if (Int32.Parse(month_textBox.Text.ToString()) > 0 && Int32.Parse(month_textBox.Text.ToString()) <= 12)
            {
                month_textBox.TextChanged -= month_textBox_TextChanged;

                int num = 0;
                Int32.TryParse(month_textBox.Text.ToString(), out num);

                if (num > 1)
                {
                    month_textBox.Text = (num - 1).ToString();

                }
                else
                if (num == 1)
                {
                    month_textBox.Text = 12.ToString();


                    int year = 0;
                    Int32.TryParse(year_textBox.Text.ToString(), out year);

                    year_textBox.Text = (year - 1).ToString();

                    year_textBox.TextChanged -= year_textBox_TextChanged;
                }
                //insert_dannie_function();
                month_textBox.TextChanged += month_textBox_TextChanged;
            }
            else if (Int32.Parse(month_textBox.Text.ToString()) > 0 && Int32.Parse(month_textBox.Text.ToString()) == 12)
            {
                year_textBox.TextChanged -= year_textBox_TextChanged;


            }
        }

        private void spravochnik_toolStripButton_Click(object sender, EventArgs e)
        {
            //try
            //{

            Spravochnik spravochnik = new Spravochnik();

            spravochnik.Show();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("prixod_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //try
            //{
            month_global = month_textBox.Text;
            year_global = year_textBox.Text;

            qrcode qrcode = new qrcode(string_for_otdels, year_global, month_global);
            qrcode.WindowState = FormWindowState.Maximized;
            qrcode.Show();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("prixod_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void register_saldo_btn_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult response = MessageBox.Show("Вы уверены? ", "Регистрация",
                     MessageBoxButtons.YesNo,
                     MessageBoxIcon.Question,
                     MessageBoxDefaultButton.Button2);

                if (response == DialogResult.No)
                {

                }
                else
                {
                    
                }

            }
            catch (Exception ex)
            {
                
                MessageBox.Show("register_saldo_btn_Click " + ex.Message);
            }
        }
    }
}
