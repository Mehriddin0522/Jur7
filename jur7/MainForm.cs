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
using Spire.Xls;

namespace jur7
{
    public partial class MainForm : Form
    {
        Connect sql = new Connect();
        Connect sql1 = new Connect();
        Connect sql2 = new Connect();
        Connect sql3 = new Connect();
        Connect sql4 = new Connect();
        public string month_global = "";
        public string year_global = "";

        public string string_for_otdels;
        public MainForm(string string_for_otdels)
        {
            InitializeComponent();

            sql.Connection();
            sql1.Connection();
            sql2.Connection();
            sql3.Connection();
            sql4.Connection();


            this.string_for_otdels = string_for_otdels;
            run_main();
            saldo_date();
        }

        string iznos = "";
        public void run_main()
        {
            try
            {
                var querty = " SELECT * FROM spravochnik_main_jur7 where user_jur7='" + string_for_otdels + "' ";
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


                    Console.WriteLine(iznos);
                    if (sql.myReader.GetBoolean("iznos") == true)
                    {
                        iznos_checkBox.CheckState = System.Windows.Forms.CheckState.Checked;
                        iznos = "1";
                    }
                    else
                    {
                        iznos = "0";
                    }
                    Console.WriteLine("iznos:" + iznos);

                }
                sql.myReader.Close();

            }
            catch (Exception ex)
            {
                sql.myReader.Close();
                MessageBox.Show("prixod_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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

        String getmonth_String2;
        public string set_month_name2(int getmonth)
        {
            switch (getmonth)
            {
                case 1:
                    {
                        getmonth_String2 = "Январь";
                        break;
                    }
                case 2:
                    {
                        getmonth_String2 = "Февраль";
                        break;
                    }
                case 3:
                    {
                        getmonth_String2 = "Март";
                        break;
                    }
                case 4:
                    {
                        getmonth_String2 = "Апрель";
                        break;
                    }
                case 5:
                    {
                        getmonth_String2 = "Май";
                        break;
                    }
                case 6:
                    {
                        getmonth_String2 = "Июнь";
                        break;
                    }
                case 7:
                    {
                        getmonth_String2 = "Июль";
                        break;
                    }
                case 8:
                    {
                        getmonth_String2 = "Августь";
                        break;
                    }
                case 9:
                    {
                        getmonth_String2 = "Сентябрь";
                        break;
                    }
                case 10:
                    {
                        getmonth_String2 = "Октябрь";
                        break;
                    }
                case 11:
                    {
                        getmonth_String2 = "Ноябрь";
                        break;
                    }
                case 12:
                    {
                        getmonth_String2 = "Декабрь";
                        break;
                    }
            }
            return getmonth_String2;
        }


        private void prixod_btn_Click(object sender, EventArgs e)
        {
            try
            {
                month_global = month_textBox.Text;
                year_global = year_textBox.Text;

                prixod_document prixod_document = new prixod_document(string_for_otdels, month_global, year_global, iznos);
                prixod_document.WindowState = FormWindowState.Maximized;
                prixod_document.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("prixod_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void rasxod_btn_Click(object sender, EventArgs e)
        {
            try
            {
                month_global = month_textBox.Text;
                year_global = year_textBox.Text;

                rasxod_document rasxod_document = new rasxod_document(string_for_otdels, year_global, month_global, iznos);
                rasxod_document.WindowState = FormWindowState.Maximized;
                rasxod_document.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("rasxod_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void vnut_per_btn_Click(object sender, EventArgs e)
        {
            try
            {
                month_global = month_textBox.Text;
                year_global = year_textBox.Text;

                vnut_per_document vnut_per = new vnut_per_document(string_for_otdels, year_global, month_global, iznos);
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
            try
            {
                month_global = month_textBox.Text;
                year_global = year_textBox.Text;

                podraz podraz = new podraz(string_for_otdels, year_global, month_global);
                podraz.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("mesya_ost_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void oborotka_iznos_btn_Click(object sender, EventArgs e)
        {
            try
            {
                month_global = month_textBox.Text;
                year_global = year_textBox.Text;
                Oborotka_iznos oborotka_iznos = new Oborotka_iznos(string_for_otdels, year_global, month_global);
                //oborotka_iznos.WindowState = FormWindowState.Maximized;
                oborotka_iznos.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("oborotka_iznos_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void saldo_btn_Click(object sender, EventArgs e)
        {
            try
            {
                month_global = month_textBox.Text;
                year_global = year_textBox.Text;
                Ostatok saldo = new Ostatok(string_for_otdels, year_global, month_global);
                saldo.WindowState = FormWindowState.Maximized;
                saldo.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("saldo_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void iznos_po_sub_btn_Click(object sender, EventArgs e)
        {
            //try
            //{
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.LeftMargin = 0.4;
            sheet.PageSetup.RightMargin = 0.4;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;


            sheet.PageSetup.Orientation = PageOrientationType.Portrait;


            sheet.Range["a1:a1"].ColumnWidth = 8;
            sheet.Range["b1:b1"].ColumnWidth = 10.57;
            sheet.Range["c1:c1"].ColumnWidth = 15.29;
            sheet.Range["d1:d1"].ColumnWidth = 15;
            sheet.Range["e1:e1"].ColumnWidth = 15;
            sheet.Range["f1:f1"].ColumnWidth = 15.14;
            sheet.Range["g1:g1"].ColumnWidth = 15.29;


            sheet.Range["a1:g1"].Style.Font.IsBold = true;
            //sheet.Range["a1:g1"].Style.Font.IsItalic = true;
            sheet.Range["a1:g1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a1:g1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a1:g1"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a1:g1"].Style.Font.Size = 14;
            sheet.Range["a1:g1"].Merge(); // birlashtirish
            sheet.Range["a1:g1"].Text = "Оборотка износа по субсчетам за " + set_month_name2(Convert.ToInt32(month_textBox.Text)) + "  " + year_textBox.Text + " год";
            sheet.Range["a1:g1"].Style.WrapText = true;
            sheet.Range["a1:g1"].Style.Font.Color = Color.DarkBlue;
            sheet.SetRowHeight(1, 22);
            sheet.SetRowHeight(2, 4);

            sheet.Range["a3:a3"].Style.Font.IsBold = true;
            sheet.Range["a3:a3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a3:a3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a3:a3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a3:a3"].Style.Font.Size = 11;
            sheet.Range["a3:a3"].Merge(); // birlashtirish
            sheet.Range["a3:a3"].Text = "СЧЕТ";
            sheet.Range["a3:a3"].Style.WrapText = true;
            sheet.Range["a3:a3"].BorderAround(LineStyleType.Thin);

            sheet.Range["b3:b3"].Style.Font.IsBold = true;
            sheet.Range["b3:b3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b3:b3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b3:b3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b3:b3"].Style.Font.Size = 11;
            sheet.Range["b3:b3"].Merge(); // birlashtirish
            sheet.Range["b3:b3"].Text = "СУБСЧЕТ";
            sheet.Range["b3:b3"].Style.WrapText = true;
            sheet.Range["b3:b3"].BorderAround(LineStyleType.Thin);

            sheet.Range["c3:c3"].Style.Font.IsBold = true;
            sheet.Range["c3:c3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c3:c3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c3:c3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c3:c3"].Style.Font.Size = 11;
            sheet.Range["c3:c3"].Merge(); // birlashtirish
            sheet.Range["c3:c3"].Text = "Салъдо";
            sheet.Range["c3:c3"].Style.WrapText = true;
            sheet.Range["c3:c3"].BorderAround(LineStyleType.Thin);

            sheet.Range["d3:d3"].Style.Font.IsBold = true;
            sheet.Range["d3:d3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d3:d3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d3:d3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d3:d3"].Style.Font.Size = 11;
            sheet.Range["d3:d3"].Merge(); // birlashtirish
            sheet.Range["d3:d3"].Text = "Приход";
            sheet.Range["d3:d3"].Style.WrapText = true;
            sheet.Range["d3:d3"].BorderAround(LineStyleType.Thin);

            sheet.Range["e3:e3"].Style.Font.IsBold = true;
            sheet.Range["e3:e3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e3:e3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e3:e3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e3:e3"].Style.Font.Size = 11;
            sheet.Range["e3:e3"].Merge(); // birlashtirish
            sheet.Range["e3:e3"].Text = "Расход";
            sheet.Range["e3:e3"].Style.WrapText = true;
            sheet.Range["e3:e3"].BorderAround(LineStyleType.Thin);

            sheet.Range["f3:f3"].Style.Font.IsBold = true;
            sheet.Range["f3:f3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f3:f3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f3:f3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f3:f3"].Style.Font.Size = 11;
            sheet.Range["f3:f3"].Merge(); // birlashtirish
            sheet.Range["f3:f3"].Text = "Износ";
            sheet.Range["f3:f3"].Style.WrapText = true;
            sheet.Range["f3:f3"].BorderAround(LineStyleType.Thin);

            sheet.Range["g3:g3"].Style.Font.IsBold = true;
            sheet.Range["g3:g3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g3:g3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g3:g3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g3:g3"].Style.Font.Size = 11;
            sheet.Range["g3:g3"].Merge(); // birlashtirish
            sheet.Range["g3:g3"].Text = "Салъдо";
            sheet.Range["g3:g3"].Style.WrapText = true;
            sheet.Range["g3:g3"].BorderAround(LineStyleType.Thin);

            month_global = month_textBox.Text;
            year_global = year_textBox.Text;

            int i = 0;
            int myrow = 4;
            int j = 0;

            double saldo_sum_all = 0;
            double pri_sum_all = 0;
            double ras_sum_all = 0;
            double summa_iznos_all = 0;

            string schet = "";
            string sub_schet = "";

            double saldo_sum = 0;
            double pri_sum = 0;
            double ras_sum = 0;
            double summa_iznos = 0;


            var exl = "SELECT distinct schet,subschet FROM gruppa_jur7 ";

            sql.myReader = sql.return_MySqlCommand(exl).ExecuteReader();
            while (sql.myReader.Read())
            {
                schet = (sql.myReader["schet"] != DBNull.Value ? sql.myReader.GetString("schet") : "");
                sub_schet = (sql.myReader["subschet"] != DBNull.Value ? sql.myReader.GetString("subschet") : "");

                var saldo = " select id,sum(summa) as summa,summa_iznos from saldo_jur7 where user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' and deb_sch='" + schet + "' and deb_sch_2='" + sub_schet + "' ";

                sql1.myReader = sql1.return_MySqlCommand(saldo).ExecuteReader();
                while (sql1.myReader.Read())
                {
                    saldo_sum = (sql1.myReader["summa"] != DBNull.Value ? sql1.myReader.GetDouble("summa") : 0);
                    summa_iznos = (sql1.myReader["summa_iznos"] != DBNull.Value ? sql1.myReader.GetDouble("summa_iznos") : 0);

                }
                sql1.myReader.Close();

                saldo_sum_all += saldo_sum;
                summa_iznos_all += summa_iznos;

                var pri = " SELECT sum(summa) as summa FROM products_jur7 where vid_doc = '1' and user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' and deb_sch='" + schet + "' and deb_sch_2='" + sub_schet + "' ";

                sql1.myReader = sql1.return_MySqlCommand(pri).ExecuteReader();
                while (sql1.myReader.Read())
                {
                    pri_sum = (sql1.myReader["summa"] != DBNull.Value ? sql1.myReader.GetDouble("summa") : 0);

                }
                sql1.myReader.Close();

                pri_sum_all += pri_sum;

                var ras = " SELECT sum(summa) as summa FROM products_jur7 where vid_doc = '2' and user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' and deb_sch='" + schet + "' and deb_sch_2='" + sub_schet + "'";

                sql1.myReader = sql1.return_MySqlCommand(ras).ExecuteReader();
                while (sql1.myReader.Read())
                {
                    ras_sum = (sql1.myReader["summa"] != DBNull.Value ? sql1.myReader.GetDouble("summa") : 0);

                }
                sql1.myReader.Close();

                ras_sum_all += ras_sum;

                j = i;
                j = j + 1;

                sheet.Range["a" + myrow + ":a" + myrow].Merge();
                sheet.Range["a" + myrow + ":a" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["a" + myrow + ":a" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.Size = 11;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["a" + myrow + ":a" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["a" + myrow + ":a" + myrow].Text = schet;

                sheet.Range["b" + myrow + ":b" + myrow].Merge();
                sheet.Range["b" + myrow + ":b" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["b" + myrow + ":b" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["b" + myrow + ":b" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.Size = 11;
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["b" + myrow + ":b" + myrow].Text = sub_schet;
                sheet.Range["b" + myrow + ":b" + myrow].Style.WrapText = true;


                sheet.Range["c" + myrow + ":c" + myrow].Merge();
                sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 11;
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["c" + myrow + ":c" + myrow].Value = saldo_sum.ToString();
                sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;

                sheet.Range["d" + myrow + ":d" + myrow].Merge();
                sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
                sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["d" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 11;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["d" + myrow + ":d" + myrow].Value = pri_sum.ToString();

                sheet.Range["e" + myrow + ":e" + myrow].Merge();
                sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 11;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["e" + myrow + ":e" + myrow].Value = ras_sum.ToString();

                sheet.Range["f" + myrow + ":f" + myrow].Merge();
                sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 11;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":f" + myrow].Value = summa_iznos.ToString();

                sheet.Range["g" + myrow + ":g" + myrow].Merge();
                sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
                sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["g" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 11;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["g" + myrow + ":g" + myrow].Value = (saldo_sum - ras_sum + pri_sum + summa_iznos).ToString();


                myrow = myrow + 1;
                i = i + 1;

            }
            sql.myReader.Close();

            sheet.Range["a" + myrow + ":b" + myrow].Merge();
            sheet.Range["a" + myrow + ":b" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["a" + myrow + ":b" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a" + myrow + ":b" + myrow].BorderAround(LineStyleType.Thin);
            sheet.Range["a" + myrow + ":b" + myrow].Style.Font.Size = 11;
            sheet.Range["a" + myrow + ":b" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["a" + myrow + ":b" + myrow].Text = "Итого :";
            sheet.Range["a" + myrow + ":b" + myrow].Style.WrapText = true;
            //sheet.Range["a" + myrow + ":b" + myrow].Style.Font.IsBold = true;


            sheet.Range["c" + myrow + ":c" + myrow].Merge();
            sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
            sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["c" + myrow + ":c" + myrow].Value = saldo_sum_all.ToString();
            sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;
            //sheet.Range["c" + myrow + ":c" + myrow].Style.Font.IsBold = true;

            sheet.Range["d" + myrow + ":d" + myrow].Merge();
            sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
            //sheet.Range["d" + myrow + ":d" + myrow].Style.Font.IsBold = true;
            sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);
            sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 11;
            sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["d" + myrow + ":d" + myrow].Value = pri_sum_all.ToString();

            sheet.Range["e" + myrow + ":e" + myrow].Merge();
            sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
            //sheet.Range["e" + myrow + ":e" + myrow].Style.Font.IsBold = true;
            sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 11;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["e" + myrow + ":e" + myrow].Value = ras_sum_all.ToString();//(saldo_sum_all - pri_sum_all + ras_sum_all).ToString();

            sheet.Range["f" + myrow + ":f" + myrow].Merge();
            sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
            //sheet.Range["f" + myrow + ":f" + myrow].Style.Font.IsBold = true;
            sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 11;
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["f" + myrow + ":f" + myrow].Value = summa_iznos_all.ToString();//(saldo_sum_all - pri_sum_all + ras_sum_all).ToString();

            sheet.Range["g" + myrow + ":g" + myrow].Merge();
            sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
            //sheet.Range["g" + myrow + ":g" + myrow].Style.Font.IsBold = true;
            sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);
            sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 11;
            sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["g" + myrow + ":g" + myrow].Value = (saldo_sum_all + pri_sum_all - ras_sum_all + summa_iznos_all).ToString();


            myrow++;
            myrow++;

            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.IsBold = true;
            //sheet.Range["b" + myrow + ":d" + myrow].Style.Font.IsItalic = true;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.Size = 14;
            sheet.Range["b" + myrow + ":d" + myrow].Merge(); // birlashtirish
            sheet.Range["b" + myrow + ":d" + myrow].Text = "Гл.бухгалтер __________________";
            sheet.Range["b" + myrow + ":d" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":d" + myrow].Style.Font.Color = Color.DarkBlue;

            sheet.Range["e" + myrow + ":g" + myrow].Style.Font.IsBold = true;
            //sheet.Range["e" + myrow + ":g" + myrow].Style.Font.IsItalic = true;
            sheet.Range["e" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["e" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e" + myrow + ":g" + myrow].Style.Font.Size = 14;
            sheet.Range["e" + myrow + ":g" + myrow].Merge(); // birlashtirish
            sheet.Range["e" + myrow + ":g" + myrow].Text = "Бухгалтер __________________";
            sheet.Range["e" + myrow + ":g" + myrow].Style.WrapText = true;
            sheet.Range["e" + myrow + ":g" + myrow].Style.Font.Color = Color.DarkBlue;

            sheet.Range["b3:" + myrow + "g"].NumberFormat = "#,##0.00";


            workbook.SaveToFile(Environment.CurrentDirectory + "\\docs\\Износ по субсче.xlsx", Spire.Xls.FileFormat.Version2007);
            System.Diagnostics.Process.Start(workbook.FileName);



            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Журнал-ордер_excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        int exist_saldo = 0;
        private void MainForm_Load(object sender, EventArgs e)
        {
            main_form_dateTimePicker.Value = DateTime.Now;
            prixod_btn.Focus();
            saldo_exist();
        }

        public void saldo_exist()
        {
            try
            {
                var saldo = " select exists ( select * from saldo_jur7 where user='" + string_for_otdels + "' and month='" + month_textBox.Text + "' and year='" + year_textBox.Text + "' ) as ex ";
                sql.myReader = sql.return_MySqlCommand(saldo).ExecuteReader();
                while (sql.myReader.Read())
                {
                    exist_saldo = sql.myReader.GetInt32("ex");
                }
                sql.myReader.Close();

                if (exist_saldo == 0)
                {
                    register_saldo_btn.BackColor = Color.YellowGreen;
                }
                else
                {
                    register_saldo_btn.BackColor = Color.Gainsboro;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("saldo_exist " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void obnavit_toolStripButton_Click(object sender, EventArgs e)
        {
            register_main();
        }

        public void register_main()
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
                        var query = "select * from spravochnik_main_jur7 where user_jur7 = '" + string_for_otdels + "' ";
                        sql.myReader = sql.return_MySqlCommand(query).ExecuteReader();
                        if (sql.myReader.Read())
                        {
                            var update_sklad = "update spravochnik_main_jur7 set " +
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
                            //run_alert("");
                        }
                        else
                        {
                            sql1.return_MySqlCommand("insert into spravochnik_main_jur7 (user_jur7,naim_org,adres,telefon,ras_s,bank,mfo,gorod,inn,fio_gl_bugalter,fio_bugalter,inspektor,okxn,inven,schet,schet1,month,year,iznos) values(" +
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
                            //run_alert("");
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
                saldo_exist();
                year_textBox.TextChanged -= year_textBox_TextChanged;


            }

            saldo_exist();
            month_global = month_textBox.Text;
            year_global = year_textBox.Text;
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
            saldo_exist();
            month_global = month_textBox.Text;
            year_global = year_textBox.Text;
        }

        private void spravochnik_toolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                Spravochnik spravochnik = new Spravochnik();

                spravochnik.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("prixod_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
                    register_main();
                    run_main();
                }
                else
                {
                    register_main();
                    run_main();
                    register_saldo();
                    run_alert("Успешно");
                    saldo_exist();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("register_saldo_btn_Click " + ex.Message);
            }
        }



        string podraz_1 = "";
        int exist = 0;
        public string year_saldo = "";
        public string month_saldo = "";

        public void saldo_date()
        {
            if (Int32.Parse(month_textBox.Text.ToString()) > 0 && Int32.Parse(month_textBox.Text.ToString()) <= 12)
            {
                //month_textBox.TextChanged -= month_textBox_TextChanged;

                int num = 0;
                Int32.TryParse(month_textBox.Text.ToString(), out num);

                if (num > 1)
                {
                    month_saldo = (num - 1).ToString();
                    year_saldo = year_textBox.Text;
                }
                else
                if (num == 1)
                {
                    month_saldo = 12.ToString();


                    int year = 0;
                    Int32.TryParse(year_textBox.Text.ToString(), out year);


                    year_saldo = (year - 1).ToString();

                    //year_textBox.TextChanged -= year_textBox_TextChanged;
                }

                //insert_dannie_function();
                // month_textBox.TextChanged += month_textBox_TextChanged;
            }
            else if (Int32.Parse(month_textBox.Text.ToString()) > 0 && Int32.Parse(month_textBox.Text.ToString()) == 12)
            {
                //year_textBox.TextChanged -= year_textBox_TextChanged;


            }
        }

        public string refresh_strings_to_mysql(string mystring)
        {
            string str = string.Format("{0:#0.00}", Convert.ToDouble(mystring.Replace('.', ','))).Replace(',', '.');
            Console.WriteLine(str);
            return str;
        }

        public void register_saldo()
        {
            try
            {
                saldo_date();

                sql.myReader = sql.return_MySqlCommand(" select exists ( select * from saldo_jur7 where user='" + string_for_otdels + "' and month='" + month_textBox.Text + "' and year='" + year_textBox.Text + "' ) as ex ").ExecuteReader();
                while (sql.myReader.Read())
                {
                    exist = sql.myReader.GetInt32("ex");
                }
                sql.myReader.Close();

                if (exist == 0)
                {
                    var querty = " SELECT podraz_naim FROM podraz_jur7 group by podraz_naim ";
                    sql.myReader = sql.return_MySqlCommand(querty).ExecuteReader();
                    while (sql.myReader.Read())
                    {
                        podraz_1 = sql.myReader["podraz_naim"] != DBNull.Value ? sql.myReader.GetString("podraz_naim") : "";

                        string podraz_2 = "";
                        var querty2 = " SELECT fio FROM podraz_jur7 where podraz_naim='" + podraz_1 + "' ";
                        sql1.myReader = sql1.return_MySqlCommand(querty2).ExecuteReader();
                        while (sql1.myReader.Read())
                        {
                            podraz_2 = sql1.myReader["fio"] != DBNull.Value ? sql1.myReader.GetString("fio") : "";

                            var products = "  select t.id,t.vid_doc,t.kod_doc,t.product_id,t.gruppa,t.naim_tov,t.edin,t.inventar_num,t.seria_num,sum(t.kol) as kol,t.sena,sum(t.summa) as summa,t.deb_sch,t.deb_sch_2,t.kre_sch,t.kre_sch_2,t.provodka_iznos," +
                                             " t.provodka_iznos_2,t.summa_iznos,t.date_pr from(" +
                                             " SELECT id, vid_doc, kod_doc, product_id, gruppa, naim_tov, edin, inventar_num, seria_num, sum(kol) as kol, sena, sum(summa) as summa, deb_sch, deb_sch_2, kre_sch, kre_sch_2, provodka_iznos, provodka_iznos_2," +
                                             " summa_iznos, date_pr FROM products_jur7 where vid_doc = '1' and user = '" + string_for_otdels + "' and year = '" + year_saldo + "' and month = '" + month_saldo + "' and kol > 0 and komu_1 = '" + podraz_1 + "' and komu_2 = '" + podraz_2 + "' group by product_id" +
                                             " union all" +
                                             " SELECT id, vid_doc, kod_doc, product_id, gruppa, naim_tov, edin, inventar_num, seria_num, sum(kol) as kol, sena, summa, deb_sch, deb_sch_2, kre_sch, kre_sch_2, provodka_iznos, provodka_iznos_2," +
                                             " summa_iznos, date_pr FROM products_jur7 where vid_doc = '3' and user = '" + string_for_otdels + "' and year = '" + year_saldo + "' and month = '" + month_saldo + "' and kol > 0 and komu_1 = '" + podraz_1 + "' and komu_2 = '" + podraz_2 + "' group by product_id" +
                                             " ) as t where t.kol > 0 group by t.product_id" +
                                             " union all" +
                                             " select id, '' as vid_doc,'' as kod_doc,product_id,gruppa, naim_tov, edin, inventar_num, seria_num,kol,sena,summa,deb_sch, deb_sch_2, kre_sch, kre_sch_2, '' as provodka_iznos, '' as provodka_iznos_2,summa_iznos," +
                                              " data_pr from saldo_jur7 where kol > 0 and user = '" + string_for_otdels + "' and year = '" + year_saldo + "' and month = '" + month_saldo + "' and kol > 0 and podraz_1 = '" + podraz_1 + "' and podraz_2 = '" + podraz_2 + "' ";


                            sql2.myReader = sql2.return_MySqlCommand(products).ExecuteReader();
                            while (sql2.myReader.Read())
                            {

                                int product_id = 0;
                                string gruppa = "";
                                string naim_tov = "";
                                string edin = "";
                                string inventar_num = "";
                                string seria_num = "";
                                double kol = 0;
                                double sena = 0;
                                double summa = 0;
                                string deb_sch = "";
                                string deb_sch_2 = "";
                                string kre_sch = "";
                                string kre_sch_2 = "";
                                double summa_iznos = 0;
                                double summa_iznos_2 = 0;
                                double prosent_izn = 0;
                                string date_pr;

                                double pri_kol = 0;
                                double ras_kol = 0;
                                double vnut_ras_kol = 0;
                                double vnut_pri_kol = 0;


                                product_id = sql2.myReader["product_id"] != DBNull.Value ? sql2.myReader.GetInt32("product_id") : 0;
                                gruppa = sql2.myReader["gruppa"] != DBNull.Value ? sql2.myReader.GetString("gruppa") : "";

                                naim_tov = sql2.myReader["naim_tov"] != DBNull.Value ? sql2.myReader.GetString("naim_tov") : "";
                                edin = sql2.myReader["edin"] != DBNull.Value ? sql2.myReader.GetString("edin") : "";
                                inventar_num = sql2.myReader["inventar_num"] != DBNull.Value ? sql2.myReader.GetString("inventar_num") : "";
                                seria_num = sql2.myReader["seria_num"] != DBNull.Value ? sql2.myReader.GetString("seria_num") : "";

                                ///////////
                                pri_kol = (sql2.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql2.myReader.GetString("kol"))) : 0);

                                var products_pri = " SELECT sum(kol) as kol FROM products_jur7 where vid_doc='2' and product_id='" + (sql2.myReader["product_id"] != DBNull.Value ? sql2.myReader.GetString("product_id") : "") + "' and user='" + string_for_otdels + "' and year='" + year_saldo + "' and month='" + month_saldo + "' and ot_kogo='" + podraz_1 + "' and ot_kogo_2='" + podraz_2 + "' and kol > 0 ";

                                sql3.myReader = sql3.return_MySqlCommand(products_pri).ExecuteReader();

                                while (sql3.myReader.Read())
                                {
                                    ras_kol = (sql3.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql3.myReader.GetString("kol"))) : 0);
                                }
                                sql3.myReader.Close();

                                var products_vnut = " SELECT id,sum(kol) as kol FROM products_jur7 where vid_doc = '3' and product_id='" + (sql2.myReader["product_id"] != DBNull.Value ? sql2.myReader.GetString("product_id") : "") + "' and user = '" + string_for_otdels + "' and year = '" + year_saldo + "' and month = '" + month_saldo + "' and kol > 0 and komu_1 = '" + podraz_1 + "' and komu_2 = '" + podraz_2 + "' group by product_id";

                                sql3.myReader = sql3.return_MySqlCommand(products_vnut).ExecuteReader();

                                while (sql3.myReader.Read())
                                {
                                    vnut_ras_kol = (sql3.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql3.myReader.GetString("kol"))) : 0);
                                }
                                sql3.myReader.Close();

                                var products_vnut_ras = " SELECT id,sum(kol) as kol FROM products_jur7 where vid_doc = '3' and product_id='" + (sql2.myReader["product_id"] != DBNull.Value ? sql2.myReader.GetString("product_id") : "") + "' and user = '" + string_for_otdels + "' and year = '" + year_saldo + "' and month = '" + month_saldo + "' and kol > 0 and ot_kogo = '" + podraz_1 + "' and ot_kogo_2 = '" + podraz_2 + "' group by product_id";

                                sql3.myReader = sql3.return_MySqlCommand(products_vnut_ras).ExecuteReader();

                                while (sql3.myReader.Read())
                                {
                                    vnut_pri_kol = (sql3.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql3.myReader.GetString("kol").Replace(",", "."))) : 0);
                                }
                                sql3.myReader.Close();
                                //////////
                                kol = (pri_kol - ras_kol - vnut_pri_kol);

                                sena = (sql2.myReader["sena"] != DBNull.Value ? sql2.myReader.GetDouble("sena") : 0);
                                summa = kol * sena;


                                deb_sch = sql2.myReader["deb_sch"] != DBNull.Value ? sql2.myReader.GetString("deb_sch") : "";
                                deb_sch_2 = sql2.myReader["deb_sch_2"] != DBNull.Value ? sql2.myReader.GetString("deb_sch_2") : "";
                                kre_sch = sql2.myReader["kre_sch"] != DBNull.Value ? sql2.myReader.GetString("kre_sch") : "";
                                kre_sch_2 = sql2.myReader["kre_sch_2"] != DBNull.Value ? sql2.myReader.GetString("kre_sch_2") : "";
                                date_pr = (sql2.myReader["date_pr"] != DBNull.Value ? (DateTime.Parse(sql2.myReader.GetString("date_pr")).ToString("yyyy-MM-dd")) : null);
                                summa_iznos_2 = sql2.myReader["summa_iznos"] != DBNull.Value ? sql2.myReader.GetDouble("summa_iznos") : 0;
                                int date_pr_month = 0;
                                int date_pr_year = 0;
                                int kol_month = 0;
                                int sub_year = 0;

                                if (iznos == "1")
                                {
                                    DateTime now = DateTime.Parse(sql2.myReader.GetString("date_pr"));
                                    date_pr_month = now.Month;
                                    date_pr_year = now.Year;

                                    sub_year = (Convert.ToInt32(year_saldo) - date_pr_year);
                                    if (sub_year == 0)
                                    {
                                        kol_month = Convert.ToInt32(month_saldo) + 1 - date_pr_month;
                                    }
                                    else
                                    {
                                        kol_month = Convert.ToInt32(month_saldo) + 1 - date_pr_month + 12 * sub_year;
                                    }

                                    sql3.myReader = sql3.return_MySqlCommand(" SELECT * FROM gruppa_jur7 where kod_gruppa='" + gruppa + "' ").ExecuteReader();
                                    while (sql3.myReader.Read())
                                    {
                                        prosent_izn = sql3.myReader["prosent_izn"] != DBNull.Value ? Convert.ToDouble(sql3.myReader.GetString("prosent_izn")) : 0;

                                    }
                                    sql3.myReader.Close();

                                    if (prosent_izn == 0)
                                    {
                                        summa_iznos = 0;
                                    }
                                    else
                                    {
                                        if (kol_month == 0)
                                        {
                                            kol_month = 1;
                                        }
                                        summa_iznos = (summa / 12) * kol_month;
                                    }
                                }
                                else
                                {
                                    summa_iznos = summa_iznos_2;
                                }

                                if (summa_iznos > summa)
                                {
                                    summa_iznos = summa;
                                }


                                var insert_product = "insert into saldo_jur7 (data_saldo,product_id,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,deb_sch,deb_sch_2,kre_sch,kre_sch_2,summa_iznos,data_pr,year,month,user,podraz_1,podraz_2) values(" +
                                                   "'" + main_form_dateTimePicker.Value.ToString("yyyy-MM-dd") + "'," +
                                                   "'" + (product_id) + "'," +
                                                   "'" + (gruppa) + "'," +
                                                   "'" + (naim_tov) + "'," +
                                                   "'" + (edin) + "'," +
                                                   "'" + (inventar_num) + "'," +
                                                   "'" + (seria_num) + "'," +
                                                   "'" + refresh_strings_to_mysql(kol.ToString()) + "'," +
                                                   "'" + refresh_strings_to_mysql(sena.ToString()) + "'," +
                                                   "'" + refresh_strings_to_mysql(summa.ToString()) + "'," +
                                                   "'" + (deb_sch) + "'," +
                                                   "'" + (deb_sch_2) + "'," +
                                                   "'" + (kre_sch) + "'," +
                                                   "'" + (kre_sch_2) + "'," +
                                                   "'" + refresh_strings_to_mysql(summa_iznos.ToString()) + "'," +
                                                   "'" + (date_pr) + "'," +
                                                   "'" + (year_textBox.Text) + "'," +
                                                   "'" + (month_textBox.Text) + "'," +
                                                   "'" + (string_for_otdels) + "'," +
                                                   "'" + (podraz_1) + "'," +
                                                   "'" + (podraz_2) + "' " +
                                                   ")";
                                sql4.return_MySqlCommand(insert_product).ExecuteNonQuery();

                            }
                            sql2.myReader.Close();

                        }
                        sql1.myReader.Close();

                    }
                    sql.myReader.Close();

                }
                else
                {
                    var delete = " delete from saldo_jur7 where user='" + string_for_otdels + "' and year='" + year_textBox.Text + "' and month='" + month_textBox.Text + "' ";
                    sql1.return_MySqlCommand(delete).ExecuteNonQuery();

                    var querty = " SELECT podraz_naim FROM podraz_jur7 group by podraz_naim ";
                    sql.myReader = sql.return_MySqlCommand(querty).ExecuteReader();
                    while (sql.myReader.Read())
                    {
                        podraz_1 = sql.myReader["podraz_naim"] != DBNull.Value ? sql.myReader.GetString("podraz_naim") : "";

                        string podraz_2 = "";
                        var querty2 = " SELECT fio FROM podraz_jur7 where podraz_naim='" + podraz_1 + "' ";
                        sql1.myReader = sql1.return_MySqlCommand(querty2).ExecuteReader();
                        while (sql1.myReader.Read())
                        {
                            podraz_2 = sql1.myReader["fio"] != DBNull.Value ? sql1.myReader.GetString("fio") : "";


                            var products = "  select t.id,t.vid_doc,t.kod_doc,t.product_id,t.gruppa,t.naim_tov,t.edin,t.inventar_num,t.seria_num,sum(t.kol) as kol,t.sena,sum(t.summa) as summa,t.deb_sch,t.deb_sch_2,t.kre_sch,t.kre_sch_2,t.provodka_iznos," +
                                             " t.provodka_iznos_2,t.summa_iznos,t.date_pr from(" +
                                             " SELECT id, vid_doc, kod_doc, product_id, gruppa, naim_tov, edin, inventar_num, seria_num, sum(kol) as kol, sena, sum(summa) as summa, deb_sch, deb_sch_2, kre_sch, kre_sch_2, provodka_iznos, provodka_iznos_2," +
                                             " summa_iznos, date_pr FROM products_jur7 where vid_doc = '1' and user = '" + string_for_otdels + "' and year = '" + year_saldo + "' and month = '" + month_saldo + "' and kol > 0 and komu_1 = '" + podraz_1 + "' and komu_2 = '" + podraz_2 + "' group by product_id" +
                                             " union all" +
                                             " SELECT id, vid_doc, kod_doc, product_id, gruppa, naim_tov, edin, inventar_num, seria_num, sum(kol) as kol, sena, summa, deb_sch, deb_sch_2, kre_sch, kre_sch_2, provodka_iznos, provodka_iznos_2," +
                                             " summa_iznos, date_pr FROM products_jur7 where vid_doc = '3' and user = '" + string_for_otdels + "' and year = '" + year_saldo + "' and month = '" + month_saldo + "' and kol > 0 and komu_1 = '" + podraz_1 + "' and komu_2 = '" + podraz_2 + "' group by product_id" +
                                             " ) as t where t.kol > 0 group by t.product_id" +
                                             " union all" +
                                             " select id, '' as vid_doc,'' as kod_doc,product_id,gruppa, naim_tov, edin, inventar_num, seria_num,kol,sena,summa,deb_sch, deb_sch_2, kre_sch, kre_sch_2, '' as provodka_iznos, '' as provodka_iznos_2,summa_iznos," +
                                              " data_pr from saldo_jur7 where kol > 0 and user = '" + string_for_otdels + "' and year = '" + year_saldo + "' and month = '" + month_saldo + "' and kol > 0 and podraz_1 = '" + podraz_1 + "' and podraz_2 = '" + podraz_2 + "' ";


                            sql2.myReader = sql2.return_MySqlCommand(products).ExecuteReader();
                            while (sql2.myReader.Read())
                            {

                                int product_id = 0;
                                string gruppa = "";
                                string naim_tov = "";
                                string edin = "";
                                string inventar_num = "";
                                string seria_num = "";
                                double kol = 0;
                                double sena = 0;
                                double summa = 0;
                                string deb_sch = "";
                                string deb_sch_2 = "";
                                string kre_sch = "";
                                string kre_sch_2 = "";
                                double summa_iznos = 0;
                                double summa_iznos_2 = 0;
                                double prosent_izn = 0;
                                string date_pr;

                                double pri_kol = 0;
                                double ras_kol = 0;
                                double vnut_ras_kol = 0;
                                double vnut_pri_kol = 0;

                                product_id = sql2.myReader["product_id"] != DBNull.Value ? sql2.myReader.GetInt32("product_id") : 0;
                                gruppa = sql2.myReader["gruppa"] != DBNull.Value ? sql2.myReader.GetString("gruppa") : "";
                                naim_tov = sql2.myReader["naim_tov"] != DBNull.Value ? sql2.myReader.GetString("naim_tov") : "";
                                edin = sql2.myReader["edin"] != DBNull.Value ? sql2.myReader.GetString("edin") : "";
                                inventar_num = sql2.myReader["inventar_num"] != DBNull.Value ? sql2.myReader.GetString("inventar_num") : "";
                                seria_num = sql2.myReader["seria_num"] != DBNull.Value ? sql2.myReader.GetString("seria_num") : "";

                                ///////////
                                pri_kol = (sql2.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql2.myReader.GetString("kol"))) : 0);

                                var products_pri = " SELECT sum(kol) as kol FROM products_jur7 where vid_doc='2' and product_id='" + (sql2.myReader["product_id"] != DBNull.Value ? sql2.myReader.GetString("product_id") : "") + "' and user='" + string_for_otdels + "' and year='" + year_saldo + "' and month='" + month_saldo + "' and ot_kogo='" + podraz_1 + "' and ot_kogo_2='" + podraz_2 + "' and kol > 0 ";

                                sql3.myReader = sql3.return_MySqlCommand(products_pri).ExecuteReader();

                                while (sql3.myReader.Read())
                                {
                                    ras_kol = (sql3.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql3.myReader.GetString("kol"))) : 0);
                                }
                                sql3.myReader.Close();

                                var products_vnut = " SELECT id,sum(kol) as kol FROM products_jur7 where vid_doc = '3' and product_id='" + (sql2.myReader["product_id"] != DBNull.Value ? sql2.myReader.GetString("product_id") : "") + "' and user = '" + string_for_otdels + "' and year = '" + year_saldo + "' and month = '" + month_saldo + "' and kol > 0 and komu_1 = '" + podraz_1 + "' and komu_2 = '" + podraz_2 + "' group by product_id";

                                sql3.myReader = sql3.return_MySqlCommand(products_vnut).ExecuteReader();

                                while (sql3.myReader.Read())
                                {
                                    vnut_ras_kol = (sql3.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql3.myReader.GetString("kol"))) : 0);
                                }
                                sql3.myReader.Close();

                                var products_vnut_ras = " SELECT id,sum(kol) as kol FROM products_jur7 where vid_doc = '3' and product_id='" + (sql2.myReader["product_id"] != DBNull.Value ? sql2.myReader.GetString("product_id") : "") + "' and user = '" + string_for_otdels + "' and year = '" + year_saldo + "' and month = '" + month_saldo + "' and kol > 0 and ot_kogo = '" + podraz_1 + "' and ot_kogo_2 = '" + podraz_2 + "' group by product_id";

                                sql3.myReader = sql3.return_MySqlCommand(products_vnut_ras).ExecuteReader();

                                while (sql3.myReader.Read())
                                {
                                    vnut_pri_kol = (sql3.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql3.myReader.GetString("kol"))) : 0);
                                }
                                sql3.myReader.Close();
                                //////////
                                kol = (pri_kol - ras_kol - vnut_pri_kol);

                                sena = (sql2.myReader["sena"] != DBNull.Value ? sql2.myReader.GetDouble("sena") : 0);
                                summa = kol * sena;

                                deb_sch = sql2.myReader["deb_sch"] != DBNull.Value ? sql2.myReader.GetString("deb_sch") : "";
                                deb_sch_2 = sql2.myReader["deb_sch_2"] != DBNull.Value ? sql2.myReader.GetString("deb_sch_2") : "";
                                kre_sch = sql2.myReader["kre_sch"] != DBNull.Value ? sql2.myReader.GetString("kre_sch") : "";
                                kre_sch_2 = sql2.myReader["kre_sch_2"] != DBNull.Value ? sql2.myReader.GetString("kre_sch_2") : "";

                                //summa_iznos = sql2.myReader["summa_iznos"] != DBNull.Value ? sql2.myReader.GetDouble("summa_iznos") : 0;
                                date_pr = (sql2.myReader["date_pr"] != DBNull.Value ? (DateTime.Parse(sql2.myReader.GetString("date_pr")).ToString("yyyy-MM-dd")) : null);

                                int date_pr_month = 0;
                                int date_pr_year = 0;
                                int kol_month = 0;
                                int sub_year = 0;

                                if (iznos == "1")
                                {
                                    DateTime now = DateTime.Parse(sql2.myReader.GetString("date_pr"));
                                    date_pr_month = now.Month;
                                    date_pr_year = now.Year;

                                    sub_year = (Convert.ToInt32(year_saldo) - date_pr_year);
                                    if (sub_year == 0)
                                    {
                                        kol_month = Convert.ToInt32(month_saldo) + 1 - date_pr_month;
                                    }
                                    else
                                    {
                                        kol_month = Convert.ToInt32(month_saldo) + 1 - date_pr_month + 12 * sub_year;
                                    }

                                    sql3.myReader = sql3.return_MySqlCommand(" SELECT * FROM gruppa_jur7 where kod_gruppa='" + gruppa + "' ").ExecuteReader();
                                    while (sql3.myReader.Read())
                                    {
                                        prosent_izn = sql3.myReader["prosent_izn"] != DBNull.Value ? Convert.ToDouble(sql3.myReader.GetString("prosent_izn")) : 0;

                                    }
                                    sql3.myReader.Close();

                                    if (prosent_izn == 0)
                                    {
                                        summa_iznos = 0;
                                    }
                                    else
                                    {
                                        if (kol_month == 0)
                                        {
                                            kol_month = 1;
                                        }
                                        summa_iznos = (summa / 12) * kol_month;
                                    }
                                }
                                else
                                {
                                    summa_iznos = summa_iznos_2;
                                }

                                if (summa_iznos > summa)
                                {
                                    summa_iznos = summa;
                                }



                                var insert_product = "insert into saldo_jur7 (data_saldo,product_id,gruppa,naim_tov,edin,inventar_num,seria_num,kol,sena,summa,deb_sch,deb_sch_2,kre_sch,kre_sch_2,summa_iznos,data_pr,year,month,user,podraz_1,podraz_2) values(" +
                                                   "'" + main_form_dateTimePicker.Value.ToString("yyyy-MM-dd") + "'," +
                                                   "'" + (product_id) + "'," +
                                                   "'" + (gruppa) + "'," +
                                                   "'" + (naim_tov) + "'," +
                                                   "'" + (edin) + "'," +
                                                   "'" + (inventar_num) + "'," +
                                                   "'" + (seria_num) + "'," +
                                                   "'" + refresh_strings_to_mysql(kol.ToString()) + "'," +
                                                   "'" + refresh_strings_to_mysql(sena.ToString()) + "'," +
                                                   "'" + refresh_strings_to_mysql(summa.ToString()) + "'," +
                                                   "'" + (deb_sch) + "'," +
                                                   "'" + (deb_sch_2) + "'," +
                                                   "'" + (kre_sch) + "'," +
                                                   "'" + (kre_sch_2) + "'," +
                                                   "'" + refresh_strings_to_mysql(summa_iznos.ToString()) + "'," +
                                                   "'" + (date_pr) + "'," +
                                                   "'" + (year_textBox.Text) + "'," +
                                                   "'" + (month_textBox.Text) + "'," +
                                                   "'" + (string_for_otdels) + "'," +
                                                   "'" + (podraz_1) + "'," +
                                                   "'" + (podraz_2) + "' " +
                                                   ")";
                                sql4.return_MySqlCommand(insert_product).ExecuteNonQuery();

                            }
                            sql2.myReader.Close();

                        }
                        sql1.myReader.Close();

                    }
                    sql.myReader.Close();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("saldo_exist " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void qrcode_btn_Click(object sender, EventArgs e)
        {

        }

        private void qr_code_btn_Click(object sender, EventArgs e)
        {
            try
            {
                month_global = month_textBox.Text;
                year_global = year_textBox.Text;

                QrCode qrcode = new QrCode(string_for_otdels, year_global, month_global);
                qrcode.WindowState = FormWindowState.Maximized;
                qrcode.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("vnut_per_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
