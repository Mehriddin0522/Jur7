using Spire.Xls;
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
    public partial class formirovanie_jur7 : Form
    {
        Connect sql = new Connect();
        Connect sql1 = new Connect();
        Connect sql2 = new Connect();

        public string string_for_otdels;
        public string month_global;
        public string year_global;

        Number_To_Words_russian number_russian = new Number_To_Words_russian();

        public formirovanie_jur7(string string_for_otdels, string year_global, string month_global)
        {
            InitializeComponent();

            sql.Connection();
            sql2.Connection();
            sql1.Connection();

            this.string_for_otdels = string_for_otdels;
            this.month_global = month_global;
            this.year_global = year_global;

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


        private void prixod_btn_Click(object sender, EventArgs e)
        {
            //try
            //{
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.LeftMargin = 0.3;
            sheet.PageSetup.RightMargin = 0.3;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;


            sheet.PageSetup.Orientation = PageOrientationType.Portrait;


            sheet.Range["a1:a1"].ColumnWidth = 9;
            sheet.Range["b1:b1"].ColumnWidth = 9.29;
            sheet.Range["c1:c1"].ColumnWidth = 24.57;
            sheet.Range["d1:d1"].ColumnWidth = 24.43;
            sheet.Range["e1:e1"].ColumnWidth = 7;
            sheet.Range["f1:f1"].ColumnWidth = 7.71;
            sheet.Range["g1:g1"].ColumnWidth = 15;

            string name_organ = "";
            var name_org = "SELECT * FROM spravochnik_main where user_jur7='" + string_for_otdels + "'";

            sql.myReader = sql.return_MySqlCommand(name_org).ExecuteReader();
            while (sql.myReader.Read())
            {
                name_organ = (sql.myReader["naim_org"] != DBNull.Value ? sql.myReader.GetString("naim_org") : "");
            }
            sql.myReader.Close();


            sheet.Range["a1:f1"].Style.Font.IsBold = true;
            sheet.Range["a1:f1"].Style.Font.IsItalic = true;
            sheet.Range["a1:f1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a1:f1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a1:f1"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["a1:f1"].Style.Font.Size = 12;
            sheet.Range["a1:f1"].Merge(); // birlashtirish
            sheet.Range["a1:f1"].Text = name_organ;
            sheet.Range["a1:f1"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["a1:f1"].Style.Font.Underline = FontUnderlineType.Single;
            sheet.SetRowHeight(1, 18);


            sheet.Range["a2:g2"].Style.Font.IsBold = true;
            sheet.Range["a2:g2"].Style.Font.IsItalic = true;
            sheet.Range["a2:g2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a2:g2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a2:g2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a2:g2"].Style.Font.Size = 14;
            sheet.Range["a2:g2"].Merge(); // birlashtirish
            sheet.Range["a2:g2"].Text = "Журнал-ордер №7";
            sheet.Range["a2:g2"].Style.WrapText = true;
            sheet.Range["a2:g2"].Style.Font.Color = Color.DarkBlue;
            sheet.SetRowHeight(2, 20);

            sheet.Range["a3:g3"].Style.Font.IsBold = true;
            sheet.Range["a3:g3"].Style.Font.IsItalic = true;
            sheet.Range["a3:g3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a3:g3"].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["a3:g3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a3:g3"].Style.Font.Size = 11;
            sheet.Range["a3:g3"].Merge(); // birlashtirish
            sheet.Range["a3:g3"].Text = "За "+ set_month_name2(Convert.ToInt32(month_global))+" "+year_global+ " год";//"За Май 2021 год";
            sheet.Range["a3:g3"].Style.WrapText = true;
            sheet.SetRowHeight(3, 18);

            sheet.Range["a4:a4"].Style.Font.IsBold = true;
            sheet.Range["a4:a4"].Style.Font.IsItalic = true;
            sheet.Range["a4:a4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a4:a4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a4:a4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a4:a4"].Style.Font.Size = 11;
            sheet.Range["a4:a4"].Merge(); // birlashtirish
            sheet.Range["a4:a4"].Text = "№ док";
            sheet.Range["a4:a4"].Style.WrapText = true;
            sheet.Range["a4:a4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["a4:a4"].BorderAround(LineStyleType.Thin);

            sheet.Range["b4:b4"].Style.Font.IsBold = true;
            sheet.Range["b4:b4"].Style.Font.IsItalic = true;
            sheet.Range["b4:b4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b4:b4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b4:b4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b4:b4"].Style.Font.Size = 11;
            sheet.Range["b4:b4"].Merge(); // birlashtirish
            sheet.Range["b4:b4"].Text = "Дата";
            sheet.Range["b4:b4"].Style.WrapText = true;
            sheet.Range["b4:b4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["b4:b4"].BorderAround(LineStyleType.Thin);

            sheet.Range["c4:c4"].Style.Font.IsBold = true;
            sheet.Range["c4:c4"].Style.Font.IsItalic = true;
            sheet.Range["c4:c4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c4:c4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c4:c4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c4:c4"].Style.Font.Size = 11;
            sheet.Range["c4:c4"].Merge(); // birlashtirish
            sheet.Range["c4:c4"].Text = "Отпустил";
            sheet.Range["c4:c4"].Style.WrapText = true;
            sheet.Range["c4:c4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["c4:c4"].BorderAround(LineStyleType.Thin);

            sheet.Range["d4:d4"].Style.Font.IsBold = true;
            sheet.Range["d4:d4"].Style.Font.IsItalic = true;
            sheet.Range["d4:d4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d4:d4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d4:d4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d4:d4"].Style.Font.Size = 11;
            sheet.Range["d4:d4"].Merge(); // birlashtirish
            sheet.Range["d4:d4"].Text = "Получил";
            sheet.Range["d4:d4"].Style.WrapText = true;
            sheet.Range["d4:d4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["d4:d4"].BorderAround(LineStyleType.Thin);

            sheet.Range["e4:e4"].Style.Font.IsBold = true;
            sheet.Range["e4:e4"].Style.Font.IsItalic = true;
            sheet.Range["e4:e4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e4:e4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e4:e4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e4:e4"].Style.Font.Size = 11;
            sheet.Range["e4:e4"].Merge(); // birlashtirish
            sheet.Range["e4:e4"].Text = "Дебит";
            sheet.Range["e4:e4"].Style.WrapText = true;
            sheet.Range["e4:e4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["e4:e4"].BorderAround(LineStyleType.Thin);

            sheet.Range["f4:f4"].Style.Font.IsBold = true;
            sheet.Range["f4:f4"].Style.Font.IsItalic = true;
            sheet.Range["f4:f4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f4:f4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f4:f4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f4:f4"].Style.Font.Size = 11;
            sheet.Range["f4:f4"].Merge(); // birlashtirish
            sheet.Range["f4:f4"].Text = "Кредит";
            sheet.Range["f4:f4"].Style.WrapText = true;
            sheet.Range["f4:f4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["f4:f4"].BorderAround(LineStyleType.Thin);

            sheet.Range["g4:g4"].Style.Font.IsBold = true;
            sheet.Range["g4:g4"].Style.Font.IsItalic = true;
            sheet.Range["g4:g4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g4:g4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g4:g4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g4:g4"].Style.Font.Size = 11;
            sheet.Range["g4:g4"].Merge(); // birlashtirish
            sheet.Range["g4:g4"].Text = "Сумма";
            sheet.Range["g4:g4"].Style.WrapText = true;
            sheet.Range["g4:g4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["g4:g4"].BorderAround(LineStyleType.Thin);
            sheet.SetRowHeight(4, 18);

            int i = 0;
            int myrow = 5;
            int j = 0;

            double all_kol_count = 0;

            var top = " SELECT id,user,year,month,vid_doc,kod_doc,date_doc,ot_kogo_2,komu_1,deb_sch,kre_sch,summa FROM products_rasxod where user='" + string_for_otdels+"' and year='"+year_global+"' and month='"+month_global+"' " +
                      "  union all" +
                      "  SELECT id,user,year,month,vid_doc,kod_doc,date_doc,ot_kogo_2,komu_2,deb_sch,kre_sch,summa FROM products_vnut_per where user = '" + string_for_otdels+"' and year = '"+year_global+"' and month = '"+month_global+"' ";

            sql.myReader = sql.return_MySqlCommand(top).ExecuteReader();
            while (sql.myReader.Read())
            {
                j = i;
                j = j + 1;

                sheet.Range["a" + myrow + ":a" + myrow].Merge();
                sheet.Range["a" + myrow + ":a" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["a" + myrow + ":a" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.Size = 11;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["a" + myrow + ":a" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["a" + myrow + ":a" + myrow].Text = sql.myReader["kod_doc"] != DBNull.Value ? sql.myReader.GetString("kod_doc").ToString() : "";

                sheet.Range["b" + myrow + ":b" + myrow].Merge();
                sheet.Range["b" + myrow + ":b" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["b" + myrow + ":b" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["b" + myrow + ":b" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.Size = 11;
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["b" + myrow + ":b" + myrow].Text = (Convert.ToDateTime(sql.myReader["date_doc"] != DBNull.Value ? sql.myReader.GetString("date_doc").ToString() : "").ToString("dd.MM.yyyy"));
                sheet.Range["b" + myrow + ":b" + myrow].Style.WrapText = true;


                sheet.Range["c" + myrow + ":c" + myrow].Merge();
                sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 11;
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["c" + myrow + ":c" + myrow].Text = sql.myReader["ot_kogo_2"] != DBNull.Value ? sql.myReader.GetString("ot_kogo_2").ToString() : "";
                sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;

                sheet.Range["d" + myrow + ":d" + myrow].Merge();
                sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
                sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["d" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 11;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["d" + myrow + ":d" + myrow].Text = sql.myReader["komu_1"] != DBNull.Value ? sql.myReader.GetString("komu_1").ToString() : "";

                sheet.Range["e" + myrow + ":e" + myrow].Merge();
                sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 11;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["e" + myrow + ":e" + myrow].Text = sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch").ToString() : "";

                sheet.Range["f" + myrow + ":f" + myrow].Merge();
                sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 11;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":f" + myrow].Text = sql.myReader["kre_sch"] != DBNull.Value ? sql.myReader.GetString("kre_sch").ToString() : "";

                sheet.Range["g" + myrow + ":g" + myrow].Merge();
                sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
                sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["g" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 11;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["g" + myrow + ":g" + myrow].Value = sql.myReader["summa"] != DBNull.Value ? sql.myReader.GetString("summa").ToString() : "";



                myrow = myrow + 1;
                i = i + 1;


            }

            sql.myReader.Close();


            myrow++;

            double all_kredit_sum = 0;
            var itog = "select t.kre_sch as kre_sch,sum(t.summa) as summa from(" +
                       " SELECT SUBSTRING(kre_sch, 1, 2) as kre_sch,sum(summa) as summa FROM products_rasxod where user = '"+string_for_otdels+"' and year = '"+year_global+"' and month = '"+month_global+"' group by SUBSTRING(kre_sch, 1, 2)"+
                       " union all"+
                       " SELECT SUBSTRING(kre_sch, 1, 2) as kre_sch,sum(summa) as summa FROM products_vnut_per where user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' group by SUBSTRING(kre_sch, 1, 2)" +
                       " ) as t group by t.kre_sch";

            sql.myReader = sql.return_MySqlCommand(itog).ExecuteReader();
            while (sql.myReader.Read())
            {
                all_kredit_sum+= (sql.myReader["summa"] != DBNull.Value ? sql.myReader.GetDouble("summa") : 0);

                sheet.Range["d" + myrow + ":e" + myrow].Merge();
                sheet.Range["d" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["d" + myrow + ":e" + myrow].Style.WrapText = true;
                //sheet.Range["d" + myrow + ":e" + myrow].Style.Font.IsBold = true;
                sheet.Range["d" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["d" + myrow + ":e" + myrow].Style.Font.Size = 11;
                sheet.Range["d" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["d" + myrow + ":e" + myrow].Text = "Итого для 'Кредит' : " + (sql.myReader["kre_sch"] != DBNull.Value ? sql.myReader.GetString("kre_sch") : "") + " ";
                //sheet.Range["d" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);

                sheet.Range["f" + myrow + ":g" + myrow].Merge();
                sheet.Range["f" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["f" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":g" + myrow].Style.Font.Size = 11;
                sheet.Range["f" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":g" + myrow].Value = (sql.myReader["summa"] != DBNull.Value ? sql.myReader.GetString("summa") : "");
                sheet.Range["f" + myrow + ":g" + myrow].Style.WrapText = true;
                //sheet.Range["f" + myrow + ":g" + myrow].Style.Font.IsBold = true;
                //sheet.Range["f" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);

                myrow++;
            }
            sql.myReader.Close();
            myrow++;

            sheet.Range["d" + myrow + ":e" + myrow].Merge();
            sheet.Range["d" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["d" + myrow + ":e" + myrow].Style.WrapText = true;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.IsBold = true;
            sheet.Range["d" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.Size = 11;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["d" + myrow + ":e" + myrow].Text = "ИТОГО : ";
            //sheet.Range["d" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);

            sheet.Range["f" + myrow + ":g" + myrow].Merge();
            sheet.Range["f" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["f" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.Size = 11;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["f" + myrow + ":g" + myrow].Value = all_kredit_sum.ToString(); ;
            sheet.Range["f" + myrow + ":g" + myrow].Style.WrapText = true;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.IsBold = true;

            myrow++;
            myrow++;

            sheet.Range["b" + myrow + ":c" + myrow].Merge();
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Гл.бухгалтер:________________";
            sheet.SetRowHeight(myrow, 18);

            myrow++;

            sheet.Range["b" + myrow + ":c" + myrow].Merge();
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Бухгалтер: __________________";
            sheet.SetRowHeight(myrow, 18);


            sheet.Range["d5:" + myrow + "k"].NumberFormat = "#,##0.00";


            workbook.SaveToFile(Environment.CurrentDirectory + "\\docs\\Журнал-ордер.xlsx", Spire.Xls.FileFormat.Version2007);
            System.Diagnostics.Process.Start(workbook.FileName);



            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Журнал-ордер_excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void button15_Click(object sender, EventArgs e)
        {
            //try
            //{
            // month_global = month_textBox.Text;
            // year_global = year_textBox.Text;

            gruppa_tovar tovar = new gruppa_tovar();
            tovar.Show();
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
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.LeftMargin = 0.3;
            sheet.PageSetup.RightMargin = 0.3;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;


            sheet.PageSetup.Orientation = PageOrientationType.Landscape;


            sheet.Range["a1:a1"].ColumnWidth = 3;
            sheet.Range["b1:b1"].ColumnWidth = 16;
            sheet.Range["c1:c1"].ColumnWidth = 13.14;
            sheet.Range["d1:d1"].ColumnWidth = 11.86;
            sheet.Range["e1:e1"].ColumnWidth = 11.86;
            sheet.Range["f1:f1"].ColumnWidth = 11.86;
            sheet.Range["g1:g1"].ColumnWidth = 11.86;
            sheet.Range["h1:h1"].ColumnWidth = 11.86;
            sheet.Range["i1:i1"].ColumnWidth = 11.86;
            sheet.Range["j1:j1"].ColumnWidth = 11.86;
            sheet.Range["k1:k1"].ColumnWidth = 11.86;
            sheet.Range["l1:l1"].ColumnWidth = 11.86;


            sheet.Range["b1:g1"].Style.Font.IsBold = true;
            sheet.Range["b1:g1"].Style.Font.IsItalic = true;
            sheet.Range["b1:g1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b1:g1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b1:g1"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["b1:g1"].Style.Font.Size = 11;
            sheet.Range["b1:g1"].Merge(); // birlashtirish
            sheet.Range["b1:g1"].Text = "СВОДНАЯ ОБОРОТЪ ЗА Декабръ 2020 год";
            sheet.Range["b1:g1"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["b1:g1"].Style.Font.Underline = FontUnderlineType.Single;

            sheet.Range["h1:i1"].Style.Font.IsBold = true;
            sheet.Range["h1:i1"].Style.Font.IsItalic = true;
            sheet.Range["h1:i1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h1:i1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h1:i1"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["h1:i1"].Style.Font.Size = 11;
            sheet.Range["h1:i1"].Merge(); // birlashtirish
            sheet.Range["h1:i1"].Text = "ГУВД1";
            sheet.Range["h1:i1"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["h1:i1"].Style.Font.Underline = FontUnderlineType.Single;
            sheet.SetRowHeight(1, 18);


            sheet.Range["b2:g2"].Style.Font.IsBold = true;
            sheet.Range["b2:g2"].Style.Font.IsItalic = true;
            sheet.Range["b2:g2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b2:g2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b2:g2"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b2:g2"].Style.Font.Size = 12;
            sheet.Range["b2:g2"].Merge(); // birlashtirish
            sheet.Range["b2:g2"].Text = " Талмут 01";
            sheet.Range["b2:g2"].Style.WrapText = true;
            sheet.Range["b2:g2"].Style.Font.Color = Color.DarkBlue;
            sheet.SetRowHeight(2, 20);

            sheet.Range["a3:c3"].Style.Font.IsBold = true;
            //sheet.Range["a3:c3"].Style.Font.IsItalic = true;
            sheet.Range["a3:c3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a3:c3"].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["a3:c3"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a3:c3"].Style.Font.Size = 10;
            sheet.Range["a3:c3"].Merge(); // birlashtirish
            sheet.Range["a3:c3"].Text = "   Счет 010";
            sheet.Range["a3:c3"].Style.WrapText = true;
            sheet.SetRowHeight(3, 18);

            // sheet.Range["a4:a5"].Style.Font.IsBold = true;
            //sheet.Range["a4:a5"].Style.Font.IsItalic = true;
            sheet.Range["a4:a5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a4:a5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a4:a5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a4:a5"].Style.Font.Size = 10;
            sheet.Range["a4:a5"].Merge(); // birlashtirish
            sheet.Range["a4:a5"].Text = "№ п.п";
            sheet.Range["a4:a5"].Style.WrapText = true;
            //heet.Range["a4:a5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["a4:a5"].BorderAround(LineStyleType.Thin);

            //sheet.Range["b4:b5"].Style.Font.IsBold = true;
            //sheet.Range["b4:b5"].Style.Font.IsItalic = true;
            sheet.Range["b4:b5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b4:b5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b4:b5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            //sheet.Range["b4:b5"].Style.Font.Size = 10;
            sheet.Range["b4:b5"].Merge(); // birlashtirish
            sheet.Range["b4:b5"].Text = "ФИО";
            sheet.Range["b4:b5"].Style.WrapText = true;
            //sheet.Range["b4:b5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["b4:b5"].BorderAround(LineStyleType.Thin);

            //sheet.Range["c4:c5"].Style.Font.IsBold = true;
            // sheet.Range["c4:c5"].Style.Font.IsItalic = true;
            sheet.Range["c4:c5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c4:c5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c4:c5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c4:c5"].Style.Font.Size = 10;
            sheet.Range["c4:c5"].Merge(); // birlashtirish
            sheet.Range["c4:c5"].Text = "Подраздел.";
            sheet.Range["c4:c5"].Style.WrapText = true;
            // sheet.Range["c4:c5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["c4:c5"].BorderAround(LineStyleType.Thin);

            //sheet.Range["d4:g4"].Style.Font.IsBold = true;
            //sheet.Range["d4:g4"].Style.Font.IsItalic = true;
            sheet.Range["d4:g4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d4:g4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d4:g4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d4:g4"].Style.Font.Size = 10;
            sheet.Range["d4:g4"].Merge(); // birlashtirish
            sheet.Range["d4:g4"].Text = "Материалный оборот";
            sheet.Range["d4:g4"].Style.WrapText = true;
            sheet.Range["d4:g4"].BorderAround(LineStyleType.Thin);


            //sheet.Range["d5:d5"].Style.Font.IsBold = true;
            //sheet.Range["d5:d5"].Style.Font.IsItalic = true;
            sheet.Range["d5:d5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d5:d5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d5:d5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d5:d5"].Style.Font.Size = 10;
            sheet.Range["d5:d5"].Merge(); // birlashtirish
            sheet.Range["d5:d5"].Text = "Нач.Салъдо";
            sheet.Range["d5:d5"].Style.WrapText = true;
            //        sheet.Range["d5:d5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["d5:d5"].BorderAround(LineStyleType.Thin);

            //sheet.Range["e5:e5"].Style.Font.IsBold = true;
            //sheet.Range["e5:e5"].Style.Font.IsItalic = true;
            sheet.Range["e5:e5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e5:e5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e5:e5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e5:e5"].Style.Font.Size = 10;
            sheet.Range["e5:e5"].Merge(); // birlashtirish
            sheet.Range["e5:e5"].Text = "Приход";
            sheet.Range["e5:e5"].Style.WrapText = true;
            //sheet.Range["e5:e5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["e5:e5"].BorderAround(LineStyleType.Thin);

            //sheet.Range["f5:f5"].Style.Font.IsBold = true;
            //sheet.Range["f5:f5"].Style.Font.IsItalic = true;
            sheet.Range["f5:f5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f5:f5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f5:f5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f5:f5"].Style.Font.Size = 10;
            sheet.Range["f5:f5"].Merge(); // birlashtirish
            sheet.Range["f5:f5"].Text = "Расход";
            sheet.Range["f5:f5"].Style.WrapText = true;
            //sheet.Range["f5:f5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["f5:f5"].BorderAround(LineStyleType.Thin);

            //sheet.Range["g5:g5"].Style.Font.IsBold = true;
            //sheet.Range["g5:g5"].Style.Font.IsItalic = true;
            sheet.Range["g5:g5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g5:g5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g5:g5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g5:g5"].Style.Font.Size = 10;
            sheet.Range["g5:g5"].Merge(); // birlashtirish
            sheet.Range["g5:g5"].Text = "Остаток";
            sheet.Range["g5:g5"].Style.WrapText = true;
            //sheet.Range["g5:g5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["g5:g5"].BorderAround(LineStyleType.Thin);
            //sheet.SetRowHeight(4, 18);


            //sheet.Range["h4:l4"].Style.Font.IsBold = true;
            //sheet.Range["h4:l4"].Style.Font.IsItalic = true;
            sheet.Range["h4:l4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h4:l4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h4:l4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["h4:l4"].Style.Font.Size = 10;
            sheet.Range["h4:l4"].Merge(); // birlashtirish
            sheet.Range["h4:l4"].Text = "Износ оборот";
            sheet.Range["h4:l4"].Style.WrapText = true;
            sheet.Range["h4:l4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["h5:h5"].Style.Font.IsBold = true;
            //sheet.Range["h5:h5"].Style.Font.IsItalic = true;
            sheet.Range["h5:h5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h5:h5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h5:h5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["h5:h5"].Style.Font.Size = 10;
            sheet.Range["h5:h5"].Merge(); // birlashtirish
            sheet.Range["h5:h5"].Text = "Салдо изн.";
            sheet.Range["h5:h5"].Style.WrapText = true;
            //sheet.Range["h5:h5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["h5:h5"].BorderAround(LineStyleType.Thin);

            //sheet.Range["i5:i5"].Style.Font.IsBold = true;
            //sheet.Range["i5:i5"].Style.Font.IsItalic = true;
            sheet.Range["i5:i5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["i5:i5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["i5:i5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["i5:i5"].Style.Font.Size = 10;
            sheet.Range["i5:i5"].Merge(); // birlashtirish
            sheet.Range["i5:i5"].Text = "Приход";
            sheet.Range["i5:i5"].Style.WrapText = true;
            //sheet.Range["i5:i5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["i5:i5"].BorderAround(LineStyleType.Thin);

            //sheet.Range["j5:j5"].Style.Font.IsBold = true;
            //sheet.Range["j5:j5"].Style.Font.IsItalic = true;
            sheet.Range["j5:j5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["j5:j5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["j5:j5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["j5:j5"].Style.Font.Size = 10;
            sheet.Range["j5:j5"].Merge(); // birlashtirish
            sheet.Range["j5:j5"].Text = "Расход";
            sheet.Range["j5:j5"].Style.WrapText = true;
            //sheet.Range["j5:j5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["j5:j5"].BorderAround(LineStyleType.Thin);


            //sheet.Range["k5:k5"].Style.Font.IsBold = true;
            //sheet.Range["k5:k5"].Style.Font.IsItalic = true;
            sheet.Range["k5:k5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["k5:k5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["k5:k5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["k5:k5"].Style.Font.Size = 10;
            sheet.Range["k5:k5"].Merge(); // birlashtirish
            sheet.Range["k5:k5"].Text = "Износ";
            sheet.Range["k5:k5"].Style.WrapText = true;
            //sheet.Range["k5:k5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["k5:k5"].BorderAround(LineStyleType.Thin);

            //sheet.Range["l5:l5"].Style.Font.IsBold = true;
            //sheet.Range["l5:l5"].Style.Font.IsItalic = true;
            sheet.Range["l5:l5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["l5:l5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["l5:l5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["l5:l5"].Style.Font.Size = 10;
            sheet.Range["l5:l5"].Merge(); // birlashtirish
            sheet.Range["l5:l5"].Text = "Салъдо изн";
            sheet.Range["l5:l5"].Style.WrapText = true;
            //sheet.Range["l5:l5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["l5:l5"].BorderAround(LineStyleType.Thin);


            int i = 0;
            int myrow = 6;
            int j = 0;
            int row_1 = 0;
            int r_count = 15;
            int my_row = 4 + r_count;

            while (row_1 < 2)
            {
                j = i;
                j = j + 1;
                row_1++;
                sheet.Range["a" + myrow + ":a" + myrow].Merge();
                sheet.Range["a" + myrow + ":a" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["a" + myrow + ":a" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.Size = 10;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["a" + myrow + ":a" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["a" + myrow + ":a" + myrow].Text = (String)j.ToString();

                sheet.Range["b" + myrow + ":b" + myrow].Merge();
                sheet.Range["b" + myrow + ":b" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["b" + myrow + ":b" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["b" + myrow + ":b" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.Size = 10;
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["b" + myrow + ":b" + myrow].Text = "";
                sheet.Range["b" + myrow + ":b" + myrow].Style.WrapText = true;


                sheet.Range["c" + myrow + ":c" + myrow].Merge();
                sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 10;
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["c" + myrow + ":c" + myrow].Text = "";
                sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;

                sheet.Range["d" + myrow + ":d" + myrow].Merge();
                sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
                sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["d" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 10;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["d" + myrow + ":d" + myrow].Text = "1000";

                sheet.Range["e" + myrow + ":e" + myrow].Merge();
                sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 10;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["e" + myrow + ":e" + myrow].Text = "100000";

                sheet.Range["f" + myrow + ":f" + myrow].Merge();
                sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":f" + myrow].Text = "1000000";

                sheet.Range["g" + myrow + ":g" + myrow].Merge();
                sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
                sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["g" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 10;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["g" + myrow + ":g" + myrow].Value = "";

                sheet.Range["h" + myrow + ":h" + myrow].Merge();
                sheet.Range["h" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["h" + myrow + ":h" + myrow].Style.WrapText = true;
                sheet.Range["h" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["h" + myrow + ":h" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.Size = 10;
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["h" + myrow + ":h" + myrow].Value = "";

                sheet.Range["i" + myrow + ":i" + myrow].Merge();
                sheet.Range["i" + myrow + ":i" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["i" + myrow + ":i" + myrow].Style.WrapText = true;
                sheet.Range["i" + myrow + ":i" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["i" + myrow + ":i" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["i" + myrow + ":i" + myrow].Style.Font.Size = 10;
                sheet.Range["i" + myrow + ":i" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["i" + myrow + ":i" + myrow].Value = "";

                sheet.Range["j" + myrow + ":j" + myrow].Merge();
                sheet.Range["j" + myrow + ":j" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["j" + myrow + ":j" + myrow].Style.WrapText = true;
                sheet.Range["j" + myrow + ":j" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["j" + myrow + ":j" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["j" + myrow + ":j" + myrow].Style.Font.Size = 10;
                sheet.Range["j" + myrow + ":j" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["j" + myrow + ":j" + myrow].Value = "";

                sheet.Range["k" + myrow + ":k" + myrow].Merge();
                sheet.Range["k" + myrow + ":k" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["k" + myrow + ":k" + myrow].Style.WrapText = true;
                sheet.Range["k" + myrow + ":k" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["k" + myrow + ":k" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["k" + myrow + ":k" + myrow].Style.Font.Size = 10;
                sheet.Range["k" + myrow + ":k" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["k" + myrow + ":k" + myrow].Value = "";

                sheet.Range["l" + myrow + ":l" + myrow].Merge();
                sheet.Range["l" + myrow + ":l" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["l" + myrow + ":l" + myrow].Style.WrapText = true;
                sheet.Range["l" + myrow + ":l" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["l" + myrow + ":l" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["l" + myrow + ":l" + myrow].Style.Font.Size = 10;
                sheet.Range["l" + myrow + ":l" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["l" + myrow + ":l" + myrow].Value = "";



                myrow = myrow + 1;
                i = i + 1;


            }

            myrow++;
            sheet.Range["d" + myrow + ":e" + myrow].Merge();
            sheet.Range["d" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["d" + myrow + ":e" + myrow].Style.WrapText = true;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.IsBold = true;
            sheet.Range["d" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.Size = 11;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["d" + myrow + ":e" + myrow].Text = "ИТОГО : ";
            //sheet.Range["d" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);

            sheet.Range["f" + myrow + ":g" + myrow].Merge();
            sheet.Range["f" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["f" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.Size = 11;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["f" + myrow + ":g" + myrow].Value = "2 066 913,00";
            sheet.Range["f" + myrow + ":g" + myrow].Style.WrapText = true;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.IsBold = true;
            sheet.Range["f" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);

            myrow++;
            myrow++;

            sheet.Range["b" + myrow + ":c" + myrow].Merge();
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Гл.бухгалтер:________________";
            sheet.SetRowHeight(myrow, 18);

            myrow++;

            sheet.Range["b" + myrow + ":c" + myrow].Merge();
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Бухгалтер: __________________";
            sheet.SetRowHeight(myrow, 18);


            sheet.Range["d5:" + myrow + "k"].NumberFormat = "#,##0.00";


            workbook.SaveToFile(Environment.CurrentDirectory + "\\docs\\Журнал-ордер.xlsx", Spire.Xls.FileFormat.Version2007);
            System.Diagnostics.Process.Start(workbook.FileName);



            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Журнал-ордер_excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //try
            //{
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.LeftMargin = 0.6;
            sheet.PageSetup.RightMargin = 0.5;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;


            sheet.PageSetup.Orientation = PageOrientationType.Portrait;


            sheet.Range["a1:a1"].ColumnWidth = 9;
            sheet.Range["b1:b1"].ColumnWidth = 9;
            sheet.Range["c1:c1"].ColumnWidth = 20;
            sheet.Range["d1:d1"].ColumnWidth = 2;
            sheet.Range["e1:e1"].ColumnWidth = 9;
            sheet.Range["f1:f1"].ColumnWidth = 9;
            sheet.Range["g1:g1"].ColumnWidth = 9;
            sheet.Range["h1:h1"].ColumnWidth = 20;



            sheet.Range["a1:g1"].Style.Font.IsBold = true;
            sheet.Range["a1:g1"].Style.Font.IsItalic = true;
            sheet.Range["a1:g1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a1:g1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a1:g1"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["a1:g1"].Style.Font.Size = 11;
            sheet.Range["a1:g1"].Merge(); // birlashtirish
            sheet.Range["a1:g1"].Text = "ГУВД1";
            sheet.Range["a1:g1"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["a1:g1"].Style.Font.Underline = FontUnderlineType.Single;
            sheet.SetRowHeight(1, 18);


            sheet.Range["a2:h2"].Style.Font.IsBold = true;
            //sheet.Range["a2:h2"].Style.Font.IsItalic = true;
            sheet.Range["a2:h2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a2:h2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a2:h2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a2:h2"].Style.Font.Size = 16;
            sheet.Range["a2:h2"].Merge(); // birlashtirish
            sheet.Range["a2:h2"].Text = "Журнал-ордер №7";
            sheet.Range["a2:h2"].Style.WrapText = true;
            //sheet.Range["a2:h2"].Style.Font.Color = Color.DarkBlue;
            sheet.SetRowHeight(2, 20);

            sheet.Range["a3:h3"].Style.Font.IsBold = true;
            //sheet.Range["a3:h3"].Style.Font.IsItalic = true;
            sheet.Range["a3:h3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a3:h3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a3:h3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a3:h3"].Style.Font.Size = 11;
            sheet.Range["a3:h3"].Merge(); // birlashtirish
            sheet.Range["a3:h3"].Text = "За Май 2021 год";
            sheet.Range["a3:h3"].Style.WrapText = true;
            sheet.SetRowHeight(3, 20);

            sheet.Range["a4:c4"].Style.Font.IsBold = true;
            //sheet.Range["a4:c4"].Style.Font.IsItalic = true;
            sheet.Range["a4:c4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a4:c4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a4:c4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a4:c4"].Style.Font.Size = 11;
            sheet.Range["a4:c4"].Merge(); // birlashtirish
            sheet.Range["a4:c4"].Text = "Подлежит записи в главную книгу";
            sheet.Range["a4:c4"].Style.WrapText = true;
            //sheet.Range["a4:c4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["a4:c4"].BorderAround(LineStyleType.Thin);

            sheet.Range["a5:a5"].Style.Font.IsBold = true;
            //sheet.Range["a5:a5"].Style.Font.IsItalic = true;
            sheet.Range["a5:a5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a5:a5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a5:a5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a5:a5"].Style.Font.Size = 11;
            sheet.Range["a5:a5"].Merge(); // birlashtirish
            sheet.Range["a5:a5"].Text = "Дебет";
            sheet.Range["a5:a5"].Style.WrapText = true;
            //sheet.Range["a5:a5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["a5:a5"].BorderAround(LineStyleType.Thin);

            sheet.Range["b5:b5"].Style.Font.IsBold = true;
            //sheet.Range["b5:b5"].Style.Font.IsItalic = true;
            sheet.Range["b5:b5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b5:b5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b5:b5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b5:b5"].Style.Font.Size = 11;
            sheet.Range["b5:b5"].Merge(); // birlashtirish
            sheet.Range["b5:b5"].Text = "Кредит";
            sheet.Range["b5:b5"].Style.WrapText = true;
            //sheet.Range["b5:b5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["b5:b5"].BorderAround(LineStyleType.Thin);

            sheet.Range["c5:c5"].Style.Font.IsBold = true;
            //sheet.Range["c5:c5"].Style.Font.IsItalic = true;
            sheet.Range["c5:c5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c5:c5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c5:c5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c5:c5"].Style.Font.Size = 11;
            sheet.Range["c5:c5"].Merge(); // birlashtirish
            sheet.Range["c5:c5"].Text = "Сумма";
            sheet.Range["c5:c5"].Style.WrapText = true;
            //sheet.Range["c5:c5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["c5:c5"].BorderAround(LineStyleType.Thin);

            int i = 0;
            int myrow = 6;
            int j = 0;
            int row_1 = 0;
            int r_count = 15;
            int my_row = 4 + r_count;

            while (row_1 < 2)
            {
                j = i;
                j = j + 1;
                row_1++;
                sheet.Range["a" + myrow + ":a" + myrow].Merge();
                sheet.Range["a" + myrow + ":a" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["a" + myrow + ":a" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.Size = 11;
                sheet.Range["a" + myrow + ":a" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["a" + myrow + ":a" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["a" + myrow + ":a" + myrow].Text = (String)j.ToString();

                sheet.Range["b" + myrow + ":b" + myrow].Merge();
                sheet.Range["b" + myrow + ":b" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["b" + myrow + ":b" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["b" + myrow + ":b" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.Size = 11;
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["b" + myrow + ":b" + myrow].Text = "";
                sheet.Range["b" + myrow + ":b" + myrow].Style.WrapText = true;


                sheet.Range["c" + myrow + ":c" + myrow].Merge();
                sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 11;
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["c" + myrow + ":c" + myrow].Text = "";
                sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;



                myrow = myrow + 1;
                i = i + 1;


            }



            sheet.Range["d5:d5"].Merge();

            sheet.Range["e4:h4"].Style.Font.IsBold = true;
            sheet.Range["e4:h4"].Style.Font.IsItalic = true;
            sheet.Range["e4:h4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e4:h4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e4:h4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e4:h4"].Style.Font.Size = 11;
            sheet.Range["e4:h4"].Merge(); // birlashtirish
            sheet.Range["e4:h4"].Text = "Расшифировка дебета 231";
            sheet.Range["e4:h4"].Style.WrapText = true;
            sheet.Range["e4:h4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["e4:h4"].BorderAround(LineStyleType.Thin);

            sheet.Range["e5:e5"].Style.Font.IsBold = true;
            sheet.Range["e5:e5"].Style.Font.IsItalic = true;
            sheet.Range["e5:e5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e5:e5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e5:e5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e5:e5"].Style.Font.Size = 11;
            sheet.Range["e5:e5"].Merge(); // birlashtirish
            sheet.Range["e5:e5"].Text = "Тип расхода";
            sheet.Range["e5:e5"].Style.WrapText = true;
            sheet.Range["e5:e5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["e5:e5"].BorderAround(LineStyleType.Thin);

            sheet.Range["f5:f5"].Style.Font.IsBold = true;
            sheet.Range["f5:f5"].Style.Font.IsItalic = true;
            sheet.Range["f5:f5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f5:f5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f5:f5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f5:f5"].Style.Font.Size = 11;
            sheet.Range["f5:f5"].Merge(); // birlashtirish
            sheet.Range["f5:f5"].Text = "Объект";
            sheet.Range["f5:f5"].Style.WrapText = true;
            sheet.Range["f5:f5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["f5:f5"].BorderAround(LineStyleType.Thin);

            sheet.Range["g5:g5"].Style.Font.IsBold = true;
            sheet.Range["g5:g5"].Style.Font.IsItalic = true;
            sheet.Range["g5:g5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g5:g5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g5:g5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g5:g5"].Style.Font.Size = 11;
            sheet.Range["g5:g5"].Merge(); // birlashtirish
            sheet.Range["g5:g5"].Text = "Под";
            sheet.Range["g5:g5"].Style.WrapText = true;
            sheet.Range["g5:g5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["g5:g5"].BorderAround(LineStyleType.Thin);

            sheet.Range["h5:h5"].Style.Font.IsBold = true;
            sheet.Range["h5:h5"].Style.Font.IsItalic = true;
            sheet.Range["h5:h5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h5:h5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h5:h5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["h5:h5"].Style.Font.Size = 11;
            sheet.Range["h5:h5"].Merge(); // birlashtirish
            sheet.Range["h5:h5"].Text = "Сумма";
            sheet.Range["h5:h5"].Style.WrapText = true;
            sheet.Range["h5:h5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["h5:h5"].BorderAround(LineStyleType.Thin);

            //////////////////
            ///under form

            sheet.Range["e" + myrow + ":h" + myrow].Style.Font.IsBold = true;
            sheet.Range["e" + myrow + ":h" + myrow].Style.Font.IsItalic = true;
            sheet.Range["e" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["e" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e" + myrow + ":h" + myrow].Style.Font.Size = 11;
            sheet.Range["e" + myrow + ":h" + myrow].Merge(); // birlashtirish
            sheet.Range["e" + myrow + ":h" + myrow].Text = "Расшифировка дебета износ 231";
            sheet.Range["e" + myrow + ":h" + myrow].Style.WrapText = true;
            sheet.Range["e" + myrow + ":h" + myrow].Style.Font.Color = Color.DarkBlue;
            sheet.Range["e" + myrow + ":h" + myrow].BorderAround(LineStyleType.Thin);

            myrow++;

            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.IsBold = true;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.IsItalic = true;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 11;
            sheet.Range["e" + myrow + ":e" + myrow].Merge(); // birlashtirish
            sheet.Range["e" + myrow + ":e" + myrow].Text = "Тип расхода";
            sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
            sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Color = Color.DarkBlue;
            sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);

            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.IsBold = true;
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.IsItalic = true;
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 11;
            sheet.Range["f" + myrow + ":f" + myrow].Merge(); // birlashtirish
            sheet.Range["f" + myrow + ":f" + myrow].Text = "Объект";
            sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
            sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Color = Color.DarkBlue;
            sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);

            sheet.Range["g" + myrow + ":g" + myrow].Style.Font.IsBold = true;
            sheet.Range["g" + myrow + ":g" + myrow].Style.Font.IsItalic = true;
            sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 11;
            sheet.Range["g" + myrow + ":g" + myrow].Merge(); // birlashtirish
            sheet.Range["g" + myrow + ":g" + myrow].Text = "Под";
            sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
            sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Color = Color.DarkBlue;
            sheet.Range["g" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);

            sheet.Range["h" + myrow + ":h" + myrow].Style.Font.IsBold = true;
            sheet.Range["h" + myrow + ":h" + myrow].Style.Font.IsItalic = true;
            sheet.Range["h" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["h" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["h" + myrow + ":h" + myrow].Style.Font.Size = 11;
            sheet.Range["h" + myrow + ":h" + myrow].Merge(); // birlashtirish
            sheet.Range["h" + myrow + ":h" + myrow].Text = "Сумма";
            sheet.Range["h" + myrow + ":h" + myrow].Style.WrapText = true;
            sheet.Range["h" + myrow + ":h" + myrow].Style.Font.Color = Color.DarkBlue;
            sheet.Range["h" + myrow + ":h" + myrow].BorderAround(LineStyleType.Thin);


            myrow++;
            sheet.Range["d" + myrow + ":e" + myrow].Merge();
            sheet.Range["d" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["d" + myrow + ":e" + myrow].Style.WrapText = true;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.IsBold = true;
            sheet.Range["d" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.Size = 11;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["d" + myrow + ":e" + myrow].Text = "ИТОГО : ";
            //sheet.Range["d" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);

            sheet.Range["f" + myrow + ":g" + myrow].Merge();
            sheet.Range["f" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["f" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.Size = 11;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["f" + myrow + ":g" + myrow].Value = "2 066 913,00";
            sheet.Range["f" + myrow + ":g" + myrow].Style.WrapText = true;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.IsBold = true;
            sheet.Range["f" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);

            myrow++;
            myrow++;

            sheet.Range["b" + myrow + ":c" + myrow].Merge();
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Гл.бухгалтер:________________";
            sheet.SetRowHeight(myrow, 18);

            myrow++;

            sheet.Range["b" + myrow + ":c" + myrow].Merge();
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Бухгалтер: __________________";
            sheet.SetRowHeight(myrow, 18);


            sheet.Range["d5:" + myrow + "k"].NumberFormat = "#,##0.00";


            workbook.SaveToFile(Environment.CurrentDirectory + "\\docs\\Журнал-ордер.xlsx", Spire.Xls.FileFormat.Version2007);
            System.Diagnostics.Process.Start(workbook.FileName);



            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Журнал-ордер_excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //try
            //{
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.LeftMargin = 0.3;
            sheet.PageSetup.RightMargin = 0.3;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;


            sheet.PageSetup.Orientation = PageOrientationType.Portrait;


            sheet.Range["a1:a1"].ColumnWidth = 3;
            sheet.Range["b1:b1"].ColumnWidth = 21;
            sheet.Range["c1:c1"].ColumnWidth = 21;
            sheet.Range["d1:d1"].ColumnWidth = 13;
            sheet.Range["e1:e1"].ColumnWidth = 13;
            sheet.Range["f1:f1"].ColumnWidth = 13;
            sheet.Range["g1:g1"].ColumnWidth = 13;


            sheet.Range["b1:e1"].Style.Font.IsBold = true;
            sheet.Range["b1:e1"].Style.Font.IsItalic = true;
            sheet.Range["b1:e1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b1:e1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b1:e1"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b1:e1"].Style.Font.Size = 11;
            sheet.Range["b1:e1"].Merge(); // birlashtirish
            sheet.Range["b1:e1"].Text = "СВОДНАЯ ОБОРОТЪ ЗА Декабръ 2020 год";
            sheet.Range["b1:e1"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["b1:e1"].Style.Font.Underline = FontUnderlineType.Single;

            sheet.Range["f1:g1"].Style.Font.IsBold = true;
            sheet.Range["f1:g1"].Style.Font.IsItalic = true;
            sheet.Range["f1:g1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f1:g1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f1:g1"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["f1:g1"].Style.Font.Size = 11;
            sheet.Range["f1:g1"].Merge(); // birlashtirish
            sheet.Range["f1:g1"].Text = "ГУВД1";
            sheet.Range["f1:g1"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["f1:g1"].Style.Font.Underline = FontUnderlineType.Single;
            sheet.SetRowHeight(1, 18);


            sheet.Range["b2:b2"].Style.Font.IsBold = true;
            sheet.Range["b2:b2"].Style.Font.IsItalic = true;
            sheet.Range["b2:b2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b2:b2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b2:b2"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b2:b2"].Style.Font.Size = 12;
            sheet.Range["b2:b2"].Merge(); // birlashtirish
            sheet.Range["b2:b2"].Text = " Талмут 01";
            sheet.Range["b2:b2"].Style.WrapText = true;
            sheet.Range["b2:b2"].Style.Font.Color = Color.DarkBlue;
            sheet.SetRowHeight(2, 20);

            sheet.Range["a3:b3"].Style.Font.IsBold = true;
            //sheet.Range["a:c3"].Style.Font.IsItalic = true;
            sheet.Range["a3:b3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a3:b3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a3:b3"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a3:b3"].Style.Font.Size = 10;
            sheet.Range["a3:b3"].Merge(); // birlashtirish
            sheet.Range["a3:b3"].Text = "   Счет 010";
            sheet.Range["a3:b3"].Style.WrapText = true;
            sheet.SetRowHeight(3, 18);

            // sheet.Range["a4:a5"].Style.Font.IsBold = true;
            //sheet.Range["a4:a5"].Style.Font.IsItalic = true;
            sheet.Range["a4:a5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a4:a5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a4:a5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a4:a5"].Style.Font.Size = 11;
            sheet.Range["a4:a5"].Merge(); // birlashtirish
            sheet.Range["a4:a5"].Text = "№ п.п";
            sheet.Range["a4:a5"].Style.WrapText = true;
            //heet.Range["a4:a5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["a4:a5"].BorderAround(LineStyleType.Thin);

            //sheet.Range["b4:b5"].Style.Font.IsBold = true;
            //sheet.Range["b4:b5"].Style.Font.IsItalic = true;
            sheet.Range["b4:b5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b4:b5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b4:b5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b4:b5"].Style.Font.Size = 11;
            sheet.Range["b4:b5"].Merge(); // birlashtirish
            sheet.Range["b4:b5"].Text = "ФИО";
            sheet.Range["b4:b5"].Style.WrapText = true;
            //sheet.Range["b4:b5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["b4:b5"].BorderAround(LineStyleType.Thin);

            //sheet.Range["c4:c5"].Style.Font.IsBold = true;
            // sheet.Range["c4:c5"].Style.Font.IsItalic = true;
            sheet.Range["c4:c5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c4:c5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c4:c5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c4:c5"].Style.Font.Size = 11;
            sheet.Range["c4:c5"].Merge(); // birlashtirish
            sheet.Range["c4:c5"].Text = "Подраздел.";
            sheet.Range["c4:c5"].Style.WrapText = true;
            // sheet.Range["c4:c5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["c4:c5"].BorderAround(LineStyleType.Thin);

            //sheet.Range["d4:g4"].Style.Font.IsBold = true;
            //sheet.Range["d4:g4"].Style.Font.IsItalic = true;
            sheet.Range["d4:g4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d4:g4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d4:g4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d4:g4"].Style.Font.Size = 11;
            sheet.Range["d4:g4"].Merge(); // birlashtirish
            sheet.Range["d4:g4"].Text = "Материалный оборот";
            sheet.Range["d4:g4"].Style.WrapText = true;
            sheet.Range["d4:g4"].BorderAround(LineStyleType.Thin);


            //sheet.Range["d5:d5"].Style.Font.IsBold = true;
            //sheet.Range["d5:d5"].Style.Font.IsItalic = true;
            sheet.Range["d5:d5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d5:d5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d5:d5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d5:d5"].Style.Font.Size = 11;
            sheet.Range["d5:d5"].Merge(); // birlashtirish
            sheet.Range["d5:d5"].Text = "Нач.Салъдо";
            sheet.Range["d5:d5"].Style.WrapText = true;
            //        sheet.Range["d5:d5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["d5:d5"].BorderAround(LineStyleType.Thin);

            //sheet.Range["e5:e5"].Style.Font.IsBold = true;
            //sheet.Range["e5:e5"].Style.Font.IsItalic = true;
            sheet.Range["e5:e5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e5:e5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e5:e5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e5:e5"].Style.Font.Size = 11;
            sheet.Range["e5:e5"].Merge(); // birlashtirish
            sheet.Range["e5:e5"].Text = "Приход";
            sheet.Range["e5:e5"].Style.WrapText = true;
            //sheet.Range["e5:e5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["e5:e5"].BorderAround(LineStyleType.Thin);

            //sheet.Range["f5:f5"].Style.Font.IsBold = true;
            //sheet.Range["f5:f5"].Style.Font.IsItalic = true;
            sheet.Range["f5:f5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f5:f5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f5:f5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f5:f5"].Style.Font.Size = 11;
            sheet.Range["f5:f5"].Merge(); // birlashtirish
            sheet.Range["f5:f5"].Text = "Расход";
            sheet.Range["f5:f5"].Style.WrapText = true;
            //sheet.Range["f5:f5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["f5:f5"].BorderAround(LineStyleType.Thin);

            //sheet.Range["g5:g5"].Style.Font.IsBold = true;
            //sheet.Range["g5:g5"].Style.Font.IsItalic = true;
            sheet.Range["g5:g5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g5:g5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g5:g5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g5:g5"].Style.Font.Size = 11;
            sheet.Range["g5:g5"].Merge(); // birlashtirish
            sheet.Range["g5:g5"].Text = "Остаток";
            sheet.Range["g5:g5"].Style.WrapText = true;
            //sheet.Range["g5:g5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["g5:g5"].BorderAround(LineStyleType.Thin);
            //sheet.SetRowHeight(4, 18);




            int i = 0;
            int myrow = 6;
            int j = 0;
            int row_1 = 0;
            int r_count = 15;
            int my_row = 4 + r_count;

            while (row_1 < 2)
            {
                j = i;
                j = j + 1;
                row_1++;
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
                sheet.Range["b" + myrow + ":b" + myrow].Text = "";
                sheet.Range["b" + myrow + ":b" + myrow].Style.WrapText = true;


                sheet.Range["c" + myrow + ":c" + myrow].Merge();
                sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 10;
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["c" + myrow + ":c" + myrow].Text = "";
                sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;

                sheet.Range["d" + myrow + ":d" + myrow].Merge();
                sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
                sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["d" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 10;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["d" + myrow + ":d" + myrow].Text = "1000";

                sheet.Range["e" + myrow + ":e" + myrow].Merge();
                sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 10;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["e" + myrow + ":e" + myrow].Text = "100000";

                sheet.Range["f" + myrow + ":f" + myrow].Merge();
                sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":f" + myrow].Text = "1000000";

                sheet.Range["g" + myrow + ":g" + myrow].Merge();
                sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
                sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["g" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 10;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["g" + myrow + ":g" + myrow].Value = "";


                myrow = myrow + 1;
                i = i + 1;


            }

            myrow++;
            sheet.Range["d" + myrow + ":e" + myrow].Merge();
            sheet.Range["d" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["d" + myrow + ":e" + myrow].Style.WrapText = true;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.IsBold = true;
            sheet.Range["d" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.Size = 11;
            sheet.Range["d" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["d" + myrow + ":e" + myrow].Text = "ИТОГО : ";
            //sheet.Range["d" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);

            sheet.Range["f" + myrow + ":g" + myrow].Merge();
            sheet.Range["f" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["f" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.Size = 11;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["f" + myrow + ":g" + myrow].Value = "2 066 913,00";
            sheet.Range["f" + myrow + ":g" + myrow].Style.WrapText = true;
            sheet.Range["f" + myrow + ":g" + myrow].Style.Font.IsBold = true;
            sheet.Range["f" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);

            myrow++;
            myrow++;

            sheet.Range["b" + myrow + ":c" + myrow].Merge();
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Гл.бухгалтер:________________";
            sheet.SetRowHeight(myrow, 18);

            myrow++;

            sheet.Range["b" + myrow + ":c" + myrow].Merge();
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Бухгалтер: __________________";
            sheet.SetRowHeight(myrow, 18);


            sheet.Range["d5:" + myrow + "k"].NumberFormat = "#,##0.00";


            workbook.SaveToFile(Environment.CurrentDirectory + "\\docs\\Журнал-ордер.xlsx", Spire.Xls.FileFormat.Version2007);
            System.Diagnostics.Process.Start(workbook.FileName);



            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Журнал-ордер_excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //try
            //{
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.LeftMargin = 0.3;
            sheet.PageSetup.RightMargin = 0.3;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;


            sheet.PageSetup.Orientation = PageOrientationType.Portrait;


            sheet.Range["a1:a1"].ColumnWidth = 7;
            sheet.Range["b1:b1"].ColumnWidth = 18;
            sheet.Range["c1:c1"].ColumnWidth = 26.71;
            sheet.Range["d1:d1"].ColumnWidth = 7;
            sheet.Range["e1:e1"].ColumnWidth = 12.57;
            sheet.Range["f1:f1"].ColumnWidth = 6;
            sheet.Range["g1:g1"].ColumnWidth = 6;
            sheet.Range["h1:h1"].ColumnWidth = 13;


            sheet.Range["a1:h1"].Style.Font.IsBold = true;
            sheet.Range["a1:h1"].Style.Font.IsItalic = true;
            sheet.Range["a1:h1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a1:h1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a1:h1"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a1:h1"].Style.Font.Size = 11;
            sheet.Range["a1:h1"].Merge(); // birlashtirish
            sheet.Range["a1:h1"].Text = "Расшифировка дебета 231";
            sheet.Range["a1:h1"].Style.WrapText = true;
            sheet.Range["a1:h1"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["a1:h1"].BorderAround(LineStyleType.Thin);

            sheet.Range["a2:a2"].Style.Font.IsBold = true;
            sheet.Range["a2:a2"].Style.Font.IsItalic = true;
            sheet.Range["a2:a2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a2:a2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a2:a2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a2:a2"].Style.Font.Size = 11;
            sheet.Range["a2:a2"].Merge(); // birlashtirish
            sheet.Range["a2:a2"].Text = "№ док";
            sheet.Range["a2:a2"].Style.WrapText = true;
            sheet.Range["a2:a2"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["a2:a2"].BorderAround(LineStyleType.Thin);

            sheet.Range["b2:b2"].Style.Font.IsBold = true;
            sheet.Range["b2:b2"].Style.Font.IsItalic = true;
            sheet.Range["b2:b2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b2:b2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b2:b2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b2:b2"].Style.Font.Size = 11;
            sheet.Range["b2:b2"].Merge(); // birlashtirish
            sheet.Range["b2:b2"].Text = "Отпустил";
            sheet.Range["b2:b2"].Style.WrapText = true;
            sheet.Range["b2:b2"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["b2:b2"].BorderAround(LineStyleType.Thin);


            sheet.Range["c2:c2"].Style.Font.IsBold = true;
            sheet.Range["c2:c2"].Style.Font.IsItalic = true;
            sheet.Range["c2:c2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c2:c2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c2:c2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c2:c2"].Style.Font.Size = 11;
            sheet.Range["c2:c2"].Merge(); // birlashtirish
            sheet.Range["c2:c2"].Text = "Наименование";
            sheet.Range["c2:c2"].Style.WrapText = true;
            sheet.Range["c2:c2"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["c2:c2"].BorderAround(LineStyleType.Thin);

            sheet.Range["d2:d2"].Style.Font.IsBold = true;
            sheet.Range["d2:d2"].Style.Font.IsItalic = true;
            sheet.Range["d2:d2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d2:d2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d2:d2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d2:d2"].Style.Font.Size = 11;
            sheet.Range["d2:d2"].Merge(); // birlashtirish
            sheet.Range["d2:d2"].Text = "Кол.";
            sheet.Range["d2:d2"].Style.WrapText = true;
            sheet.Range["d2:d2"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["d2:d2"].BorderAround(LineStyleType.Thin);

            sheet.Range["e2:e2"].Style.Font.IsBold = true;
            sheet.Range["e2:e2"].Style.Font.IsItalic = true;
            sheet.Range["e2:e2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e2:e2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e2:e2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e2:e2"].Style.Font.Size = 11;
            sheet.Range["e2:e2"].Merge(); // birlashtirish
            sheet.Range["e2:e2"].Text = "Сумма";
            sheet.Range["e2:e2"].Style.WrapText = true;
            sheet.Range["e2:e2"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["e2:e2"].BorderAround(LineStyleType.Thin);

            sheet.Range["f2:f2"].Style.Font.IsBold = true;
            sheet.Range["f2:f2"].Style.Font.IsItalic = true;
            sheet.Range["f2:f2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f2:f2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f2:f2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f2:f2"].Style.Font.Size = 11;
            sheet.Range["f2:f2"].Merge(); // birlashtirish
            sheet.Range["f2:f2"].Text = "Деб.";
            sheet.Range["f2:f2"].Style.WrapText = true;
            sheet.Range["f2:f2"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["f2:f2"].BorderAround(LineStyleType.Thin);

            sheet.Range["g2:g2"].Style.Font.IsBold = true;
            sheet.Range["g2:g2"].Style.Font.IsItalic = true;
            sheet.Range["g2:g2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g2:g2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g2:g2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g2:g2"].Style.Font.Size = 11;
            sheet.Range["g2:g2"].Merge(); // birlashtirish
            sheet.Range["g2:g2"].Text = "Кред.";
            sheet.Range["g2:g2"].Style.WrapText = true;
            sheet.Range["g2:g2"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["g2:g2"].BorderAround(LineStyleType.Thin);

            sheet.Range["h2:h2"].Style.Font.IsBold = true;
            sheet.Range["h2:h2"].Style.Font.IsItalic = true;
            sheet.Range["h2:h2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h2:h2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h2:h2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["h2:h2"].Style.Font.Size = 11;
            sheet.Range["h2:h2"].Merge(); // birlashtirish
            sheet.Range["h2:h2"].Text = "Статъи";
            sheet.Range["h2:h2"].Style.WrapText = true;
            sheet.Range["h2:h2"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["h2:h2"].BorderAround(LineStyleType.Thin);

            int i = 0;
            int myrow = 3;
            int j = 0;
            int row_1 = 0;
            int r_count = 15;
            int my_row = 4 + r_count;

            while (row_1 < 2)
            {
                j = i;
                j = j + 1;
                row_1++;
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
                sheet.Range["b" + myrow + ":b" + myrow].Text = "";
                sheet.Range["b" + myrow + ":b" + myrow].Style.WrapText = true;


                sheet.Range["c" + myrow + ":c" + myrow].Merge();
                sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 10;
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["c" + myrow + ":c" + myrow].Text = "";
                sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;

                sheet.Range["d" + myrow + ":d" + myrow].Merge();
                sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
                sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["d" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 10;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["d" + myrow + ":d" + myrow].Text = "1000";

                sheet.Range["e" + myrow + ":e" + myrow].Merge();
                sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 10;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["e" + myrow + ":e" + myrow].Text = "100000";

                sheet.Range["f" + myrow + ":f" + myrow].Merge();
                sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":f" + myrow].Text = "1000000";

                sheet.Range["g" + myrow + ":g" + myrow].Merge();
                sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
                sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["g" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 10;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["g" + myrow + ":g" + myrow].Value = "";

                sheet.Range["h" + myrow + ":h" + myrow].Merge();
                sheet.Range["h" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["h" + myrow + ":h" + myrow].Style.WrapText = true;
                sheet.Range["h" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["h" + myrow + ":h" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.Size = 10;
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["h" + myrow + ":h" + myrow].Value = "";



                myrow = myrow + 1;
                i = i + 1;


            }




            sheet.Range["d5:" + myrow + "k"].NumberFormat = "#,##0.00";


            workbook.SaveToFile(Environment.CurrentDirectory + "\\docs\\Журнал-ордер.xlsx", Spire.Xls.FileFormat.Version2007);
            System.Diagnostics.Process.Start(workbook.FileName);



            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Журнал-ордер_excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void button8_Click(object sender, EventArgs e)
        {
            //try
            //{
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.LeftMargin = 0.3;
            sheet.PageSetup.RightMargin = 0.3;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;


            sheet.PageSetup.Orientation = PageOrientationType.Portrait;


            sheet.Range["a1:a1"].ColumnWidth = 7;
            sheet.Range["b1:b1"].ColumnWidth = 18;
            sheet.Range["c1:c1"].ColumnWidth = 26.71;
            sheet.Range["d1:d1"].ColumnWidth = 7;
            sheet.Range["e1:e1"].ColumnWidth = 12.57;
            sheet.Range["f1:f1"].ColumnWidth = 6;
            sheet.Range["g1:g1"].ColumnWidth = 6;
            sheet.Range["h1:h1"].ColumnWidth = 13;


            sheet.Range["a1:h1"].Style.Font.IsBold = true;
            sheet.Range["a1:h1"].Style.Font.IsItalic = true;
            sheet.Range["a1:h1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a1:h1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a1:h1"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a1:h1"].Style.Font.Size = 11;
            sheet.Range["a1:h1"].Merge(); // birlashtirish
            sheet.Range["a1:h1"].Text = "Расшифировка дебета 231";
            sheet.Range["a1:h1"].Style.WrapText = true;
            sheet.Range["a1:h1"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["a1:h1"].BorderAround(LineStyleType.Thin);

            sheet.Range["a2:a2"].Style.Font.IsBold = true;
            sheet.Range["a2:a2"].Style.Font.IsItalic = true;
            sheet.Range["a2:a2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a2:a2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a2:a2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a2:a2"].Style.Font.Size = 11;
            sheet.Range["a2:a2"].Merge(); // birlashtirish
            sheet.Range["a2:a2"].Text = "№ док";
            sheet.Range["a2:a2"].Style.WrapText = true;
            sheet.Range["a2:a2"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["a2:a2"].BorderAround(LineStyleType.Thin);

            sheet.Range["b2:b2"].Style.Font.IsBold = true;
            sheet.Range["b2:b2"].Style.Font.IsItalic = true;
            sheet.Range["b2:b2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b2:b2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b2:b2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b2:b2"].Style.Font.Size = 11;
            sheet.Range["b2:b2"].Merge(); // birlashtirish
            sheet.Range["b2:b2"].Text = "Отпустил";
            sheet.Range["b2:b2"].Style.WrapText = true;
            sheet.Range["b2:b2"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["b2:b2"].BorderAround(LineStyleType.Thin);


            sheet.Range["c2:c2"].Style.Font.IsBold = true;
            sheet.Range["c2:c2"].Style.Font.IsItalic = true;
            sheet.Range["c2:c2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c2:c2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c2:c2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c2:c2"].Style.Font.Size = 11;
            sheet.Range["c2:c2"].Merge(); // birlashtirish
            sheet.Range["c2:c2"].Text = "Наименование";
            sheet.Range["c2:c2"].Style.WrapText = true;
            sheet.Range["c2:c2"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["c2:c2"].BorderAround(LineStyleType.Thin);

            sheet.Range["d2:d2"].Style.Font.IsBold = true;
            sheet.Range["d2:d2"].Style.Font.IsItalic = true;
            sheet.Range["d2:d2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d2:d2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d2:d2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d2:d2"].Style.Font.Size = 11;
            sheet.Range["d2:d2"].Merge(); // birlashtirish
            sheet.Range["d2:d2"].Text = "Кол.";
            sheet.Range["d2:d2"].Style.WrapText = true;
            sheet.Range["d2:d2"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["d2:d2"].BorderAround(LineStyleType.Thin);

            sheet.Range["e2:e2"].Style.Font.IsBold = true;
            sheet.Range["e2:e2"].Style.Font.IsItalic = true;
            sheet.Range["e2:e2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e2:e2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e2:e2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e2:e2"].Style.Font.Size = 11;
            sheet.Range["e2:e2"].Merge(); // birlashtirish
            sheet.Range["e2:e2"].Text = "Сумма";
            sheet.Range["e2:e2"].Style.WrapText = true;
            sheet.Range["e2:e2"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["e2:e2"].BorderAround(LineStyleType.Thin);

            sheet.Range["f2:f2"].Style.Font.IsBold = true;
            sheet.Range["f2:f2"].Style.Font.IsItalic = true;
            sheet.Range["f2:f2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f2:f2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f2:f2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f2:f2"].Style.Font.Size = 11;
            sheet.Range["f2:f2"].Merge(); // birlashtirish
            sheet.Range["f2:f2"].Text = "Деб.";
            sheet.Range["f2:f2"].Style.WrapText = true;
            sheet.Range["f2:f2"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["f2:f2"].BorderAround(LineStyleType.Thin);

            sheet.Range["g2:g2"].Style.Font.IsBold = true;
            sheet.Range["g2:g2"].Style.Font.IsItalic = true;
            sheet.Range["g2:g2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g2:g2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g2:g2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g2:g2"].Style.Font.Size = 11;
            sheet.Range["g2:g2"].Merge(); // birlashtirish
            sheet.Range["g2:g2"].Text = "Кред.";
            sheet.Range["g2:g2"].Style.WrapText = true;
            sheet.Range["g2:g2"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["g2:g2"].BorderAround(LineStyleType.Thin);

            sheet.Range["h2:h2"].Style.Font.IsBold = true;
            sheet.Range["h2:h2"].Style.Font.IsItalic = true;
            sheet.Range["h2:h2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h2:h2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h2:h2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["h2:h2"].Style.Font.Size = 11;
            sheet.Range["h2:h2"].Merge(); // birlashtirish
            sheet.Range["h2:h2"].Text = "Статъи";
            sheet.Range["h2:h2"].Style.WrapText = true;
            sheet.Range["h2:h2"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["h2:h2"].BorderAround(LineStyleType.Thin);

            int i = 0;
            int myrow = 3;
            int j = 0;
            int row_1 = 0;
            int r_count = 15;
            int my_row = 4 + r_count;

            while (row_1 < 2)
            {
                j = i;
                j = j + 1;
                row_1++;
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
                sheet.Range["b" + myrow + ":b" + myrow].Text = "";
                sheet.Range["b" + myrow + ":b" + myrow].Style.WrapText = true;


                sheet.Range["c" + myrow + ":c" + myrow].Merge();
                sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 10;
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["c" + myrow + ":c" + myrow].Text = "";
                sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;

                sheet.Range["d" + myrow + ":d" + myrow].Merge();
                sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
                sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["d" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 10;
                sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["d" + myrow + ":d" + myrow].Text = "1000";

                sheet.Range["e" + myrow + ":e" + myrow].Merge();
                sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
                sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 10;
                sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["e" + myrow + ":e" + myrow].Text = "100000";

                sheet.Range["f" + myrow + ":f" + myrow].Merge();
                sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
                sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 10;
                sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["f" + myrow + ":f" + myrow].Text = "1000000";

                sheet.Range["g" + myrow + ":g" + myrow].Merge();
                sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
                sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["g" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 10;
                sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["g" + myrow + ":g" + myrow].Value = "";

                sheet.Range["h" + myrow + ":h" + myrow].Merge();
                sheet.Range["h" + myrow + ":h" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["h" + myrow + ":h" + myrow].Style.WrapText = true;
                sheet.Range["h" + myrow + ":h" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["h" + myrow + ":h" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.Size = 10;
                sheet.Range["h" + myrow + ":h" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["h" + myrow + ":h" + myrow].Value = "";



                myrow = myrow + 1;
                i = i + 1;


            }




            sheet.Range["d5:" + myrow + "k"].NumberFormat = "#,##0.00";


            workbook.SaveToFile(Environment.CurrentDirectory + "\\docs\\Журнал-ордер.xlsx", Spire.Xls.FileFormat.Version2007);
            System.Diagnostics.Process.Start(workbook.FileName);



            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Журнал-ордер_excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void button9_Click(object sender, EventArgs e)
        {
            //try
            //{
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.LeftMargin = 0.6;
            sheet.PageSetup.RightMargin = 0.5;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;


            sheet.PageSetup.Orientation = PageOrientationType.Portrait;


            sheet.Range["a1:a1"].ColumnWidth = 5;
            sheet.Range["b1:b1"].ColumnWidth = 10;
            sheet.Range["c1:c1"].ColumnWidth = 21;
            sheet.Range["d1:d1"].ColumnWidth = 5;
            sheet.Range["e1:e1"].ColumnWidth = 5;
            sheet.Range["f1:f1"].ColumnWidth = 5;
            sheet.Range["g1:g1"].ColumnWidth = 10;
            sheet.Range["h1:h1"].ColumnWidth = 21;
            sheet.Range["i1:i1"].ColumnWidth = 5;



            sheet.Range["h1:i1"].Style.Font.IsBold = true;
            sheet.Range["h1:i1"].Style.Font.IsItalic = true;
            sheet.Range["h1:i1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h1:i1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h1:i1"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["h1:i1"].Style.Font.Size = 12;
            sheet.Range["h1:i1"].Merge(); // birlashtirish
            sheet.Range["h1:i1"].Text = "ГУВД1";
            sheet.Range["h1:i1"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["h1:i1"].Style.Font.Underline = FontUnderlineType.Single;
            sheet.SetRowHeight(1, 18);


            sheet.Range["a2:i2"].Style.Font.IsBold = true;
            //sheet.Range["a2:h2"].Style.Font.IsItalic = true;
            sheet.Range["a2:i2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a2:i2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a2:i2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a2:i2"].Style.Font.Size = 16;
            sheet.Range["a2:i2"].Merge(); // birlashtirish
            sheet.Range["a2:i2"].Text = "Журнал-ордер №7";
            sheet.Range["a2:i2"].Style.WrapText = true;
            //sheet.Range["a2:h2"].Style.Font.Color = Color.DarkBlue;
            sheet.SetRowHeight(2, 20);

            sheet.Range["a3:i3"].Style.Font.IsBold = true;
            //sheet.Range["a3:h3"].Style.Font.IsItalic = true;
            sheet.Range["a3:i3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a3:i3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a3:i3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a3:i3"].Style.Font.Size = 10;
            sheet.Range["a3:i3"].Merge(); // birlashtirish
            sheet.Range["a3:i3"].Text = "За Май 2021 год";
            sheet.Range["a3:i3"].Style.WrapText = true;
            sheet.SetRowHeight(3, 20);

            sheet.Range["a4:d4"].Style.Font.IsBold = true;
            sheet.Range["a4:d4"].Style.Font.IsItalic = true;
            sheet.Range["a4:d4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a4:d4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a4:d4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a4:d4"].Style.Font.Size = 13;
            sheet.Range["a4:d4"].Merge(); // birlashtirish
            sheet.Range["a4:d4"].Text = "Шапка для Журнал-Ордера №7     (дебет)";
            sheet.Range["a4:d4"].Style.WrapText = true;
            sheet.Range["a4:d4"].Style.Font.Color = Color.DarkBlue;


            sheet.Range["b5:b5"].Style.Font.IsBold = true;
            sheet.Range["b5:b5"].Style.Font.IsItalic = true;
            sheet.Range["b5:b5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b5:b5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b5:b5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b5:b5"].Style.Font.Size = 11;
            sheet.Range["b5:b5"].Merge(); // birlashtirish
            sheet.Range["b5:b5"].Text = "Деб_счет";
            sheet.Range["b5:b5"].Style.WrapText = true;
            sheet.Range["b5:b5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["b5:b5"].BorderAround(LineStyleType.Thin);

            sheet.Range["c5:c5"].Style.Font.IsBold = true;
            sheet.Range["c5:c5"].Style.Font.IsItalic = true;
            sheet.Range["c5:c5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c5:c5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c5:c5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c5:c5"].Style.Font.Size = 11;
            sheet.Range["c5:c5"].Merge(); // birlashtirish
            sheet.Range["c5:c5"].Text = "Сумма";
            sheet.Range["c5:c5"].Style.WrapText = true;
            sheet.Range["c5:c5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["c5:c5"].BorderAround(LineStyleType.Thin);

            sheet.Range["f4:i4"].Style.Font.IsBold = true;
            sheet.Range["f4:i4"].Style.Font.IsItalic = true;
            sheet.Range["f4:i4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f4:i4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f4:i4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f4:i4"].Style.Font.Size = 13;
            sheet.Range["f4:i4"].Merge(); // birlashtirish
            sheet.Range["f4:i4"].Text = "Шапка для Журнал-Ордера №7 (кредит)";
            sheet.Range["f4:i4"].Style.WrapText = true;
            sheet.Range["f4:i4"].Style.Font.Color = Color.DarkBlue;

            sheet.SetRowHeight(4, 35);

            sheet.Range["g5:g5"].Style.Font.IsBold = true;
            sheet.Range["g5:g5"].Style.Font.IsItalic = true;
            sheet.Range["g5:g5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g5:g5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g5:g5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g5:g5"].Style.Font.Size = 11;
            sheet.Range["g5:g5"].Merge(); // birlashtirish
            sheet.Range["g5:g5"].Text = "Кре_счет";
            sheet.Range["g5:g5"].Style.WrapText = true;
            sheet.Range["g5:g5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["g5:g5"].BorderAround(LineStyleType.Thin);

            sheet.Range["h5:h5"].Style.Font.IsBold = true;
            sheet.Range["h5:h5"].Style.Font.IsItalic = true;
            sheet.Range["h5:h5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h5:h5"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h5:h5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["h5:h5"].Style.Font.Size = 11;
            sheet.Range["h5:h5"].Merge(); // birlashtirish
            sheet.Range["h5:h5"].Text = "Сумма";
            sheet.Range["h5:h5"].Style.WrapText = true;
            sheet.Range["h5:h5"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["h5:h5"].BorderAround(LineStyleType.Thin);

            int i = 0;
            int myrow = 6;
            int j = 0;
            int row_1 = 0;
            int r_count = 15;
            int my_row = 4 + r_count;

            while (row_1 < 2)
            {
                j = i;
                j = j + 1;
                row_1++;

                sheet.Range["b" + myrow + ":b" + myrow].Merge();
                sheet.Range["b" + myrow + ":b" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["b" + myrow + ":b" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["b" + myrow + ":b" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.Size = 11;
                sheet.Range["b" + myrow + ":b" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["b" + myrow + ":b" + myrow].Text = "";
                sheet.Range["b" + myrow + ":b" + myrow].Style.WrapText = true;


                sheet.Range["c" + myrow + ":c" + myrow].Merge();
                sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 11;
                sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["c" + myrow + ":c" + myrow].Text = "";
                sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;



                myrow = myrow + 1;
                i = i + 1;


            }
            sheet.Range["a4:" + myrow + "d"].BorderAround(LineStyleType.Thin);

            int i2 = 0;
            int myrow2 = 6;
            int j2 = 0;
            int row_12 = 0;
            int r_count2 = 15;
            int my_row2 = 4 + r_count;

            while (row_12 < 2)
            {
                j2 = i2;
                j2 = j2 + 1;
                row_12++;

                sheet.Range["b" + myrow2 + ":b" + myrow2].Merge();
                sheet.Range["b" + myrow2 + ":b" + myrow2].Style.HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range["b" + myrow2 + ":b" + myrow2].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["b" + myrow2 + ":b" + myrow2].BorderAround(LineStyleType.Thin);
                sheet.Range["b" + myrow2 + ":b" + myrow2].Style.Font.Size = 11;
                sheet.Range["b" + myrow2 + ":b" + myrow2].Style.Font.FontName = "Times New Roman";
                sheet.Range["b" + myrow2 + ":b" + myrow2].Text = "";
                sheet.Range["b" + myrow2 + ":b" + myrow2].Style.WrapText = true;


                sheet.Range["c" + myrow2 + ":c" + myrow2].Merge();
                sheet.Range["c" + myrow2 + ":c" + myrow2].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["c" + myrow2 + ":c" + myrow2].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["c" + myrow2 + ":c" + myrow2].BorderAround(LineStyleType.Thin);
                sheet.Range["c" + myrow2 + ":c" + myrow2].Style.Font.Size = 11;
                sheet.Range["c" + myrow2 + ":c" + myrow2].Style.Font.FontName = "Times New Roman";
                sheet.Range["c" + myrow2 + ":c" + myrow2].Text = "";
                sheet.Range["c" + myrow2 + ":c" + myrow2].Style.WrapText = true;



                myrow2 = myrow2 + 1;
                i2 = i2 + 1;


            }
            sheet.Range["f4:" + myrow2 + "i"].BorderAround(LineStyleType.Thin);


            myrow++;
            myrow++;

            sheet.Range["b" + myrow + ":c" + myrow].Merge();
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Гл.бухгалтер:_____________";
            sheet.SetRowHeight(myrow, 18);

            myrow++;

            sheet.Range["b" + myrow + ":c" + myrow].Merge();
            sheet.Range["b" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["b" + myrow + ":c" + myrow].Style.WrapText = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.IsBold = true;
            sheet.Range["b" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.Size = 11;
            sheet.Range["b" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            sheet.Range["b" + myrow + ":c" + myrow].Text = "Бухгалтер: _______________";
            sheet.SetRowHeight(myrow, 18);


            sheet.Range["d5:" + myrow + "k"].NumberFormat = "#,##0.00";


            workbook.SaveToFile(Environment.CurrentDirectory + "\\docs\\Журнал-ордер.xlsx", Spire.Xls.FileFormat.Version2007);
            System.Diagnostics.Process.Start(workbook.FileName);



            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Журнал-ордер_excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void button17_Click(object sender, EventArgs e)
        {
            //try
            //{
            // month_global = month_textBox.Text;
            // year_global = year_textBox.Text;

            pereotsenka pereotsenka = new pereotsenka(string_for_otdels,year_global,month_global);
            pereotsenka.WindowState = FormWindowState.Maximized;
            pereotsenka.Show();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("prixod_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }
    }
}
