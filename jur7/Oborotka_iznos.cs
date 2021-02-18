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
    public partial class Oborotka_iznos : Form
    {
        Connect sql = new Connect();
        Connect sql1 = new Connect();
        Connect sql2 = new Connect();

        public string string_for_otdels;
        public string month_global;
        public string year_global;
        public Oborotka_iznos(string string_for_otdels, string year_global, string month_global)
        {
            InitializeComponent();

            this.string_for_otdels = string_for_otdels;
            this.year_global = year_global;
            this.month_global = month_global;
        }

        private void oborotka_iznos_btn_Click(object sender, EventArgs e)
        {
            //try
            //{
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.LeftMargin = 0.4;
            sheet.PageSetup.RightMargin = 0.4;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;


            sheet.PageSetup.Orientation = PageOrientationType.Landscape;


            sheet.Range["a1:a1"].ColumnWidth = 7;
            sheet.Range["b1:b1"].ColumnWidth = 35;
            sheet.Range["c1:c1"].ColumnWidth = 13;
            sheet.Range["d1:d1"].ColumnWidth = 8;
            sheet.Range["e1:e1"].ColumnWidth = 6.86;
            sheet.Range["f1:f1"].ColumnWidth = 14;
            sheet.Range["g1:g1"].ColumnWidth = 13;
            sheet.Range["h1:h1"].ColumnWidth = 13;
            sheet.Range["i1:i1"].ColumnWidth = 14;
            sheet.Range["j1:j1"].ColumnWidth = 14;


            sheet.Range["a1:j1"].Style.Font.IsBold = true;
            //sheet.Range["a1:g1"].Style.Font.IsItalic = true;
            sheet.Range["a1:j1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a1:j1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a1:j1"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a1:j1"].Style.Font.Size = 14;
            sheet.Range["a1:j1"].Merge(); // birlashtirish
            sheet.Range["a1:j1"].Text = "Оборотка по износу ";
            sheet.Range["a1:j1"].Style.WrapText = true;
            sheet.Range["a1:j1"].Style.Font.Color = Color.DarkBlue;
            sheet.SetRowHeight(1, 22);
            sheet.SetRowHeight(2, 4);

            sheet.Range["a3:a3"].Style.Font.IsBold = true;
            sheet.Range["a3:a3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a3:a3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a3:a3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a3:a3"].Style.Font.Size = 11;
            sheet.Range["a3:a3"].Merge(); // birlashtirish
            sheet.Range["a3:a3"].Text = "Код";
            sheet.Range["a3:a3"].Style.WrapText = true;
            sheet.Range["a3:a3"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["a3:a3"].BorderAround(LineStyleType.Thin);

            sheet.Range["b3:b3"].Style.Font.IsBold = true;
            sheet.Range["b3:b3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b3:b3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b3:b3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b3:b3"].Style.Font.Size = 11;
            sheet.Range["b3:b3"].Merge(); // birlashtirish
            sheet.Range["b3:b3"].Text = "Наим_тов";
            sheet.Range["b3:b3"].Style.WrapText = true;
            sheet.Range["b3:b3"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["b3:b3"].BorderAround(LineStyleType.Thin);

            sheet.Range["c3:c3"].Style.Font.IsBold = true;
            sheet.Range["c3:c3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c3:c3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c3:c3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c3:c3"].Style.Font.Size = 11;
            sheet.Range["c3:c3"].Merge(); // birlashtirish
            sheet.Range["c3:c3"].Text = "Цена";
            sheet.Range["c3:c3"].Style.WrapText = true;
            sheet.Range["c3:c3"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["c3:c3"].BorderAround(LineStyleType.Thin);

            sheet.Range["d3:d3"].Style.Font.IsBold = true;
            sheet.Range["d3:d3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d3:d3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d3:d3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d3:d3"].Style.Font.Size = 11;
            sheet.Range["d3:d3"].Merge(); // birlashtirish
            sheet.Range["d3:d3"].Text = "Дата";
            sheet.Range["d3:d3"].Style.WrapText = true;
            sheet.Range["d3:d3"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["d3:d3"].BorderAround(LineStyleType.Thin);

            sheet.Range["e3:e3"].Style.Font.IsBold = true;
            sheet.Range["e3:e3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e3:e3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e3:e3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e3:e3"].Style.Font.Size = 11;
            sheet.Range["e3:e3"].Merge(); // birlashtirish
            sheet.Range["e3:e3"].Text = "%";
            sheet.Range["e3:e3"].Style.WrapText = true;
            sheet.Range["e3:e3"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["e3:e3"].BorderAround(LineStyleType.Thin);

            sheet.Range["f3:f3"].Style.Font.IsBold = true;
            sheet.Range["f3:f3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f3:f3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f3:f3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f3:f3"].Style.Font.Size = 11;
            sheet.Range["f3:f3"].Merge(); // birlashtirish
            sheet.Range["f3:f3"].Text = "Салъдо";
            sheet.Range["f3:f3"].Style.WrapText = true;
            sheet.Range["f3:f3"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["f3:f3"].BorderAround(LineStyleType.Thin);

            sheet.Range["g3:g3"].Style.Font.IsBold = true;
            sheet.Range["g3:g3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g3:g3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g3:g3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g3:g3"].Style.Font.Size = 11;
            sheet.Range["g3:g3"].Merge(); // birlashtirish
            sheet.Range["g3:g3"].Text = "Приход";
            sheet.Range["g3:g3"].Style.WrapText = true;
            sheet.Range["g3:g3"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["g3:g3"].BorderAround(LineStyleType.Thin);

            sheet.Range["h3:h3"].Style.Font.IsBold = true;
            sheet.Range["h3:h3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h3:h3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h3:h3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["h3:h3"].Style.Font.Size = 11;
            sheet.Range["h3:h3"].Merge(); // birlashtirish
            sheet.Range["h3:h3"].Text = "Расход";
            sheet.Range["h3:h3"].Style.WrapText = true;
            sheet.Range["h3:h3"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["h3:h3"].BorderAround(LineStyleType.Thin);

            sheet.Range["i3:i3"].Style.Font.IsBold = true;
            sheet.Range["i3:i3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["i3:i3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["i3:i3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["i3:i3"].Style.Font.Size = 11;
            sheet.Range["i3:i3"].Merge(); // birlashtirish
            sheet.Range["i3:i3"].Text = "Износ";
            sheet.Range["i3:i3"].Style.WrapText = true;
            sheet.Range["i3:i3"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["i3:i3"].BorderAround(LineStyleType.Thin);

            sheet.Range["j3:j3"].Style.Font.IsBold = true;
            sheet.Range["j3:j3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["j3:j3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["j3:j3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["j3:j3"].Style.Font.Size = 11;
            sheet.Range["j3:j3"].Merge(); // birlashtirish
            sheet.Range["j3:j3"].Text = "Салъдо";
            sheet.Range["j3:j3"].Style.WrapText = true;
            sheet.Range["j3:j3"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["j3:j3"].BorderAround(LineStyleType.Thin);

            sheet.Range["a3:j3"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.SetRowHeight(3,20);
            //int i = 0;
            //int myrow = 4;
            //int j = 0;

            
            //var exl = "SELECT distinct schet,subschet FROM gruppa ";

            //sql.myReader = sql.return_MySqlCommand(exl).ExecuteReader();
            //while (sql.myReader.Read())
            //{

            //    j = i;
            //    j = j + 1;

            //    sheet.Range["a" + myrow + ":a" + myrow].Merge();
            //    sheet.Range["a" + myrow + ":a" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            //    sheet.Range["a" + myrow + ":a" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            //    sheet.Range["a" + myrow + ":a" + myrow].Style.Font.Size = 11;
            //    sheet.Range["a" + myrow + ":a" + myrow].Style.Font.FontName = "Times New Roman";
            //    sheet.Range["a" + myrow + ":a" + myrow].BorderAround(LineStyleType.Thin);
            //    sheet.Range["a" + myrow + ":a" + myrow].Text = "";

            //    sheet.Range["b" + myrow + ":b" + myrow].Merge();
            //    sheet.Range["b" + myrow + ":b" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            //    sheet.Range["b" + myrow + ":b" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            //    sheet.Range["b" + myrow + ":b" + myrow].BorderAround(LineStyleType.Thin);
            //    sheet.Range["b" + myrow + ":b" + myrow].Style.Font.Size = 11;
            //    sheet.Range["b" + myrow + ":b" + myrow].Style.Font.FontName = "Times New Roman";
            //    sheet.Range["b" + myrow + ":b" + myrow].Text = "";
            //    sheet.Range["b" + myrow + ":b" + myrow].Style.WrapText = true;


            //    sheet.Range["c" + myrow + ":c" + myrow].Merge();
            //    sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            //    sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            //    sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
            //    sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 11;
            //    sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            //    sheet.Range["c" + myrow + ":c" + myrow].Value = "";
            //    sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;

            //    sheet.Range["d" + myrow + ":d" + myrow].Merge();
            //    sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            //    sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
            //    sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            //    sheet.Range["d" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);
            //    sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 11;
            //    sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
            //    sheet.Range["d" + myrow + ":d" + myrow].Value = "";

            //    sheet.Range["e" + myrow + ":e" + myrow].Merge();
            //    sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            //    sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
            //    sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            //    sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);
            //    sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 11;
            //    sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
            //    sheet.Range["e" + myrow + ":e" + myrow].Value = "";

            //    sheet.Range["f" + myrow + ":f" + myrow].Merge();
            //    sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            //    sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
            //    sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            //    sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
            //    sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 11;
            //    sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
            //    sheet.Range["f" + myrow + ":f" + myrow].Value = "";

            //    sheet.Range["g" + myrow + ":g" + myrow].Merge();
            //    sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            //    sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
            //    sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            //    sheet.Range["g" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);
            //    sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 11;
            //    sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
            //    sheet.Range["g" + myrow + ":g" + myrow].Value = "";


            //    myrow = myrow + 1;
            //    i = i + 1;

            //}
            //sql.myReader.Close();

            //sheet.Range["a" + myrow + ":b" + myrow].Merge();
            //sheet.Range["a" + myrow + ":b" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            //sheet.Range["a" + myrow + ":b" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            //sheet.Range["a" + myrow + ":b" + myrow].BorderAround(LineStyleType.Thin);
            //sheet.Range["a" + myrow + ":b" + myrow].Style.Font.Size = 11;
            //sheet.Range["a" + myrow + ":b" + myrow].Style.Font.FontName = "Times New Roman";
            //sheet.Range["a" + myrow + ":b" + myrow].Text = "Итого :";
            //sheet.Range["a" + myrow + ":b" + myrow].Style.WrapText = true;
            ////sheet.Range["a" + myrow + ":b" + myrow].Style.Font.IsBold = true;


            //sheet.Range["c" + myrow + ":c" + myrow].Merge();
            //sheet.Range["c" + myrow + ":c" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            //sheet.Range["c" + myrow + ":c" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            //sheet.Range["c" + myrow + ":c" + myrow].BorderAround(LineStyleType.Thin);
            //sheet.Range["c" + myrow + ":c" + myrow].Style.Font.Size = 11;
            //sheet.Range["c" + myrow + ":c" + myrow].Style.Font.FontName = "Times New Roman";
            //sheet.Range["c" + myrow + ":c" + myrow].Value = "";
            //sheet.Range["c" + myrow + ":c" + myrow].Style.WrapText = true;
            ////sheet.Range["c" + myrow + ":c" + myrow].Style.Font.IsBold = true;

            //sheet.Range["d" + myrow + ":d" + myrow].Merge();
            //sheet.Range["d" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            //sheet.Range["d" + myrow + ":d" + myrow].Style.WrapText = true;
            ////sheet.Range["d" + myrow + ":d" + myrow].Style.Font.IsBold = true;
            //sheet.Range["d" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            //sheet.Range["d" + myrow + ":d" + myrow].BorderAround(LineStyleType.Thin);
            //sheet.Range["d" + myrow + ":d" + myrow].Style.Font.Size = 11;
            //sheet.Range["d" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
            //sheet.Range["d" + myrow + ":d" + myrow].Value = "";

            //sheet.Range["e" + myrow + ":e" + myrow].Merge();
            //sheet.Range["e" + myrow + ":e" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            //sheet.Range["e" + myrow + ":e" + myrow].Style.WrapText = true;
            ////sheet.Range["e" + myrow + ":e" + myrow].Style.Font.IsBold = true;
            //sheet.Range["e" + myrow + ":e" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            //sheet.Range["e" + myrow + ":e" + myrow].BorderAround(LineStyleType.Thin);
            //sheet.Range["e" + myrow + ":e" + myrow].Style.Font.Size = 11;
            //sheet.Range["e" + myrow + ":e" + myrow].Style.Font.FontName = "Times New Roman";
            //sheet.Range["e" + myrow + ":e" + myrow].Value = "";

            //sheet.Range["f" + myrow + ":f" + myrow].Merge();
            //sheet.Range["f" + myrow + ":f" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            //sheet.Range["f" + myrow + ":f" + myrow].Style.WrapText = true;
            ////sheet.Range["f" + myrow + ":f" + myrow].Style.Font.IsBold = true;
            //sheet.Range["f" + myrow + ":f" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            //sheet.Range["f" + myrow + ":f" + myrow].BorderAround(LineStyleType.Thin);
            //sheet.Range["f" + myrow + ":f" + myrow].Style.Font.Size = 11;
            //sheet.Range["f" + myrow + ":f" + myrow].Style.Font.FontName = "Times New Roman";
            //sheet.Range["f" + myrow + ":f" + myrow].Value = "";

            //sheet.Range["g" + myrow + ":g" + myrow].Merge();
            //sheet.Range["g" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
            //sheet.Range["g" + myrow + ":g" + myrow].Style.WrapText = true;
            ////sheet.Range["g" + myrow + ":g" + myrow].Style.Font.IsBold = true;
            //sheet.Range["g" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            //sheet.Range["g" + myrow + ":g" + myrow].BorderAround(LineStyleType.Thin);
            //sheet.Range["g" + myrow + ":g" + myrow].Style.Font.Size = 11;
            //sheet.Range["g" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
            //sheet.Range["g" + myrow + ":g" + myrow].Value = "";


            //myrow++;
            //myrow++;

            //sheet.Range["b" + myrow + ":d" + myrow].Style.Font.IsBold = true;
            ////sheet.Range["b" + myrow + ":d" + myrow].Style.Font.IsItalic = true;
            //sheet.Range["b" + myrow + ":d" + myrow].Style.Font.FontName = "Times New Roman";
            //sheet.Range["b" + myrow + ":d" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            //sheet.Range["b" + myrow + ":d" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            //sheet.Range["b" + myrow + ":d" + myrow].Style.Font.Size = 14;
            //sheet.Range["b" + myrow + ":d" + myrow].Merge(); // birlashtirish
            //sheet.Range["b" + myrow + ":d" + myrow].Text = "Гл.бухгалтер __________________";
            //sheet.Range["b" + myrow + ":d" + myrow].Style.WrapText = true;
            //sheet.Range["b" + myrow + ":d" + myrow].Style.Font.Color = Color.DarkBlue;

            //sheet.Range["e" + myrow + ":g" + myrow].Style.Font.IsBold = true;
            ////sheet.Range["e" + myrow + ":g" + myrow].Style.Font.IsItalic = true;
            //sheet.Range["e" + myrow + ":g" + myrow].Style.Font.FontName = "Times New Roman";
            //sheet.Range["e" + myrow + ":g" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
            //sheet.Range["e" + myrow + ":g" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Center;
            //sheet.Range["e" + myrow + ":g" + myrow].Style.Font.Size = 14;
            //sheet.Range["e" + myrow + ":g" + myrow].Merge(); // birlashtirish
            //sheet.Range["e" + myrow + ":g" + myrow].Text = "Бухгалтер __________________";
            //sheet.Range["e" + myrow + ":g" + myrow].Style.WrapText = true;
            //sheet.Range["e" + myrow + ":g" + myrow].Style.Font.Color = Color.DarkBlue;

            //sheet.Range["b3:" + myrow + "g"].NumberFormat = "#,##0.00";


            workbook.SaveToFile(Environment.CurrentDirectory + "\\docs\\Износ по субсче.xlsx", Spire.Xls.FileFormat.Version2007);
            System.Diagnostics.Process.Start(workbook.FileName);



            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Журнал-ордер_excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }
    }
}
