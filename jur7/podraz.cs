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
    public partial class podraz : Form
    {
        Connect sql = new Connect();
        Connect sql1 = new Connect();

        public string string_for_otdels;
        public string month_global;
        public string year_global;

        public podraz(string string_for_otdels, string year_global, string month_global)
        {
            InitializeComponent();

            sql.Connection();
            sql1.Connection();

            this.string_for_otdels = string_for_otdels;
            this.year_global = year_global;
            this.month_global = month_global;

            run_main();
        }

        public void run_main()
        {
            try
            { 
            var query = " SELECT * FROM podraz_jur7 where kod_pod is not null order by cast(kod_pod as unsigned)  ";
            sql1.myReader = sql1.return_MySqlCommand(query).ExecuteReader();
            while (sql1.myReader.Read())
            {
                //gruppa,kod_gruppa,naim,schet,prosent_izn,debet,subschet,kredit
                podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Add()].Cells[0].Value = (sql1.myReader["id"] != DBNull.Value ? sql1.myReader.GetString("id") : "");
                podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[1].Value = (sql1.myReader["kod_pod"] != DBNull.Value ? sql1.myReader.GetString("kod_pod") : "");
                podrazdelenie_dataGridView.Rows[podrazdelenie_dataGridView.Rows.Count - 2].Cells[2].Value = (sql1.myReader["podraz_naim"] != DBNull.Value ? sql1.myReader.GetString("podraz_naim") : "");
            }
            sql1.myReader.Close();
        }
            catch (Exception ex)
            {
                MessageBox.Show("run_treeview " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
}

        private void podraz_Load(object sender, EventArgs e)
        {
            this.podrazdelenie_dataGridView.RowsDefaultCellStyle.BackColor = Color.White;
            this.podrazdelenie_dataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(233, 233, 234);

            podrazdelenie_dataGridView.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            podrazdelenie_dataGridView.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
        }

        private void podrazdelenie_dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (podrazdelenie_dataGridView.CurrentRow != null)
            {
                DataGridViewRow dgvRow = podrazdelenie_dataGridView.CurrentRow;




                if (dgvRow.Cells[0].Value == null)
                {
                    Console.WriteLine("insert");

                    sql.return_MySqlCommand("insert into podraz_jur7 (kod_pod,podraz_naim) values" +
                                        "('" + (podrazdelenie_dataGridView.CurrentRow.Cells[1].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[1].Value : "") + "', " +
                                        "'" + (podrazdelenie_dataGridView.CurrentRow.Cells[2].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[2].Value : "") + "' " +
                                        ") ").ExecuteNonQuery();

                    this.podrazdelenie_dataGridView.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.podrazdelenie_dataGridView_CellValueChanged);
                    sql1.myReader = sql1.return_MySqlCommand("select max(id) as id from podraz_jur7").ExecuteReader();
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

                    sql.return_MySqlCommand("update podraz_jur7 set " +
                     "kod_pod = '" + (podrazdelenie_dataGridView.CurrentRow.Cells[1].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[1].Value : "") + "', " +
                     "podraz_naim = '" + (podrazdelenie_dataGridView.CurrentRow.Cells[2].Value != null ? podrazdelenie_dataGridView.CurrentRow.Cells[2].Value : "") + "' " +
                     " where id = '" + podrazdelenie_dataGridView.CurrentRow.Cells[0].Value + "' ").ExecuteNonQuery();
                }
            }


            }
            catch (Exception ex)
            {
                MessageBox.Show("gruppa_dataGridView_CellValueChanged " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        string kod_pod = "";
        string podraz_naim = "";
        private void podrazdelenie_dataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = podrazdelenie_dataGridView.CurrentRow;

            kod_pod = row.Cells[1].Value.ToString();
            podraz_naim = row.Cells[2].Value.ToString();
            podraz_fio gruppa = new podraz_fio(kod_pod, podraz_naim,string_for_otdels,year_global,month_global);
            if (e.ColumnIndex == 2)
            {
                if (gruppa.ShowDialog() == DialogResult.OK)
                {

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

                            sql.return_MySqlCommand("delete from podraz_jur7 where id = " + row.Cells[0].Value + "").ExecuteNonQuery();
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

        private void button1_Click(object sender, EventArgs e)
        {
            //try
            //{
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.LeftMargin = 0.1;
            sheet.PageSetup.RightMargin = 0.1;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;


            sheet.PageSetup.Orientation = PageOrientationType.Portrait;


            sheet.Range["a1:a1"].ColumnWidth = 22.14;
            sheet.Range["b1:b1"].ColumnWidth = 5;
            sheet.Range["c1:c1"].ColumnWidth = 6;
            sheet.Range["d1:d1"].ColumnWidth = 10;
            sheet.Range["e1:e1"].ColumnWidth = 6;
            sheet.Range["f1:f1"].ColumnWidth = 10;
            sheet.Range["g1:g1"].ColumnWidth = 6;
            sheet.Range["h1:h1"].ColumnWidth = 10;
            sheet.Range["i1:i1"].ColumnWidth = 6;
            sheet.Range["j1:j1"].ColumnWidth = 10;
            sheet.Range["k1:k1"].ColumnWidth = 8.43;



            sheet.Range["f1:h1"].Style.Font.IsBold = true;
            sheet.Range["f1:h1"].Style.Font.IsItalic = true;
            sheet.Range["f1:h1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f1:h1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f1:h1"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["f1:h1"].Style.Font.Size = 11;
            sheet.Range["f1:h1"].Merge(); // birlashtirish
            sheet.Range["f1:h1"].Text = "ГУВД1";
            sheet.Range["f1:h1"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["f1:h1"].Style.Font.Underline = FontUnderlineType.Single;
            sheet.SetRowHeight(1, 18);


            sheet.Range["a2:g2"].Style.Font.IsBold = true;
            sheet.Range["a2:g2"].Style.Font.IsItalic = true;
            sheet.Range["a2:g2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a2:g2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a2:g2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a2:g2"].Style.Font.Size = 14;
            sheet.Range["a2:g2"].Merge(); // birlashtirish
            sheet.Range["a2:g2"].Text = "СВОДНАЯ ОБОРОТЪ ЗА Декабръ 2020 год";
            sheet.Range["a2:g2"].Style.WrapText = true;
            sheet.Range["a2:g2"].Style.Font.Color = Color.DarkBlue;
            sheet.SetRowHeight(2, 20);


            //sheet.Range["a3:a4"].Style.Font.IsBold = true;
            //sheet.Range["a3:a4"].Style.Font.IsItalic = true;
            sheet.Range["a3:a4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a3:a4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a3:a4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a3:a4"].Style.Font.Size = 11;
            sheet.Range["a3:a4"].Merge(); // birlashtirish
            sheet.Range["a3:a4"].Text = "Наименования предмета";
            sheet.Range["a3:a4"].Style.WrapText = true;
            //sheet.Range["a3:a4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["a3:a4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["b3:b4"].Style.Font.IsBold = true;
            //sheet.Range["b3:b4"].Style.Font.IsItalic = true;
            sheet.Range["b3:b4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b3:b4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b3:b4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b3:b4"].Style.Font.Size = 10;
            sheet.Range["b3:b4"].Merge(); // birlashtirish
            sheet.Range["b3:b4"].Text = "Ед.из.";
            sheet.Range["b3:b4"].Style.WrapText = true;
            //sheet.Range["b3:b4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["b3:b4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["c3:c3"].Style.Font.IsBold = true;
            //sheet.Range["c3:c3"].Style.Font.IsItalic = true;
            sheet.Range["c3:d3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c3:d3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c3:d3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c3:d3"].Style.Font.Size = 10;
            sheet.Range["c3:d3"].Merge(); // birlashtirish
            sheet.Range["c3:d3"].Text = "ОСТАТОК на нач.";
            sheet.Range["c3:d3"].Style.WrapText = true;
            //sheet.Range["c3:c3"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["c3:d3"].BorderAround(LineStyleType.Thin);

            //sheet.Range["c3:c3"].Style.Font.IsBold = true;
            //sheet.Range["c3:c3"].Style.Font.IsItalic = true;
            sheet.Range["c4:c4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c4:c4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c4:c4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c4:c4"].Style.Font.Size = 10;
            sheet.Range["c4:c4"].Merge(); // birlashtirish
            sheet.Range["c4:c4"].Text = "Кол";
            sheet.Range["c4:c4"].Style.WrapText = true;
            //sheet.Range["c3:c3"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["c4:c4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["d4:d4"].Style.Font.IsBold = true;
            //sheet.Range["d4:d4"].Style.Font.IsItalic = true;
            sheet.Range["d4:d4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d4:d4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d4:d4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d4:d4"].Style.Font.Size = 10;
            sheet.Range["d4:d4"].Merge(); // birlashtirish
            sheet.Range["d4:d4"].Text = "Остаток";
            sheet.Range["d4:d4"].Style.WrapText = true;
            //sheet.Range["d4:d4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["d4:d4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["e3:e3"].Style.Font.IsBold = true;
            //sheet.Range["e3:e3"].Style.Font.IsItalic = true;
            sheet.Range["e3:h3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e3:h3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e3:h3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e3:h3"].Style.Font.Size = 10;
            sheet.Range["e3:h3"].Merge(); // birlashtirish
            sheet.Range["e3:h3"].Text = "ОБОРОТ";
            sheet.Range["e3:h3"].Style.WrapText = true;
            //sheet.Range["e3:e3"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["e3:h3"].BorderAround(LineStyleType.Thin);

            //sheet.Range["e3:e3"].Style.Font.IsBold = true;
            //sheet.Range["e3:e3"].Style.Font.IsItalic = true;
            sheet.Range["e4:e4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e4:e4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e4:e4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e4:e4"].Style.Font.Size = 10;
            sheet.Range["e4:e4"].Merge(); // birlashtirish
            sheet.Range["e4:e4"].Text = "Кол";
            sheet.Range["e4:e4"].Style.WrapText = true;
            //sheet.Range["e3:e3"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["e4:e4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["f4:4"].Style.Font.IsBold = true;
            //sheet.Range["f4:4"].Style.Font.IsItalic = true;
            sheet.Range["f4:f4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f4:f4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f4:f4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f4:f4"].Style.Font.Size = 10;
            sheet.Range["f4:f4"].Merge(); // birlashtirish
            sheet.Range["f4:f4"].Text = "Приход";
            sheet.Range["f4:f4"].Style.WrapText = true;
            //sheet.Range["f4:4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["f4:f4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["g4:g4"].Style.Font.IsBold = true;
            //sheet.Range["g4:g4"].Style.Font.IsItalic = true;
            sheet.Range["g4:g4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g4:g4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g4:g4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g4:g4"].Style.Font.Size = 10;
            sheet.Range["g4:g4"].Merge(); // birlashtirish
            sheet.Range["g4:g4"].Text = "Кол";
            sheet.Range["g4:g4"].Style.WrapText = true;
            //sheet.Range["g4:g4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["g4:g4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["h4:h4"].Style.Font.IsBold = true;
            //sheet.Range["h4:h4"].Style.Font.IsItalic = true;
            sheet.Range["h4:h4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h4:h4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h4:h4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["h4:h4"].Style.Font.Size = 10;
            sheet.Range["h4:h4"].Merge(); // birlashtirish
            sheet.Range["h4:h4"].Text = "Расход";
            sheet.Range["h4:h4"].Style.WrapText = true;
            //sheet.Range["h4:h4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["h4:h4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["i3:i3"].Style.Font.IsBold = true;
            //sheet.Range["i3:i3"].Style.Font.IsItalic = true;
            sheet.Range["i3:j3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["i3:j3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["i3:j3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["i3:j3"].Style.Font.Size = 10;
            sheet.Range["i3:j3"].Merge(); // birlashtirish
            sheet.Range["i3:j3"].Text = "ОСТАТОК на кон.";
            sheet.Range["i3:j3"].Style.WrapText = true;
            //sheet.Range["i3:i3"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["i3:j3"].BorderAround(LineStyleType.Thin);

            //sheet.Range["i4:i4"].Style.Font.IsBold = true;
            //sheet.Range["i4:i4"].Style.Font.IsItalic = true;
            sheet.Range["i4:i4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["i4:i4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["i4:i4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["i4:i4"].Style.Font.Size = 10;
            sheet.Range["i4:i4"].Merge(); // birlashtirish
            sheet.Range["i4:i4"].Text = "Кол";
            sheet.Range["i4:i4"].Style.WrapText = true;
            //sheet.Range["i4:i4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["i4:i4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["j4:j4"].Style.Font.IsBold = true;
            //sheet.Range["j4:j4"].Style.Font.IsItalic = true;
            sheet.Range["j4:j4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["j4:j4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["j4:j4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["j4:j4"].Style.Font.Size = 10;
            sheet.Range["j4:j4"].Merge(); // birlashtirish
            sheet.Range["j4:j4"].Text = "Остаток";
            sheet.Range["j4:j4"].Style.WrapText = true;
            //sheet.Range["j4:j4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["j4:j4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["k3:k4"].Style.Font.IsBold = true;
            //sheet.Range["k3:k4"].Style.Font.IsItalic = true;
            sheet.Range["k3:k4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["k3:k4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["k3:k4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["k3:k4"].Style.Font.Size = 10;
            sheet.Range["k3:k4"].Merge(); // birlashtirish
            sheet.Range["k3:k4"].Text = "Дата";
            sheet.Range["k3:k4"].Style.WrapText = true;
            //sheet.Range["k3:k4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["k3:k4"].BorderAround(LineStyleType.Thin);

            int i = 0;
            int myrow = 5;
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

            sheet.PageSetup.LeftMargin = 0.1;
            sheet.PageSetup.RightMargin = 0.1;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;


            sheet.PageSetup.Orientation = PageOrientationType.Portrait;


            sheet.Range["a1:a1"].ColumnWidth = 22.14;
            sheet.Range["b1:b1"].ColumnWidth = 5;
            sheet.Range["c1:c1"].ColumnWidth = 6;
            sheet.Range["d1:d1"].ColumnWidth = 10;
            sheet.Range["e1:e1"].ColumnWidth = 6;
            sheet.Range["f1:f1"].ColumnWidth = 10;
            sheet.Range["g1:g1"].ColumnWidth = 6;
            sheet.Range["h1:h1"].ColumnWidth = 10;
            sheet.Range["i1:i1"].ColumnWidth = 6;
            sheet.Range["j1:j1"].ColumnWidth = 10;
            sheet.Range["k1:k1"].ColumnWidth = 8.43;



            sheet.Range["f1:h1"].Style.Font.IsBold = true;
            sheet.Range["f1:h1"].Style.Font.IsItalic = true;
            sheet.Range["f1:h1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f1:h1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f1:h1"].Style.HorizontalAlignment = HorizontalAlignType.Right;
            sheet.Range["f1:h1"].Style.Font.Size = 11;
            sheet.Range["f1:h1"].Merge(); // birlashtirish
            sheet.Range["f1:h1"].Text = "ГУВД1";
            sheet.Range["f1:h1"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["f1:h1"].Style.Font.Underline = FontUnderlineType.Single;
            sheet.SetRowHeight(1, 18);


            sheet.Range["a2:g2"].Style.Font.IsBold = true;
            sheet.Range["a2:g2"].Style.Font.IsItalic = true;
            sheet.Range["a2:g2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a2:g2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a2:g2"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a2:g2"].Style.Font.Size = 14;
            sheet.Range["a2:g2"].Merge(); // birlashtirish
            sheet.Range["a2:g2"].Text = "СВОДНАЯ ОБОРОТЪ ЗА Декабръ 2020 год";
            sheet.Range["a2:g2"].Style.WrapText = true;
            sheet.Range["a2:g2"].Style.Font.Color = Color.DarkBlue;
            sheet.SetRowHeight(2, 20);


            //sheet.Range["a3:a4"].Style.Font.IsBold = true;
            //sheet.Range["a3:a4"].Style.Font.IsItalic = true;
            sheet.Range["a3:a4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a3:a4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a3:a4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a3:a4"].Style.Font.Size = 11;
            sheet.Range["a3:a4"].Merge(); // birlashtirish
            sheet.Range["a3:a4"].Text = "Наименования предмета";
            sheet.Range["a3:a4"].Style.WrapText = true;
            //sheet.Range["a3:a4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["a3:a4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["b3:b4"].Style.Font.IsBold = true;
            //sheet.Range["b3:b4"].Style.Font.IsItalic = true;
            sheet.Range["b3:b4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b3:b4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b3:b4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b3:b4"].Style.Font.Size = 10;
            sheet.Range["b3:b4"].Merge(); // birlashtirish
            sheet.Range["b3:b4"].Text = "Ед.из.";
            sheet.Range["b3:b4"].Style.WrapText = true;
            //sheet.Range["b3:b4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["b3:b4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["c3:c3"].Style.Font.IsBold = true;
            //sheet.Range["c3:c3"].Style.Font.IsItalic = true;
            sheet.Range["c3:d3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c3:d3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c3:d3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c3:d3"].Style.Font.Size = 10;
            sheet.Range["c3:d3"].Merge(); // birlashtirish
            sheet.Range["c3:d3"].Text = "ОСТАТОК на нач.";
            sheet.Range["c3:d3"].Style.WrapText = true;
            //sheet.Range["c3:c3"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["c3:d3"].BorderAround(LineStyleType.Thin);

            //sheet.Range["c3:c3"].Style.Font.IsBold = true;
            //sheet.Range["c3:c3"].Style.Font.IsItalic = true;
            sheet.Range["c4:c4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c4:c4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c4:c4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c4:c4"].Style.Font.Size = 10;
            sheet.Range["c4:c4"].Merge(); // birlashtirish
            sheet.Range["c4:c4"].Text = "Кол";
            sheet.Range["c4:c4"].Style.WrapText = true;
            //sheet.Range["c3:c3"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["c4:c4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["d4:d4"].Style.Font.IsBold = true;
            //sheet.Range["d4:d4"].Style.Font.IsItalic = true;
            sheet.Range["d4:d4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d4:d4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d4:d4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d4:d4"].Style.Font.Size = 10;
            sheet.Range["d4:d4"].Merge(); // birlashtirish
            sheet.Range["d4:d4"].Text = "Остаток";
            sheet.Range["d4:d4"].Style.WrapText = true;
            //sheet.Range["d4:d4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["d4:d4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["e3:e3"].Style.Font.IsBold = true;
            //sheet.Range["e3:e3"].Style.Font.IsItalic = true;
            sheet.Range["e3:h3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e3:h3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e3:h3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e3:h3"].Style.Font.Size = 10;
            sheet.Range["e3:h3"].Merge(); // birlashtirish
            sheet.Range["e3:h3"].Text = "ОБОРОТ";
            sheet.Range["e3:h3"].Style.WrapText = true;
            //sheet.Range["e3:e3"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["e3:h3"].BorderAround(LineStyleType.Thin);

            //sheet.Range["e3:e3"].Style.Font.IsBold = true;
            //sheet.Range["e3:e3"].Style.Font.IsItalic = true;
            sheet.Range["e4:e4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e4:e4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e4:e4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e4:e4"].Style.Font.Size = 10;
            sheet.Range["e4:e4"].Merge(); // birlashtirish
            sheet.Range["e4:e4"].Text = "Кол";
            sheet.Range["e4:e4"].Style.WrapText = true;
            //sheet.Range["e3:e3"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["e4:e4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["f4:4"].Style.Font.IsBold = true;
            //sheet.Range["f4:4"].Style.Font.IsItalic = true;
            sheet.Range["f4:f4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f4:f4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f4:f4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f4:f4"].Style.Font.Size = 10;
            sheet.Range["f4:f4"].Merge(); // birlashtirish
            sheet.Range["f4:f4"].Text = "Приход";
            sheet.Range["f4:f4"].Style.WrapText = true;
            //sheet.Range["f4:4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["f4:f4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["g4:g4"].Style.Font.IsBold = true;
            //sheet.Range["g4:g4"].Style.Font.IsItalic = true;
            sheet.Range["g4:g4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g4:g4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g4:g4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g4:g4"].Style.Font.Size = 10;
            sheet.Range["g4:g4"].Merge(); // birlashtirish
            sheet.Range["g4:g4"].Text = "Кол";
            sheet.Range["g4:g4"].Style.WrapText = true;
            //sheet.Range["g4:g4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["g4:g4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["h4:h4"].Style.Font.IsBold = true;
            //sheet.Range["h4:h4"].Style.Font.IsItalic = true;
            sheet.Range["h4:h4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h4:h4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h4:h4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["h4:h4"].Style.Font.Size = 10;
            sheet.Range["h4:h4"].Merge(); // birlashtirish
            sheet.Range["h4:h4"].Text = "Расход";
            sheet.Range["h4:h4"].Style.WrapText = true;
            //sheet.Range["h4:h4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["h4:h4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["i3:i3"].Style.Font.IsBold = true;
            //sheet.Range["i3:i3"].Style.Font.IsItalic = true;
            sheet.Range["i3:j3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["i3:j3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["i3:j3"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["i3:j3"].Style.Font.Size = 10;
            sheet.Range["i3:j3"].Merge(); // birlashtirish
            sheet.Range["i3:j3"].Text = "ОСТАТОК на кон.";
            sheet.Range["i3:j3"].Style.WrapText = true;
            //sheet.Range["i3:i3"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["i3:j3"].BorderAround(LineStyleType.Thin);

            //sheet.Range["i4:i4"].Style.Font.IsBold = true;
            //sheet.Range["i4:i4"].Style.Font.IsItalic = true;
            sheet.Range["i4:i4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["i4:i4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["i4:i4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["i4:i4"].Style.Font.Size = 10;
            sheet.Range["i4:i4"].Merge(); // birlashtirish
            sheet.Range["i4:i4"].Text = "Кол";
            sheet.Range["i4:i4"].Style.WrapText = true;
            //sheet.Range["i4:i4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["i4:i4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["j4:j4"].Style.Font.IsBold = true;
            //sheet.Range["j4:j4"].Style.Font.IsItalic = true;
            sheet.Range["j4:j4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["j4:j4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["j4:j4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["j4:j4"].Style.Font.Size = 10;
            sheet.Range["j4:j4"].Merge(); // birlashtirish
            sheet.Range["j4:j4"].Text = "Остаток";
            sheet.Range["j4:j4"].Style.WrapText = true;
            //sheet.Range["j4:j4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["j4:j4"].BorderAround(LineStyleType.Thin);

            //sheet.Range["k3:k4"].Style.Font.IsBold = true;
            //sheet.Range["k3:k4"].Style.Font.IsItalic = true;
            sheet.Range["k3:k4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["k3:k4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["k3:k4"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["k3:k4"].Style.Font.Size = 10;
            sheet.Range["k3:k4"].Merge(); // birlashtirish
            sheet.Range["k3:k4"].Text = "Дата";
            sheet.Range["k3:k4"].Style.WrapText = true;
            //sheet.Range["k3:k4"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["k3:k4"].BorderAround(LineStyleType.Thin);

            int i = 0;
            int myrow = 5;
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

        private void button2_Click(object sender, EventArgs e)
        {
            //try
            //{
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.LeftMargin = 0.1;
            sheet.PageSetup.RightMargin = 0.1;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;


            sheet.PageSetup.Orientation = PageOrientationType.Portrait;


            sheet.Range["a1:a1"].ColumnWidth = 3;
            sheet.Range["b1:b1"].ColumnWidth = 17.57;
            sheet.Range["c1:c1"].ColumnWidth = 8;
            sheet.Range["d1:d1"].ColumnWidth = 8.43;
            sheet.Range["e1:e1"].ColumnWidth = 4;
            sheet.Range["f1:f1"].ColumnWidth = 6;
            sheet.Range["g1:g1"].ColumnWidth = 9;
            sheet.Range["h1:h1"].ColumnWidth = 6;
            sheet.Range["i1:i1"].ColumnWidth = 8;
            sheet.Range["j1:j1"].ColumnWidth = 6;
            sheet.Range["k1:k1"].ColumnWidth = 8;
            sheet.Range["l1:l1"].ColumnWidth = 6;
            sheet.Range["m1:m1"].ColumnWidth = 8;



            //sheet.Range["f1:k1"].Style.Font.IsBold = true;
            //sheet.Range["f1:k1"].Style.Font.IsItalic = true;
            sheet.Range["f1:k1"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f1:k1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f1:k1"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["f1:k1"].Style.Font.Size = 11;
            sheet.Range["f1:k1"].Merge(); // birlashtirish
            sheet.Range["f1:k1"].Text = "Приложение_______________";
            //sheet.Range["f1:k1"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["f1:k1"].Style.Font.Underline = FontUnderlineType.Single;
            sheet.SetRowHeight(1, 20);

            //sheet.Range["f2:k2"].Style.Font.IsBold = true;
            //sheet.Range["f2:k2"].Style.Font.IsItalic = true;
            sheet.Range["f2:k2"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f2:k2"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f2:k2"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["f2:k2"].Style.Font.Size = 11;
            sheet.Range["f2:k2"].Merge(); // birlashtirish
            sheet.Range["f2:k2"].Text = "Графа 'значится по учёту' з аполняется";
            //sheet.Range["f2:k2"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["f2:k2"].Style.Font.Underline = FontUnderlineType.Single;

            //sheet.Range["f3:k3"].Style.Font.IsBold = true;
            //sheet.Range["f3:k3"].Style.Font.IsItalic = true;
            sheet.Range["f3:k3"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f3:k3"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f3:k3"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["f3:k3"].Style.Font.Size = 11;
            sheet.Range["f3:k3"].Merge(); // birlashtirish
            sheet.Range["f3:k3"].Text = "бухгалтерией ТОЛКО после снятия остаткое";
            //sheet.Range["f3:k3"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["f3:k3"].Style.Font.Underline = FontUnderlineType.Single;

            //sheet.Range["f4:k4"].Style.Font.IsBold = true;
            //sheet.Range["f4:k4"].Style.Font.IsItalic = true;
            sheet.Range["f4:k4"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f4:k4"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f4:k4"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["f4:k4"].Style.Font.Size = 11;
            sheet.Range["f4:k4"].Merge(); // birlashtirish
            sheet.Range["f4:k4"].Text = "материалъных ценностей";
            //sheet.Range["f4:k4"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["f4:k4"].Style.Font.Underline = FontUnderlineType.Single;


            sheet.Range["a5:k5"].Style.Font.IsBold = true;
            sheet.Range["a5:k5"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a5:k5"].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["a5:k5"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a5:k5"].Style.Font.Size = 11;
            sheet.Range["a5:k5"].Merge(); // birlashtirish
            sheet.Range["a5:k5"].Text = "АКТ       ________";
            sheet.Range["a5:k5"].Style.WrapText = true;
            sheet.SetRowHeight(5, 24);

            sheet.Range["a6:k6"].Style.Font.IsBold = true;
            sheet.Range["a6:k6"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a6:k6"].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["a6:k6"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a6:k6"].Style.Font.Size = 10;
            sheet.Range["a6:k6"].Merge(); // birlashtirish
            sheet.Range["a6:k6"].Text = "снятия остатков";
            sheet.Range["a6:k6"].Style.WrapText = true;
            sheet.SetRowHeight(6, 16);

            // sheet.Range["a7:m7"].Style.Font.IsBold = true;
            sheet.Range["a7:k7"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a7:k7"].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["a7:k7"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a7:k7"].Style.Font.Size = 10;
            sheet.Range["a7:k7"].Merge(); // birlashtirish
            sheet.Range["a7:k7"].Text = " '_____'  ________________________  20____г";
            sheet.Range["a7:k7"].Style.WrapText = true;
            sheet.SetRowHeight(7, 16);

            //sheet.Range["a8:k8"].Style.Font.IsBold = true;
            sheet.Range["a8:k8"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a8:k8"].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["a8:k8"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a8:k8"].Style.Font.Size = 10;
            sheet.Range["a8:k8"].Merge(); // birlashtirish
            sheet.Range["a8:k8"].Text = "   Мною,";
            sheet.Range["a8:k8"].Style.WrapText = true;
            sheet.Range["a8:k8"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.SetRowHeight(8, 18);

            //sheet.Range["a9:k9"].Style.Font.IsBold = true;
            sheet.Range["a9:k9"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a9:k9"].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["a9:k9"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a9:k9"].Style.Font.Size = 10;
            sheet.Range["a9:k9"].Merge(); // birlashtirish
            sheet.Range["a9:k9"].Text = "   на основании";
            sheet.Range["a9:k9"].Style.WrapText = true;
            sheet.Range["a9:k9"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.SetRowHeight(9, 18);

            //sheet.Range["a10:k10"].Style.Font.IsBold = true;
            sheet.Range["a10:k10"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a10:k10"].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["a10:k10"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a10:k10"].Style.Font.Size = 10;
            sheet.Range["a10:k10"].Merge(); // birlashtirish
            sheet.Range["a10:k10"].Text = "   в присутствин";
            sheet.Range["a10:k10"].Style.WrapText = true;
            sheet.Range["a10:k10"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.SetRowHeight(10, 18);

            //sheet.Range["a11:k11"].Style.Font.IsBold = true;
            sheet.Range["a11:k11"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a11:k11"].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["a11:k11"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a11:k11"].Style.Font.Size = 10;
            sheet.Range["a11:k11"].Merge(); // birlashtirish
            sheet.Range["a11:k11"].Text = "   и материалъно-ответственные лицфа :";
            sheet.Range["a11:k11"].Style.WrapText = true;
            sheet.Range["a11:k11"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.SetRowHeight(11, 18);

            //sheet.Range["a12:k12"].Style.Font.IsBold = true;
            sheet.Range["a12:k12"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a12:k12"].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["a12:k12"].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["a12:k12"].Style.Font.Size = 10;
            sheet.Range["a12:k12"].Merge(); // birlashtirish
            sheet.Range["a12:k12"].Text = "   произведено полное                                             снятие наличия остатков материалных ценностей";
            sheet.Range["a12:k12"].Style.WrapText = true;
            sheet.Range["a12:k12"].Style.Font.Underline = FontUnderlineType.SingleAccounting;
            sheet.SetRowHeight(12, 18);


            //sheet.Range["f13:k13"].Style.Font.IsBold = true;
            //sheet.Range["f13:k13"].Style.Font.IsItalic = true;
            sheet.Range["f13:k13"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f13:k13"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f13:k13"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f13:k13"].Style.Font.Size = 11;
            sheet.Range["f13:k13"].Merge(); // birlashtirish
            sheet.Range["f13:k13"].Text = "Февралъ 2021 год";
            //sheet.Range["f13:k13"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["f13:k13"].Style.Font.Underline = FontUnderlineType.Single;

            sheet.SetRowHeight(13, 16);

            //sheet.Range["a4:k13"].Style.Font.IsBold = true;
            //sheet.Range["a4:k13"].Style.Font.IsItalic = true;
            sheet.Range["a14:k14"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a14:k14"].Style.VerticalAlignment = VerticalAlignType.Bottom;
            sheet.Range["a14:k14"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a14:k14"].Style.Font.Size = 11;
            sheet.Range["a14:k14"].Merge(); // birlashtirish
            sheet.Range["a14:k14"].Text = " на_______ ' _________________________ 20____г.";
            //sheet.Range["a4:k13"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["a4:k13"].Style.Font.Underline = FontUnderlineType.Single;
            sheet.SetRowHeight(14, 16);

            sheet.Range["a15:k15"].Style.Font.IsBold = true;
            //sheet.Range["a15:k15"].Style.Font.IsItalic = true;
            sheet.Range["a15:k15"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a15:k15"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a15:k15"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a15:k15"].Style.Font.Size = 11;
            sheet.Range["a15:k15"].Merge(); // birlashtirish
            sheet.Range["a15:k15"].Text = "Подписка";
            //sheet.Range["a15:k15"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["a15:k15"].Style.Font.Underline = FontUnderlineType.Single;
            sheet.SetRowHeight(15, 18);

            sheet.Range["a16:k16"].Style.Font.IsBold = true;
            //sheet.Range["a16:k16"].Style.Font.IsItalic = true;
            sheet.Range["a16:k16"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a16:k16"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a16:k16"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a16:k16"].Style.Font.Size = 11;
            sheet.Range["a16:k16"].Merge(); // birlashtirish
            sheet.Range["a16:k16"].Text = "материалъно-ответственного лица";
            sheet.SetRowHeight(16, 16);
            //sheet.Range["a16:k16"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["a16:k16"].Style.Font.Underline = FontUnderlineType.Single;


            //sheet.Range["a17:k17"].Style.Font.IsBold = true;
            //sheet.Range["a17:k17"].Style.Font.IsItalic = true;
            sheet.Range["a17:k17"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a17:k17"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a17:k17"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a17:k17"].Style.Font.Size = 11;
            sheet.Range["a17:k17"].Merge(); // birlashtirish
            sheet.Range["a17:k17"].Text = "Все документы по приходно-расходным операциями по состоянию на '____' _____________201___года мною";
            //sheet.Range["a17:k17"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["a17:k17"].Style.Font.Underline = FontUnderlineType.Single;

            //sheet.Range["a18:k18"].Style.Font.IsBold = true;
            //sheet.Range["a18:k18"].Style.Font.IsItalic = true;
            sheet.Range["a18:k18"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a18:k18"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a18:k18"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a18:k18"].Style.Font.Size = 11;
            sheet.Range["a18:k18"].Merge(); // birlashtirish
            sheet.Range["a18:k18"].Text = "предавлены, других оправдательных документов на прием и выдачу имушесва (продовльствия) не";
            //sheet.Range["a18:k18"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["a18:k18"].Style.Font.Underline = FontUnderlineType.Single;


            //sheet.Range["a19:k19"].Style.Font.IsBold = true;
            //sheet.Range["a19:k19"].Style.Font.IsItalic = true;
            sheet.Range["a19:k19"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a19:k19"].Style.VerticalAlignment = VerticalAlignType.Top;
            sheet.Range["a19:k19"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a19:k19"].Style.Font.Size = 11;
            sheet.Range["a19:k19"].Merge(); // birlashtirish
            sheet.Range["a19:k19"].Text = "именю. Бездокументалъного отпуска и према ";
            //sheet.Range["a19:k19"].Style.Font.Color = Color.DarkBlue;
            //sheet.Range["a19:k19"].Style.Font.Underline = FontUnderlineType.Single;
            sheet.SetRowHeight(19, 18);

            sheet.Range["a20:a21"].Style.Font.IsBold = true;
            //sheet.Range["a20:a21"].Style.Font.IsItalic = true;
            sheet.Range["a20:a21"].Style.Font.FontName = "Times New Roman";
            sheet.Range["a20:a21"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["a20:a21"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["a20:a21"].Style.Font.Size = 10;
            sheet.Range["a20:a21"].Merge(); // birlashtirish
            sheet.Range["a20:a21"].Text = "№ пп";
            sheet.Range["a20:a21"].Style.WrapText = true;
            //sheet.Range["a20:a21"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["a20:a21"].BorderAround(LineStyleType.Thin);

            sheet.Range["b20:b21"].Style.Font.IsBold = true;
            //sheet.Range["b20:b21"].Style.Font.IsItalic = true;
            sheet.Range["b20:b21"].Style.Font.FontName = "Times New Roman";
            sheet.Range["b20:b21"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["b20:b21"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["b20:b21"].Style.Font.Size = 10;
            sheet.Range["b20:b21"].Merge(); // birlashtirish
            sheet.Range["b20:b21"].Text = "Наименования предмета";
            sheet.Range["b20:b21"].Style.WrapText = true;
            //sheet.Range["b20:b21"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["b20:b21"].BorderAround(LineStyleType.Thin);



            sheet.Range["c20:c21"].Style.Font.IsBold = true;
            //sheet.Range["c20:c21"].Style.Font.IsItalic = true;
            sheet.Range["c20:c21"].Style.Font.FontName = "Times New Roman";
            sheet.Range["c20:c21"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["c20:c21"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["c20:c21"].Style.Font.Size = 10;
            sheet.Range["c20:c21"].Merge(); // birlashtirish
            sheet.Range["c20:c21"].Text = "Инвентар номер";
            sheet.Range["c20:c21"].Style.WrapText = true;
            //sheet.Range["c20:c21"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["c20:c21"].BorderAround(LineStyleType.Thin);


            sheet.Range["d20:d21"].Style.Font.IsBold = true;
            //sheet.Range["d20:d21"].Style.Font.IsItalic = true;
            sheet.Range["d20:d21"].Style.Font.FontName = "Times New Roman";
            sheet.Range["d20:d21"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["d20:d21"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["d20:d21"].Style.Font.Size = 10;
            sheet.Range["d20:d21"].Merge(); // birlashtirish
            sheet.Range["d20:d21"].Text = "Дата выпуск";
            sheet.Range["d20:d21"].Style.WrapText = true;
            //sheet.Range["d20:d21"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["d20:d21"].BorderAround(LineStyleType.Thin);

            sheet.Range["e20:e21"].Style.Font.IsBold = true;
            //sheet.Range["e20:e21"].Style.Font.IsItalic = true;
            sheet.Range["e20:e21"].Style.Font.FontName = "Times New Roman";
            sheet.Range["e20:e21"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["e20:e21"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["e20:e21"].Style.Font.Size = 10;
            sheet.Range["e20:e21"].Merge(); // birlashtirish
            sheet.Range["e20:e21"].Text = "Ед.из";
            sheet.Range["e20:e21"].Style.WrapText = true;
            //sheet.Range["e20:e21"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["e20:e21"].BorderAround(LineStyleType.Thin);

            sheet.Range["f20:g20"].Style.Font.IsBold = true;
            //sheet.Range["f20:g20"].Style.Font.IsItalic = true;
            sheet.Range["f20:g20"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f20:g20"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f20:g20"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f20:g20"].Style.Font.Size = 10;
            sheet.Range["f20:g20"].Merge(); // birlashtirish
            sheet.Range["f20:g20"].Text = "Значится по учету";
            sheet.Range["f20:g20"].Style.WrapText = true;
            //sheet.Range["f20:g20"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["f20:g20"].BorderAround(LineStyleType.Thin);

            sheet.Range["f21:f21"].Style.Font.IsBold = true;
            //sheet.Range["f21:f21"].Style.Font.IsItalic = true;
            sheet.Range["f21:f21"].Style.Font.FontName = "Times New Roman";
            sheet.Range["f21:f21"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["f21:f21"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["f21:f21"].Style.Font.Size = 10;
            sheet.Range["f21:f21"].Merge(); // birlashtirish
            sheet.Range["f21:f21"].Text = "Кол";
            sheet.Range["f21:f21"].Style.WrapText = true;
            //sheet.Range["f21:f21"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["f21:f21"].BorderAround(LineStyleType.Thin);

            sheet.Range["g21:g21"].Style.Font.IsBold = true;
            //sheet.Range["g21:g21"].Style.Font.IsItalic = true;
            sheet.Range["g21:g21"].Style.Font.FontName = "Times New Roman";
            sheet.Range["g21:g21"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["g21:g21"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["g21:g21"].Style.Font.Size = 10;
            sheet.Range["g21:g21"].Merge(); // birlashtirish
            sheet.Range["g21:g21"].Text = "Сумма";
            sheet.Range["g21:g21"].Style.WrapText = true;
            //sheet.Range["g21:g21"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["g21:g21"].BorderAround(LineStyleType.Thin);

            sheet.Range["h20:i20"].Style.Font.IsBold = true;
            //sheet.Range["h20:i20"].Style.Font.IsItalic = true;
            sheet.Range["h20:i20"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h20:i20"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h20:i20"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["h20:i20"].Style.Font.Size = 10;
            sheet.Range["h20:i20"].Merge(); // birlashtirish
            sheet.Range["h20:i20"].Text = "Факти. нали.";
            sheet.Range["h20:i20"].Style.WrapText = true;
            //sheet.Range["h20:i20"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["h20:i20"].BorderAround(LineStyleType.Thin);

            sheet.Range["h21:h21"].Style.Font.IsBold = true;
            //sheet.Range["h21:h21"].Style.Font.IsItalic = true;
            sheet.Range["h21:h21"].Style.Font.FontName = "Times New Roman";
            sheet.Range["h21:h21"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["h21:h21"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["h21:h21"].Style.Font.Size = 10;
            sheet.Range["h21:h21"].Merge(); // birlashtirish
            sheet.Range["h21:h21"].Text = "Кол";
            sheet.Range["h21:h21"].Style.WrapText = true;
            //sheet.Range["h21:h21"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["h21:h21"].BorderAround(LineStyleType.Thin);

            sheet.Range["i21:i21"].Style.Font.IsBold = true;
            //sheet.Range["i21:i21"].Style.Font.IsItalic = true;
            sheet.Range["i21:i21"].Style.Font.FontName = "Times New Roman";
            sheet.Range["i21:i21"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["i21:i21"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["i21:i21"].Style.Font.Size = 10;
            sheet.Range["i21:i21"].Merge(); // birlashtirish
            sheet.Range["i21:i21"].Text = "Сумма";
            sheet.Range["i21:i21"].Style.WrapText = true;
            //sheet.Range["i21:i21"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["i21:i21"].BorderAround(LineStyleType.Thin);

            sheet.Range["j20:k20"].Style.Font.IsBold = true;
            //sheet.Range["j20:k20"].Style.Font.IsItalic = true;
            sheet.Range["j20:k20"].Style.Font.FontName = "Times New Roman";
            sheet.Range["j20:k20"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["j20:k20"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["j20:k20"].Style.Font.Size = 10;
            sheet.Range["j20:k20"].Merge(); // birlashtirish
            sheet.Range["j20:k20"].Text = "Недостача";
            sheet.Range["j20:k20"].Style.WrapText = true;
            //sheet.Range["j20:k20"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["j20:k20"].BorderAround(LineStyleType.Thin);

            sheet.Range["j21:j21"].Style.Font.IsBold = true;
            //sheet.Range["j21:j21"].Style.Font.IsItalic = true;
            sheet.Range["j21:j21"].Style.Font.FontName = "Times New Roman";
            sheet.Range["j21:j21"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["j21:j21"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["j21:j21"].Style.Font.Size = 10;
            sheet.Range["j21:j21"].Merge(); // birlashtirish
            sheet.Range["j21:j21"].Text = "Кол";
            sheet.Range["j21:j21"].Style.WrapText = true;
            //sheet.Range["j21:j21"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["j21:j21"].BorderAround(LineStyleType.Thin);

            sheet.Range["k21:k21"].Style.Font.IsBold = true;
            //sheet.Range["k21:k21"].Style.Font.IsItalic = true;
            sheet.Range["k21:k21"].Style.Font.FontName = "Times New Roman";
            sheet.Range["k21:k21"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["k21:k21"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["k21:k21"].Style.Font.Size = 10;
            sheet.Range["k21:k21"].Merge(); // birlashtirish
            sheet.Range["k21:k21"].Text = "Сумма";
            sheet.Range["k21:k21"].Style.WrapText = true;
            //sheet.Range["k21:k21"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["k21:k21"].BorderAround(LineStyleType.Thin);

            sheet.Range["l20:m20"].Style.Font.IsBold = true;
            //sheet.Range["l20:m20"].Style.Font.IsItalic = true;
            sheet.Range["l20:m20"].Style.Font.FontName = "Times New Roman";
            sheet.Range["l20:m20"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["l20:m20"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["l20:m20"].Style.Font.Size = 10;
            sheet.Range["l20:m20"].Merge(); // birlashtirish
            sheet.Range["l20:m20"].Text = "Излишки";
            sheet.Range["l20:m20"].Style.WrapText = true;
            //sheet.Range["l20:m20"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["l20:m20"].BorderAround(LineStyleType.Thin);

            sheet.Range["l21:l21"].Style.Font.IsBold = true;
            //sheet.Range["l21:l21"].Style.Font.IsItalic = true;
            sheet.Range["l21:l21"].Style.Font.FontName = "Times New Roman";
            sheet.Range["l21:l21"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["l21:l21"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["l21:l21"].Style.Font.Size = 10;
            sheet.Range["l21:l21"].Merge(); // birlashtirish
            sheet.Range["l21:l21"].Text = "Кол";
            sheet.Range["l21:l21"].Style.WrapText = true;
            //sheet.Range["l21:l21"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["l21:l21"].BorderAround(LineStyleType.Thin);

            sheet.Range["m21:m21"].Style.Font.IsBold = true;
            //sheet.Range["m21:m21"].Style.Font.IsItalic = true;
            sheet.Range["m21:m21"].Style.Font.FontName = "Times New Roman";
            sheet.Range["m21:m21"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["m21:m21"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["m21:m21"].Style.Font.Size = 10;
            sheet.Range["m21:m21"].Merge(); // birlashtirish
            sheet.Range["m21:m21"].Text = "Сумма";
            sheet.Range["m21:m21"].Style.WrapText = true;
            //sheet.Range["m21:m21"].Style.Font.Color = Color.DarkBlue;
            sheet.Range["m21:m21"].BorderAround(LineStyleType.Thin);




            int i = 0;
            int myrow = 22;
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


                sheet.Range["m" + myrow + ":m" + myrow].Merge();
                sheet.Range["m" + myrow + ":m" + myrow].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range["m" + myrow + ":m" + myrow].Style.WrapText = true;
                sheet.Range["m" + myrow + ":m" + myrow].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["m" + myrow + ":m" + myrow].BorderAround(LineStyleType.Thin);
                sheet.Range["m" + myrow + ":m" + myrow].Style.Font.Size = 10;
                sheet.Range["m" + myrow + ":m" + myrow].Style.Font.FontName = "Times New Roman";
                sheet.Range["m" + myrow + ":m" + myrow].Value = "";




                myrow = myrow + 1;
                i = i + 1;


            }

            myrow++;



            sheet.Range["d5:" + myrow + "k"].NumberFormat = "#,##0.00";


            workbook.SaveToFile(Environment.CurrentDirectory + "\\docs\\Журнал-ордер.xlsx", Spire.Xls.FileFormat.Version2007);
            System.Diagnostics.Process.Start(workbook.FileName);



            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Журнал-ордер_excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
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
    }
}
