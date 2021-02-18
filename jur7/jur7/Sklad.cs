using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tulpep.NotificationWindow;

namespace jur7
{
    public partial class Sklad : Form
    {
        Connect sql = new Connect();
        Connect sql2 = new Connect();
        Connect sql3 = new Connect();
        Connect sql4 = new Connect();
        Connect sql5 = new Connect();
        Connect sql6 = new Connect();

        public string string_for_otdels;
        public string month_global;
        public string year_global;
        public string pereosenka_visible;

        public string ot_kogo_1;
        public string ot_kogo_2;

        public string product_id = "";
        public string schet = "";
        public string naim = "";
        public string edin = "";
        public string gruppa = "";
        public string seria_num = "";
        public string inv_num = "";
        public string date_pr = "";
        public string deb_schet_2 = "";
        public string kr_schet = "";
        public string kr_schet_2 = "";

        public DataTable table = new DataTable();
        public Sklad(string string_for_otdels, string year_global, string month_global, string ot_kogo_1, string ot_kogo_2, string pereosenka_visible)
        {
            InitializeComponent();

            this.string_for_otdels = string_for_otdels;
            this.month_global = month_global;
            this.year_global = year_global;
            this.pereosenka_visible = pereosenka_visible;
            this.ot_kogo_1 = ot_kogo_1;
            this.ot_kogo_2 = ot_kogo_2;

            sql.Connection();
            sql2.Connection();
            sql3.Connection();
            sql4.Connection();
            sql5.Connection();
            sql6.Connection();

            sklad_load();
            label_update_prixod();
        }

        public string refresh_strings_to_mysql(string mystring)
        {
            string str = string.Format("{0:#0.0000}", Convert.ToDouble(mystring.Replace('.', ','))).Replace(',', '.');
            Console.WriteLine(str);
            return str;
        }

        public void sklad_load()
        {
            //try
            //{
            txt_edin_prixod.Items.Clear();
            sql.myReader = sql.return_MySqlCommand("SELECT distinct edin FROM products where edin is not null").ExecuteReader();

            while (sql.myReader.Read())
            {
                txt_edin_prixod.Items.Add(sql.myReader.GetString("edin"));
            }
            sql.myReader.Close();


            this.sklad_dataGridView.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.sklad_dataGridView_CellValueChanged);

            sklad_dataGridView.Rows.Clear();


            double pri_kol = 0;
            double ras_kol = 0;
            double vnut_ras_kol = 0;

            var products = " select t.id,t.vid_doc,t.kod_doc,t.product_id,t.gruppa,t.naim_tov,t.edin,t.inventar_num,t.seria_num,sum(t.kol) as kol,t.sena,sum(t.summa) as summa,t.deb_sch,t.deb_sch_2,t.kre_sch,t.kre_sch_2,t.provodka_iznos,t.provodka_iznos_2,t.summa_iznos,t.date_pr" +
                            " from(SELECT id, vid_doc, kod_doc, product_id, gruppa, naim_tov, edin, inventar_num, seria_num, sum(kol) as kol, sena, sum(summa) as summa, deb_sch, deb_sch_2, kre_sch, kre_sch_2, provodka_iznos, provodka_iznos_2, summa_iznos, date_pr FROM products" +
                            " where vid_doc = '1' and user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' and kol > 0 and komu_1 = '" + ot_kogo_1 + "' and komu_2 = '" + ot_kogo_2 + "'  group by product_id" +
                            " union all" +
                            " SELECT id, vid_doc, kod_doc, product_id, gruppa, naim_tov, edin, inventar_num, seria_num, sum(kol) as kol, sena, summa, deb_sch, deb_sch_2, kre_sch, kre_sch_2, provodka_iznos, provodka_iznos_2, summa_iznos, date_pr FROM products" +
                            " where vid_doc = '3' and user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' and kol > 0 and komu_1 = '" + ot_kogo_1 + "' and komu_2 = '" + ot_kogo_2 + "'  group by product_id) as t group by t.product_id ";
            sql.myReader = sql.return_MySqlCommand(products).ExecuteReader();

            while (sql.myReader.Read())
            {
                int index = sklad_dataGridView.Rows.Add();

                sklad_dataGridView.Rows[index].Cells[0].Value = (sql.myReader["id"] != DBNull.Value ? sql.myReader.GetString("id") : "");
                sklad_dataGridView.Rows[index].Cells[1].Value = (sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetString("product_id") : "");
                sklad_dataGridView.Rows[index].Cells[2].Value = (sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch") : "");
                sklad_dataGridView.Rows[index].Cells[3].Value = (sql.myReader["naim_tov"] != DBNull.Value ? sql.myReader.GetString("naim_tov") : "");
                sklad_dataGridView.Rows[index].Cells[4].Value = (sql.myReader["edin"] != DBNull.Value ? sql.myReader.GetString("edin") : "");


                pri_kol = (sql.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("kol").Replace(".", ","))) : 0);

                var products_pri = " SELECT sum(kol) as kol FROM products where vid_doc='2' and product_id='" + (sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetString("product_id") : "") + "' and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and ot_kogo='" + ot_kogo_1 + "' and ot_kogo_2='" + ot_kogo_2 + "' and kol > 0 ";

                sql2.myReader = sql2.return_MySqlCommand(products_pri).ExecuteReader();

                while (sql2.myReader.Read())
                {
                    ras_kol = (sql2.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql2.myReader.GetString("kol").Replace(".", ","))) : 0);
                }
                sql2.myReader.Close();

                var products_vnut = " SELECT id,sum(kol) as kol FROM products where vid_doc = '3' and product_id='" + (sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetString("product_id") : "") + "' and user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' and kol > 0 and ot_kogo = '" + ot_kogo_1 + "' and ot_kogo_2 = '" + ot_kogo_2 + "' group by product_id";

                sql2.myReader = sql2.return_MySqlCommand(products_vnut).ExecuteReader();

                while (sql2.myReader.Read())
                {
                    vnut_ras_kol = (sql2.myReader["kol"] != DBNull.Value ? sql2.myReader.GetDouble("kol") : 0);
                }
                sql2.myReader.Close();


                string kols = (pri_kol - ras_kol - vnut_ras_kol).ToString();//sql3.myReader["kol"] != DBNull.Value ? sql3.myReader.GetString("kol") : "";
                //sklad_dataGridView.Rows[index].Cells[5].Style.BackColor = Color.Yellow;
                if (kols.Length <= 3)
                {
                    sklad_dataGridView.Rows[index].Cells[5].Value = string.Format("{0:#0.00}", (pri_kol - ras_kol - vnut_ras_kol));
                }
                if (kols.Length > 3)
                {
                    sklad_dataGridView.Rows[index].Cells[5].Value = string.Format("{0:#,###.00}", (pri_kol - ras_kol - vnut_ras_kol));
                }

                string sena = sql.myReader["sena"] != DBNull.Value ? sql.myReader.GetString("sena") : "";

                if (sena.Length <= 3)
                {
                    sklad_dataGridView.Rows[index].Cells[6].Value = string.Format("{0:#0.00}", (sql.myReader["sena"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("sena").Replace(".", ","))) : 0));
                }
                if (sena.Length > 3)
                {
                    sklad_dataGridView.Rows[index].Cells[6].Value = string.Format("{0:#,###.00}", (sql.myReader["sena"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("sena").Replace(".", ","))) : 0));
                }

                string summa = ((pri_kol - ras_kol - vnut_ras_kol) * (sql.myReader["sena"] != DBNull.Value ? sql.myReader.GetDouble("sena") : 0)).ToString(); //sql3.myReader["summa"] != DBNull.Value ? sql3.myReader.GetString("summa") : "";

                if (summa.Length <= 3)
                {
                    sklad_dataGridView.Rows[index].Cells[7].Value = string.Format("{0:#0.00}", ((pri_kol - ras_kol - vnut_ras_kol) * (sql.myReader["sena"] != DBNull.Value ? sql.myReader.GetDouble("sena") : 0)));
                }
                if (summa.Length > 3)
                {
                    sklad_dataGridView.Rows[index].Cells[7].Value = string.Format("{0:#,###.00}", ((pri_kol - ras_kol - vnut_ras_kol) * (sql.myReader["sena"] != DBNull.Value ? sql.myReader.GetDouble("sena") : 0)));
                }

                string sum_iznos = sql.myReader["summa_iznos"] != DBNull.Value ? sql.myReader.GetString("summa_iznos") : "";

                if (sum_iznos.Length <= 3)
                {
                    sklad_dataGridView.Rows[index].Cells[8].Value = string.Format("{0:#0.00}", (sql.myReader["summa_iznos"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa_iznos").Replace(".", ","))) : 0));
                }
                if (sum_iznos.Length > 3)
                {
                    sklad_dataGridView.Rows[index].Cells[8].Value = string.Format("{0:#,###.00}", (sql.myReader["summa_iznos"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa_iznos").Replace(".", ","))) : 0));
                }


                //deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,id_products
                sklad_dataGridView.Rows[index].Cells[9].Value = ("0");
                sklad_dataGridView.Rows[index].Cells[9].Style.BackColor = Color.GreenYellow;
                sklad_dataGridView.Rows[index].Cells[10].Value = ("0");
                sklad_dataGridView.Rows[index].Cells[11].Value = (sql.myReader["gruppa"] != DBNull.Value ? sql.myReader.GetString("gruppa") : "");
                sklad_dataGridView.Rows[index].Cells[13].Value = (sql.myReader["seria_num"] != DBNull.Value ? sql.myReader.GetString("seria_num") : "");
                sklad_dataGridView.Rows[index].Cells[14].Value = (sql.myReader["inventar_num"] != DBNull.Value ? sql.myReader.GetString("inventar_num") : "");
                sklad_dataGridView.Rows[index].Cells[15].Value = (sql.myReader["date_pr"] != DBNull.Value ? (DateTime.Parse(sql.myReader.GetString("date_pr")).ToString("dd.MM.yyyy")) : null);
                sklad_dataGridView.Rows[index].Cells[16].Value = (sql.myReader["deb_sch_2"] != DBNull.Value ? sql.myReader.GetString("deb_sch_2") : "");
                sklad_dataGridView.Rows[index].Cells[17].Value = (sql.myReader["kre_sch"] != DBNull.Value ? sql.myReader.GetString("kre_sch") : "");
                sklad_dataGridView.Rows[index].Cells[18].Value = (sql.myReader["kre_sch_2"] != DBNull.Value ? sql.myReader.GetString("kre_sch_2") : "");


                //deb_sch_2,kre_sch,kre_sch_2
            }
            sql.myReader.Close();


            this.sklad_dataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.sklad_dataGridView_CellValueChanged);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Хато маълумот киритилган (" + ex.Message + ")");
            //}
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

        public void label_update_prixod()
        {
            double kol = 0;
            double summa = 0;


            foreach (DataGridViewRow row in sklad_dataGridView.Rows)
            {
                kol = kol + (row.Cells[5].Value != null ? Double.Parse(row.Cells[5].Value.ToString()) : 0);

                summa = summa + (row.Cells[7].Value != null ? Double.Parse(row.Cells[7].Value.ToString()) : 0);

            }
            if (kol.ToString().Length <= 3)
            {
                kol_ostatok.Text = string.Format("{0:#0.00}", kol);
            }
            if (kol.ToString().Length > 3)
            {
                kol_ostatok.Text = string.Format("{0:#0,000.00}", kol);
            }

            if (summa.ToString().Length <= 3)
            {
                summa_ostatok.Text = string.Format("{0:#0.00}", summa);
            }
            if (summa.ToString().Length > 3)
            {
                summa_ostatok.Text = string.Format("{0:#0,000.00}", summa);
            }

        }


        private void sklad_dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (sklad_dataGridView.CurrentRow != null)
                {
                    DataGridViewRow dgvRow = sklad_dataGridView.CurrentRow;

                    if (e.ColumnIndex == 5)
                    {
                        // Console.WriteLine(dgvRow.Cells[7].Value);
                        dgvRow.Cells[7].Value = string.Format("{0:#0.00}", (dgvRow.Cells[5].Value != null ? (Convert.ToDouble(dgvRow.Cells[5].Value.ToString().Replace(".", ","))) : 0) *
                                                                           (dgvRow.Cells[6].Value != null ? (Convert.ToDouble(dgvRow.Cells[6].Value.ToString().Replace(".", ","))) : 0)
                                                                           );
                    }

                    if (e.ColumnIndex == 6)
                    {
                        dgvRow.Cells[7].Value = string.Format("{0:#0.00}", (dgvRow.Cells[5].Value != null ? (Convert.ToDouble(dgvRow.Cells[5].Value.ToString().Replace(".", ","))) : 0) *
                                                                           (dgvRow.Cells[6].Value != null ? (Convert.ToDouble(dgvRow.Cells[6].Value.ToString().Replace(".", ","))) : 0));
                    }

                    if (e.ColumnIndex == 9)
                    {
                        dgvRow.Cells[9].Value = string.Format("{0:#0.00}", ((dgvRow.Cells[9].Value != null ? (Convert.ToDouble(dgvRow.Cells[9].Value.ToString().Replace(".", ","))) : 0)));
                        dgvRow.Cells[10].Value = string.Format("{0:#0.00}", (dgvRow.Cells[9].Value != null ? (Convert.ToDouble(dgvRow.Cells[9].Value.ToString().Replace(".", ","))) : 0) *
                                                                       (dgvRow.Cells[6].Value != null ? (Convert.ToDouble(dgvRow.Cells[6].Value.ToString().Replace(".", ","))) : 0));
                        label_update_rasxod();

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("sklad_dataGridView_CellValueChanged_1 " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void label_update_rasxod()
        {
            try
            {
                double summa = 0;
                foreach (DataGridViewRow row in sklad_dataGridView.Rows)
                {
                    summa = summa + (row.Cells[10].Value != null ? Double.Parse(row.Cells[10].Value.ToString()) : 0);
                }
                if (summa.ToString().Length <= 3)
                {
                    rasxod_obshiy_summa_label.Text = string.Format("{0:#0.00}", summa);
                }
                if (summa.ToString().Length > 3)
                {
                    rasxod_obshiy_summa_label.Text = string.Format("{0:#0,0.00}", summa);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("label_update_rasxod " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        private void Sklad_Load(object sender, EventArgs e)
        {
            if(pereosenka_visible == "1")
            {
                pri_pereosenka_btn.Visible = true;
            }
            else
            {
                pri_pereosenka_btn.Visible = false;
            }

            this.sklad_dataGridView.RowsDefaultCellStyle.BackColor = Color.White;
            this.sklad_dataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(233, 233, 234);



            sklad_dataGridView.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            sklad_dataGridView.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            sklad_dataGridView.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            sklad_dataGridView.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            sklad_dataGridView.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            sklad_dataGridView.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            sklad_dataGridView.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            sklad_dataGridView.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            sklad_dataGridView.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            sklad_dataGridView.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            sklad_dataGridView.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            sklad_dataGridView.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            sklad_dataGridView.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            sklad_dataGridView.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        private void sklad_dataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //try
            //{
            DataGridViewRow dgvRow = sklad_dataGridView.CurrentRow;
            if (e.ColumnIndex == 9)
            {

                dgvRow.Cells[9].Value = string.Format("{0:#0.00}", ((dgvRow.Cells[5].Value != null ? (Convert.ToDouble(dgvRow.Cells[5].Value.ToString().Replace(".", ","))) : 0)));
                dgvRow.Cells[7].Value = string.Format("{0:#0.00}", (dgvRow.Cells[10].Value != null ? (Convert.ToDouble(dgvRow.Cells[10].Value.ToString().Replace(".", ","))) : 0));


            }
            if (e.ColumnIndex == 3)
            {
                //public string product_id = "";
                //public string schet = "";
                //public string naim = "";
                //public string edin = "";
                //public string gruppa = "";
                //public string seria_num = "";
                //public string inv_num = "";

                product_id = dgvRow.Cells[1].Value.ToString();
                schet = dgvRow.Cells[2].Value.ToString();
                naim = dgvRow.Cells[3].Value.ToString();
                edin = dgvRow.Cells[4].Value.ToString();
                gruppa = dgvRow.Cells[11].Value.ToString();
                seria_num = dgvRow.Cells[13].Value.ToString();
                inv_num = dgvRow.Cells[14].Value.ToString();
                date_pr = dgvRow.Cells[15].Value.ToString();
                deb_schet_2 = dgvRow.Cells[16].Value.ToString();
                kr_schet = dgvRow.Cells[17].Value.ToString();
                kr_schet_2 = dgvRow.Cells[18].Value.ToString();
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("sklad_dataGridView_CellDoubleClick " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private DataTable GetDataTableFromDGV(DataGridView dgv)
        {

            var dt = new DataTable();
            DataGridViewRow row1 = sklad_dataGridView.CurrentRow;
            foreach (DataGridViewColumn column in dgv.Columns)
            {
                dt.Columns.Add();
            }
            DataGridViewRow row11 = sklad_dataGridView.CurrentRow;
            object[] cellValues = new object[dgv.Columns.Count];

            foreach (DataGridViewRow row in dgv.Rows)
            {
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    //if (row.Cells[9].Value != null && Convert.ToDouble(row.Cells[9].Value.ToString().Replace(".", ",")) != 0 && Convert.ToString(row.Cells[9].Value.ToString()) != "0")
                    //   {
                    cellValues[i] = row.Cells[i].Value;
                    //}
                }
                if (row.Cells[9].Value != null && Convert.ToDouble(row.Cells[9].Value.ToString().Replace(".", ",")) != 0 && Convert.ToString(row.Cells[9].Value.ToString()) != "0")
                {
                    dt.Rows.Add(cellValues);
                }
            }

            return dt;
        }
        private void vzyat_btn_Click(object sender, EventArgs e)
        {
            try
            {
                DataGridViewRow row = sklad_dataGridView.CurrentRow;
                table = GetDataTableFromDGV(sklad_dataGridView);
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("toolStripMenuItem1_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        string gruppa_naim = "";
        string number_pereosenka = "";
        double pereosenka_foiz = 0;

        int exist = 0;
        private void pri_pereosenka_btn_Click(object sender, EventArgs e)
        {
            //try
            //{
            DialogResult dialogResult = MessageBox.Show("Вы уверены ? ", "Переоценка", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {

                var ex = " SELECT exists (SELECT * FROM pereosenka where user='"+string_for_otdels+"' and year='"+year_global+"' and month='"+month_global+"' and podraz_1='"+ot_kogo_1+"' and podraz_2 = '"+ot_kogo_2+"') as ex ";
                sql6.myReader = sql6.return_MySqlCommand(ex).ExecuteReader();
                while (sql6.myReader.Read())
                {
                    exist = sql6.myReader.GetInt32("ex");
                }
                sql6.myReader.Close();
                if (exist == 1)
                {
                    DialogResult dialogResult2 = MessageBox.Show("Обновлено в этом месяце ? ", "Обновление", MessageBoxButtons.YesNo);
                    if (dialogResult2 == DialogResult.Yes)
                    {

                        foreach (DataGridViewRow row in sklad_dataGridView.Rows)
                        {

                            if (row.Cells[1].Value != null && row.Cells[0].Value != null)
                            {

                                var data_between = " select * from pereosenka_data where '" + (row.Cells[15].Value != null ? DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") : "") + "' between date_start and date_finish ";
                                sql4.myReader = sql4.return_MySqlCommand(data_between).ExecuteReader();
                                while (sql4.myReader.Read())
                                {
                                    number_pereosenka = (sql4.myReader["number"] != DBNull.Value ? sql4.myReader.GetString("number") : "");


                                    var gruppa = "SELECT gruppa FROM gruppa where kod_gruppa='" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'";
                                    sql2.myReader = sql2.return_MySqlCommand(gruppa).ExecuteReader();
                                    while (sql2.myReader.Read())
                                    {
                                        gruppa_naim = (sql2.myReader["gruppa"] != DBNull.Value ? sql2.myReader.GetString("gruppa") : "");

                                        var gruppa0 = " SELECT * FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                        sql3.myReader = sql3.return_MySqlCommand(gruppa0).ExecuteReader();
                                        while (sql3.myReader.Read())
                                        {

                                            switch (number_pereosenka)
                                            {
                                                case "1":
                                                    {
                                                        var number = " SELECT id,first FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                                        sql5.myReader = sql5.return_MySqlCommand(number).ExecuteReader();
                                                        while (sql5.myReader.Read())
                                                        {
                                                            pereosenka_foiz = (sql5.myReader["first"] != DBNull.Value ? sql5.myReader.GetDouble("first") : 1);

                                                            var update_sklad = "update products set " +
                                                                               "sena = '" + refresh_strings_to_mysql(((row.Cells[6].Value != null ? (Convert.ToDouble(row.Cells[6].Value.ToString().Replace(".", ","))) : 0) * pereosenka_foiz).ToString()) + "'," +
                                                                               "summa = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'" +
                                                                               " where product_id = " + row.Cells[1].Value + " and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and komu_1='" + ot_kogo_1 + "' and komu_2='" + ot_kogo_2 + "' ";
                                                            sql.return_MySqlCommand(update_sklad).ExecuteNonQuery();
                                                        }
                                                        sql5.myReader.Close();

                                                        var insert_product = "insert into pereosenka (user,year,month,product_id,deb_sch,naim_tov,edin,kol,sena,summa,gruppa,seria_num,inventar_num,date_pr,deb_sch_2,kre_sch,kre_sch_2,podraz_1,podraz_2,gruppa0_naim,num_period,pereosenka_foiz) values(" +
                                                                                "'" + string_for_otdels + "'," +
                                                                                "'" + (year_global) + "'," +
                                                                                "'" + (month_global) + "'," +
                                                                                "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[5].Value != null ? row.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[6].Value != null ? row.Cells[6].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                                                "" + (row.Cells[15].Value != null ? "'" + DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") + "'" : DateTime.Now.ToString("yyyy-MM-dd")) + ", " +
                                                                                "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                                                "'" + (ot_kogo_1) + "'," +
                                                                                "'" + (ot_kogo_2) + "'," +
                                                                                "'" + (gruppa_naim) + "'," +
                                                                                "'" + (number_pereosenka) + "'," +
                                                                                "'" + ((pereosenka_foiz).ToString()).Replace(',', '.') + "'" +
                                                                                ")";
                                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();


                                                        break;
                                                    }
                                                case "2":
                                                    {
                                                        var number = " SELECT id,first FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                                        sql5.myReader = sql5.return_MySqlCommand(number).ExecuteReader();
                                                        while (sql5.myReader.Read())
                                                        {
                                                            pereosenka_foiz = (sql5.myReader["first"] != DBNull.Value ? sql5.myReader.GetDouble("first") : 1);

                                                            var update_sklad = "update products set " +
                                                                               "sena = '" + refresh_strings_to_mysql(((row.Cells[6].Value != null ? (Convert.ToDouble(row.Cells[6].Value.ToString().Replace(".", ","))) : 0) * pereosenka_foiz).ToString()) + "'," +
                                                                               "summa = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'" +
                                                                               " where product_id = " + row.Cells[1].Value + " and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and komu_1='" + ot_kogo_1 + "' and komu_2='" + ot_kogo_2 + "' ";
                                                            sql.return_MySqlCommand(update_sklad).ExecuteNonQuery();
                                                        }
                                                        sql5.myReader.Close();

                                                        var insert_product = "insert into pereosenka (user,year,month,product_id,deb_sch,naim_tov,edin,kol,sena,summa,gruppa,seria_num,inventar_num,date_pr,deb_sch_2,kre_sch,kre_sch_2,podraz_1,podraz_2,gruppa0_naim,num_period,pereosenka_foiz) values(" +
                                                                                "'" + string_for_otdels + "'," +
                                                                                "'" + (year_global) + "'," +
                                                                                "'" + (month_global) + "'," +
                                                                                "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[5].Value != null ? row.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[6].Value != null ? row.Cells[6].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                                                "" + (row.Cells[15].Value != null ? "'" + DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") + "'" : DateTime.Now.ToString("yyyy-MM-dd")) + ", " +
                                                                                "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                                                "'" + (ot_kogo_1) + "'," +
                                                                                "'" + (ot_kogo_2) + "'," +
                                                                                "'" + (gruppa_naim) + "'," +
                                                                                "'" + (number_pereosenka) + "'," +
                                                                                "'" + ((pereosenka_foiz).ToString()).Replace(',', '.') + "'" +
                                                                                ")";
                                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();


                                                        break;
                                                    }
                                                case "3":
                                                    {
                                                        var number = " SELECT id,first FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                                        sql5.myReader = sql5.return_MySqlCommand(number).ExecuteReader();
                                                        while (sql5.myReader.Read())
                                                        {
                                                            pereosenka_foiz = (sql5.myReader["first"] != DBNull.Value ? sql5.myReader.GetDouble("first") : 1);

                                                            var update_sklad = "update products set " +
                                                                               "sena = '" + refresh_strings_to_mysql(((row.Cells[6].Value != null ? (Convert.ToDouble(row.Cells[6].Value.ToString().Replace(".", ","))) : 0) * pereosenka_foiz).ToString()) + "'," +
                                                                               "summa = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'" +
                                                                               " where product_id = " + row.Cells[1].Value + " and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and komu_1='" + ot_kogo_1 + "' and komu_2='" + ot_kogo_2 + "' ";
                                                            sql.return_MySqlCommand(update_sklad).ExecuteNonQuery();
                                                        }
                                                        sql5.myReader.Close();

                                                        var insert_product = "insert into pereosenka (user,year,month,product_id,deb_sch,naim_tov,edin,kol,sena,summa,gruppa,seria_num,inventar_num,date_pr,deb_sch_2,kre_sch,kre_sch_2,podraz_1,podraz_2,gruppa0_naim,num_period,pereosenka_foiz) values(" +
                                                                                "'" + string_for_otdels + "'," +
                                                                                "'" + (year_global) + "'," +
                                                                                "'" + (month_global) + "'," +
                                                                                "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[5].Value != null ? row.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[6].Value != null ? row.Cells[6].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                                                "" + (row.Cells[15].Value != null ? "'" + DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") + "'" : DateTime.Now.ToString("yyyy-MM-dd")) + ", " +
                                                                                "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                                                "'" + (ot_kogo_1) + "'," +
                                                                                "'" + (ot_kogo_2) + "'," +
                                                                                "'" + (gruppa_naim) + "'," +
                                                                                "'" + (number_pereosenka) + "'," +
                                                                                "'" + ((pereosenka_foiz).ToString()).Replace(',', '.') + "'" +
                                                                                ")";
                                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();


                                                        break;
                                                    }
                                                case "4":
                                                    {
                                                        var number = " SELECT id,first FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                                        sql5.myReader = sql5.return_MySqlCommand(number).ExecuteReader();
                                                        while (sql5.myReader.Read())
                                                        {
                                                            pereosenka_foiz = (sql5.myReader["first"] != DBNull.Value ? sql5.myReader.GetDouble("first") : 1);

                                                            var update_sklad = "update products set " +
                                                                               "sena = '" + refresh_strings_to_mysql(((row.Cells[6].Value != null ? (Convert.ToDouble(row.Cells[6].Value.ToString().Replace(".", ","))) : 0) * pereosenka_foiz).ToString()) + "'," +
                                                                               "summa = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'" +
                                                                               " where product_id = " + row.Cells[1].Value + " and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and komu_1='" + ot_kogo_1 + "' and komu_2='" + ot_kogo_2 + "' ";
                                                            sql.return_MySqlCommand(update_sklad).ExecuteNonQuery();
                                                        }
                                                        sql5.myReader.Close();

                                                        var insert_product = "insert into pereosenka (user,year,month,product_id,deb_sch,naim_tov,edin,kol,sena,summa,gruppa,seria_num,inventar_num,date_pr,deb_sch_2,kre_sch,kre_sch_2,podraz_1,podraz_2,gruppa0_naim,num_period,pereosenka_foiz) values(" +
                                                                                "'" + string_for_otdels + "'," +
                                                                                "'" + (year_global) + "'," +
                                                                                "'" + (month_global) + "'," +
                                                                                "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[5].Value != null ? row.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[6].Value != null ? row.Cells[6].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                                                "" + (row.Cells[15].Value != null ? "'" + DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") + "'" : DateTime.Now.ToString("yyyy-MM-dd")) + ", " +
                                                                                "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                                                "'" + (ot_kogo_1) + "'," +
                                                                                "'" + (ot_kogo_2) + "'," +
                                                                                "'" + (gruppa_naim) + "'," +
                                                                                "'" + (number_pereosenka) + "'," +
                                                                                "'" + ((pereosenka_foiz).ToString()).Replace(',', '.') + "'" +
                                                                                ")";
                                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();


                                                        break;
                                                    }
                                                case "5":
                                                    {
                                                        var number = " SELECT id,first FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                                        sql5.myReader = sql5.return_MySqlCommand(number).ExecuteReader();
                                                        while (sql5.myReader.Read())
                                                        {
                                                            pereosenka_foiz = (sql5.myReader["first"] != DBNull.Value ? sql5.myReader.GetDouble("first") : 1);

                                                            var update_sklad = "update products set " +
                                                                               "sena = '" + refresh_strings_to_mysql(((row.Cells[6].Value != null ? (Convert.ToDouble(row.Cells[6].Value.ToString().Replace(".", ","))) : 0) * pereosenka_foiz).ToString()) + "'," +
                                                                               "summa = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'" +
                                                                               " where product_id = " + row.Cells[1].Value + " and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and komu_1='" + ot_kogo_1 + "' and komu_2='" + ot_kogo_2 + "' ";
                                                            sql.return_MySqlCommand(update_sklad).ExecuteNonQuery();
                                                        }
                                                        sql5.myReader.Close();

                                                        var insert_product = "insert into pereosenka (user,year,month,product_id,deb_sch,naim_tov,edin,kol,sena,summa,gruppa,seria_num,inventar_num,date_pr,deb_sch_2,kre_sch,kre_sch_2,podraz_1,podraz_2,gruppa0_naim,num_period,pereosenka_foiz) values(" +
                                                                                 "'" + string_for_otdels + "'," +
                                                                                 "'" + (year_global) + "'," +
                                                                                 "'" + (month_global) + "'," +
                                                                                 "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                                                   "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                                                 "'" + refresh_strings_to_mysql(row.Cells[5].Value != null ? row.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                 "'" + refresh_strings_to_mysql(row.Cells[6].Value != null ? row.Cells[6].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                 "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                 "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                                                   "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                                                   "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                                                 "" + (row.Cells[15].Value != null ? "'" + DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") + "'" : DateTime.Now.ToString("yyyy-MM-dd")) + ", " +
                                                                                 "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                                                 "'" + (ot_kogo_1) + "'," +
                                                                                 "'" + (ot_kogo_2) + "'," +
                                                                                 "'" + (gruppa_naim) + "'," +
                                                                                 "'" + (number_pereosenka) + "'," +
                                                                                 "'" + ((pereosenka_foiz).ToString()).Replace(',', '.') + "'" +
                                                                                 ")";
                                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();


                                                        break;
                                                    }
                                                case "6":
                                                    {
                                                        var number = " SELECT id,first FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                                        sql5.myReader = sql5.return_MySqlCommand(number).ExecuteReader();
                                                        while (sql5.myReader.Read())
                                                        {
                                                            pereosenka_foiz = (sql5.myReader["first"] != DBNull.Value ? sql5.myReader.GetDouble("first") : 1);

                                                            var update_sklad = "update products set " +
                                                                               "sena = '" + refresh_strings_to_mysql(((row.Cells[6].Value != null ? (Convert.ToDouble(row.Cells[6].Value.ToString().Replace(".", ","))) : 0) * pereosenka_foiz).ToString()) + "'," +
                                                                               "summa = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'" +
                                                                               " where product_id = " + row.Cells[1].Value + " and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and komu_1='" + ot_kogo_1 + "' and komu_2='" + ot_kogo_2 + "' ";
                                                            sql.return_MySqlCommand(update_sklad).ExecuteNonQuery();
                                                        }
                                                        sql5.myReader.Close();

                                                        var insert_product = "insert into pereosenka (user,year,month,product_id,deb_sch,naim_tov,edin,kol,sena,summa,gruppa,seria_num,inventar_num,date_pr,deb_sch_2,kre_sch,kre_sch_2,podraz_1,podraz_2,gruppa0_naim,num_period,pereosenka_foiz) values(" +
                                                                                "'" + string_for_otdels + "'," +
                                                                                "'" + (year_global) + "'," +
                                                                                "'" + (month_global) + "'," +
                                                                                "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[5].Value != null ? row.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[6].Value != null ? row.Cells[6].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                                                "" + (row.Cells[15].Value != null ? "'" + DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") + "'" : DateTime.Now.ToString("yyyy-MM-dd")) + ", " +
                                                                                "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                                                "'" + (ot_kogo_1) + "'," +
                                                                                "'" + (ot_kogo_2) + "'," +
                                                                                "'" + (gruppa_naim) + "'," +
                                                                                "'" + (number_pereosenka) + "'," +
                                                                                "'" + ((pereosenka_foiz).ToString()).Replace(',', '.') + "'" +
                                                                                ")";
                                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();


                                                        break;
                                                    }
                                                case "7":
                                                    {
                                                        var number = " SELECT id,first FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                                        sql5.myReader = sql5.return_MySqlCommand(number).ExecuteReader();
                                                        while (sql5.myReader.Read())
                                                        {
                                                            pereosenka_foiz = (sql5.myReader["first"] != DBNull.Value ? sql5.myReader.GetDouble("first") : 1);

                                                            var update_sklad = "update products set " +
                                                                               "sena = '" + refresh_strings_to_mysql(((row.Cells[6].Value != null ? (Convert.ToDouble(row.Cells[6].Value.ToString().Replace(".", ","))) : 0) * pereosenka_foiz).ToString()) + "'," +
                                                                               "summa = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'" +
                                                                               " where product_id = " + row.Cells[1].Value + " and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and komu_1='" + ot_kogo_1 + "' and komu_2='" + ot_kogo_2 + "' ";
                                                            sql.return_MySqlCommand(update_sklad).ExecuteNonQuery();
                                                        }
                                                        sql5.myReader.Close();

                                                        var insert_product = "insert into pereosenka (user,year,month,product_id,deb_sch,naim_tov,edin,kol,sena,summa,gruppa,seria_num,inventar_num,date_pr,deb_sch_2,kre_sch,kre_sch_2,podraz_1,podraz_2,gruppa0_naim,num_period,pereosenka_foiz) values(" +
                                                                                 "'" + string_for_otdels + "'," +
                                                                                 "'" + (year_global) + "'," +
                                                                                 "'" + (month_global) + "'," +
                                                                                 "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                                                   "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                                                 "'" + refresh_strings_to_mysql(row.Cells[5].Value != null ? row.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                 "'" + refresh_strings_to_mysql(row.Cells[6].Value != null ? row.Cells[6].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                 "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                 "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                                                   "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                                                   "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                                                 "" + (row.Cells[15].Value != null ? "'" + DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") + "'" : DateTime.Now.ToString("yyyy-MM-dd")) + ", " +
                                                                                 "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                                                 "'" + (ot_kogo_1) + "'," +
                                                                                 "'" + (ot_kogo_2) + "'," +
                                                                                 "'" + (gruppa_naim) + "'," +
                                                                                 "'" + (number_pereosenka) + "'," +
                                                                                 "'" + ((pereosenka_foiz).ToString()).Replace(',', '.') + "'" +
                                                                                 ")";
                                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();


                                                        break;
                                                    }
                                                case "8":
                                                    {
                                                        var number = " SELECT id,first FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                                        sql5.myReader = sql5.return_MySqlCommand(number).ExecuteReader();
                                                        while (sql5.myReader.Read())
                                                        {
                                                            pereosenka_foiz = (sql5.myReader["first"] != DBNull.Value ? sql5.myReader.GetDouble("first") : 1);

                                                            var update_sklad = "update products set " +
                                                                               "sena = '" + refresh_strings_to_mysql(((row.Cells[6].Value != null ? (Convert.ToDouble(row.Cells[6].Value.ToString().Replace(".", ","))) : 0) * pereosenka_foiz).ToString()) + "'," +
                                                                               "summa = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'" +
                                                                               " where product_id = " + row.Cells[1].Value + " and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and komu_1='" + ot_kogo_1 + "' and komu_2='" + ot_kogo_2 + "' ";
                                                            sql.return_MySqlCommand(update_sklad).ExecuteNonQuery();
                                                        }
                                                        sql5.myReader.Close();

                                                        var insert_product = "insert into pereosenka (user,year,month,product_id,deb_sch,naim_tov,edin,kol,sena,summa,gruppa,seria_num,inventar_num,date_pr,deb_sch_2,kre_sch,kre_sch_2,podraz_1,podraz_2,gruppa0_naim,num_period,pereosenka_foiz) values(" +
                                                                                "'" + string_for_otdels + "'," +
                                                                                "'" + (year_global) + "'," +
                                                                                "'" + (month_global) + "'," +
                                                                                "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[5].Value != null ? row.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[6].Value != null ? row.Cells[6].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                                                "" + (row.Cells[15].Value != null ? "'" + DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") + "'" : DateTime.Now.ToString("yyyy-MM-dd")) + ", " +
                                                                                "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                                                "'" + (ot_kogo_1) + "'," +
                                                                                "'" + (ot_kogo_2) + "'," +
                                                                                "'" + (gruppa_naim) + "'," +
                                                                                "'" + (number_pereosenka) + "'," +
                                                                                "'" + ((pereosenka_foiz).ToString()).Replace(',', '.') + "'" +
                                                                                ")";
                                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();



                                                        break;
                                                    }
                                            }


                                        }
                                        sql3.myReader.Close();

                                    }
                                    sql2.myReader.Close();


                                }
                                sql4.myReader.Close();


                                sklad_load();
                                run_alert("Успешно");
                                //MessageBox.Show("Добавлено ", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                
                            }
                        }
                    }
                    else
                    {

                    }
                }
                else
                {
                        foreach (DataGridViewRow row in sklad_dataGridView.Rows)
                        {

                            if (row.Cells[1].Value != null && row.Cells[0].Value != null)
                            {

                                var data_between = " select * from pereosenka_data where '" + (row.Cells[15].Value != null ? DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") : "") + "' between date_start and date_finish ";
                                sql4.myReader = sql4.return_MySqlCommand(data_between).ExecuteReader();
                                while (sql4.myReader.Read())
                                {
                                    number_pereosenka = (sql4.myReader["number"] != DBNull.Value ? sql4.myReader.GetString("number") : "");


                                    var gruppa = "SELECT gruppa FROM gruppa where kod_gruppa='" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'";
                                    sql2.myReader = sql2.return_MySqlCommand(gruppa).ExecuteReader();
                                    while (sql2.myReader.Read())
                                    {
                                        gruppa_naim = (sql2.myReader["gruppa"] != DBNull.Value ? sql2.myReader.GetString("gruppa") : "");

                                        var gruppa0 = " SELECT * FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                        sql3.myReader = sql3.return_MySqlCommand(gruppa0).ExecuteReader();
                                        while (sql3.myReader.Read())
                                        {

                                            switch (number_pereosenka)
                                            {
                                                case "1":
                                                    {
                                                        var number = " SELECT id,first FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                                        sql5.myReader = sql5.return_MySqlCommand(number).ExecuteReader();
                                                        while (sql5.myReader.Read())
                                                        {
                                                            pereosenka_foiz = (sql5.myReader["first"] != DBNull.Value ? sql5.myReader.GetDouble("first") : 1);

                                                            var update_sklad = "update products set " +
                                                                               "sena = '" + refresh_strings_to_mysql(((row.Cells[6].Value != null ? (Convert.ToDouble(row.Cells[6].Value.ToString().Replace(".", ","))) : 0) * pereosenka_foiz).ToString()) + "'," +
                                                                               "summa = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'" +
                                                                               " where product_id = " + row.Cells[1].Value + " and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and komu_1='" + ot_kogo_1 + "' and komu_2='" + ot_kogo_2 + "' ";
                                                            sql.return_MySqlCommand(update_sklad).ExecuteNonQuery();
                                                        }
                                                        sql5.myReader.Close();

                                                        var insert_product = "insert into pereosenka (user,year,month,product_id,deb_sch,naim_tov,edin,kol,sena,summa,gruppa,seria_num,inventar_num,date_pr,deb_sch_2,kre_sch,kre_sch_2,podraz_1,podraz_2,gruppa0_naim,num_period,pereosenka_foiz) values(" +
                                                                                "'" + string_for_otdels + "'," +
                                                                                "'" + (year_global) + "'," +
                                                                                "'" + (month_global) + "'," +
                                                                                "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[5].Value != null ? row.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[6].Value != null ? row.Cells[6].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                                                "" + (row.Cells[15].Value != null ? "'" + DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") + "'" : DateTime.Now.ToString("yyyy-MM-dd")) + ", " +
                                                                                "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                                                "'" + (ot_kogo_1) + "'," +
                                                                                "'" + (ot_kogo_2) + "'," +
                                                                                "'" + (gruppa_naim) + "'," +
                                                                                "'" + (number_pereosenka) + "'," +
                                                                                "'" + ((pereosenka_foiz).ToString()).Replace(',', '.') + "'" +
                                                                                ")";
                                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();


                                                        break;
                                                    }
                                                case "2":
                                                    {
                                                        var number = " SELECT id,first FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                                        sql5.myReader = sql5.return_MySqlCommand(number).ExecuteReader();
                                                        while (sql5.myReader.Read())
                                                        {
                                                            pereosenka_foiz = (sql5.myReader["first"] != DBNull.Value ? sql5.myReader.GetDouble("first") : 1);

                                                            var update_sklad = "update products set " +
                                                                               "sena = '" + refresh_strings_to_mysql(((row.Cells[6].Value != null ? (Convert.ToDouble(row.Cells[6].Value.ToString().Replace(".", ","))) : 0) * pereosenka_foiz).ToString()) + "'," +
                                                                               "summa = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'" +
                                                                               " where product_id = " + row.Cells[1].Value + " and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and komu_1='" + ot_kogo_1 + "' and komu_2='" + ot_kogo_2 + "' ";
                                                            sql.return_MySqlCommand(update_sklad).ExecuteNonQuery();
                                                        }
                                                        sql5.myReader.Close();

                                                        var insert_product = "insert into pereosenka (user,year,month,product_id,deb_sch,naim_tov,edin,kol,sena,summa,gruppa,seria_num,inventar_num,date_pr,deb_sch_2,kre_sch,kre_sch_2,podraz_1,podraz_2,gruppa0_naim,num_period,pereosenka_foiz) values(" +
                                                                                "'" + string_for_otdels + "'," +
                                                                                "'" + (year_global) + "'," +
                                                                                "'" + (month_global) + "'," +
                                                                                "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[5].Value != null ? row.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[6].Value != null ? row.Cells[6].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                                                "" + (row.Cells[15].Value != null ? "'" + DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") + "'" : DateTime.Now.ToString("yyyy-MM-dd")) + ", " +
                                                                                "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                                                "'" + (ot_kogo_1) + "'," +
                                                                                "'" + (ot_kogo_2) + "'," +
                                                                                "'" + (gruppa_naim) + "'," +
                                                                                "'" + (number_pereosenka) + "'," +
                                                                                "'" + ((pereosenka_foiz).ToString()).Replace(',', '.') + "'" +
                                                                                ")";
                                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();


                                                        break;
                                                    }
                                                case "3":
                                                    {
                                                        var number = " SELECT id,first FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                                        sql5.myReader = sql5.return_MySqlCommand(number).ExecuteReader();
                                                        while (sql5.myReader.Read())
                                                        {
                                                            pereosenka_foiz = (sql5.myReader["first"] != DBNull.Value ? sql5.myReader.GetDouble("first") : 1);

                                                            var update_sklad = "update products set " +
                                                                               "sena = '" + refresh_strings_to_mysql(((row.Cells[6].Value != null ? (Convert.ToDouble(row.Cells[6].Value.ToString().Replace(".", ","))) : 0) * pereosenka_foiz).ToString()) + "'," +
                                                                               "summa = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'" +
                                                                               " where product_id = " + row.Cells[1].Value + " and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and komu_1='" + ot_kogo_1 + "' and komu_2='" + ot_kogo_2 + "' ";
                                                            sql.return_MySqlCommand(update_sklad).ExecuteNonQuery();
                                                        }
                                                        sql5.myReader.Close();

                                                        var insert_product = "insert into pereosenka (user,year,month,product_id,deb_sch,naim_tov,edin,kol,sena,summa,gruppa,seria_num,inventar_num,date_pr,deb_sch_2,kre_sch,kre_sch_2,podraz_1,podraz_2,gruppa0_naim,num_period,pereosenka_foiz) values(" +
                                                                                "'" + string_for_otdels + "'," +
                                                                                "'" + (year_global) + "'," +
                                                                                "'" + (month_global) + "'," +
                                                                                "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[5].Value != null ? row.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[6].Value != null ? row.Cells[6].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                                                "" + (row.Cells[15].Value != null ? "'" + DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") + "'" : DateTime.Now.ToString("yyyy-MM-dd")) + ", " +
                                                                                "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                                                "'" + (ot_kogo_1) + "'," +
                                                                                "'" + (ot_kogo_2) + "'," +
                                                                                "'" + (gruppa_naim) + "'," +
                                                                                "'" + (number_pereosenka) + "'," +
                                                                                "'" + ((pereosenka_foiz).ToString()).Replace(',', '.') + "'" +
                                                                                ")";
                                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();


                                                        break;
                                                    }
                                                case "4":
                                                    {
                                                        var number = " SELECT id,first FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                                        sql5.myReader = sql5.return_MySqlCommand(number).ExecuteReader();
                                                        while (sql5.myReader.Read())
                                                        {
                                                            pereosenka_foiz = (sql5.myReader["first"] != DBNull.Value ? sql5.myReader.GetDouble("first") : 1);

                                                            var update_sklad = "update products set " +
                                                                               "sena = '" + refresh_strings_to_mysql(((row.Cells[6].Value != null ? (Convert.ToDouble(row.Cells[6].Value.ToString().Replace(".", ","))) : 0) * pereosenka_foiz).ToString()) + "'," +
                                                                               "summa = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'" +
                                                                               " where product_id = " + row.Cells[1].Value + " and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and komu_1='" + ot_kogo_1 + "' and komu_2='" + ot_kogo_2 + "' ";
                                                            sql.return_MySqlCommand(update_sklad).ExecuteNonQuery();
                                                        }
                                                        sql5.myReader.Close();

                                                        var insert_product = "insert into pereosenka (user,year,month,product_id,deb_sch,naim_tov,edin,kol,sena,summa,gruppa,seria_num,inventar_num,date_pr,deb_sch_2,kre_sch,kre_sch_2,podraz_1,podraz_2,gruppa0_naim,num_period,pereosenka_foiz) values(" +
                                                                                "'" + string_for_otdels + "'," +
                                                                                "'" + (year_global) + "'," +
                                                                                "'" + (month_global) + "'," +
                                                                                "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[5].Value != null ? row.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[6].Value != null ? row.Cells[6].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                                                "" + (row.Cells[15].Value != null ? "'" + DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") + "'" : DateTime.Now.ToString("yyyy-MM-dd")) + ", " +
                                                                                "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                                                "'" + (ot_kogo_1) + "'," +
                                                                                "'" + (ot_kogo_2) + "'," +
                                                                                "'" + (gruppa_naim) + "'," +
                                                                                "'" + (number_pereosenka) + "'," +
                                                                                "'" + ((pereosenka_foiz).ToString()).Replace(',', '.') + "'" +
                                                                                ")";
                                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();


                                                        break;
                                                    }
                                                case "5":
                                                    {
                                                        var number = " SELECT id,first FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                                        sql5.myReader = sql5.return_MySqlCommand(number).ExecuteReader();
                                                        while (sql5.myReader.Read())
                                                        {
                                                            pereosenka_foiz = (sql5.myReader["first"] != DBNull.Value ? sql5.myReader.GetDouble("first") : 1);

                                                            var update_sklad = "update products set " +
                                                                               "sena = '" + refresh_strings_to_mysql(((row.Cells[6].Value != null ? (Convert.ToDouble(row.Cells[6].Value.ToString().Replace(".", ","))) : 0) * pereosenka_foiz).ToString()) + "'," +
                                                                               "summa = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'" +
                                                                               " where product_id = " + row.Cells[1].Value + " and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and komu_1='" + ot_kogo_1 + "' and komu_2='" + ot_kogo_2 + "' ";
                                                            sql.return_MySqlCommand(update_sklad).ExecuteNonQuery();
                                                        }
                                                        sql5.myReader.Close();

                                                        var insert_product = "insert into pereosenka (user,year,month,product_id,deb_sch,naim_tov,edin,kol,sena,summa,gruppa,seria_num,inventar_num,date_pr,deb_sch_2,kre_sch,kre_sch_2,podraz_1,podraz_2,gruppa0_naim,num_period,pereosenka_foiz) values(" +
                                                                                 "'" + string_for_otdels + "'," +
                                                                                 "'" + (year_global) + "'," +
                                                                                 "'" + (month_global) + "'," +
                                                                                 "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                                                   "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                                                 "'" + refresh_strings_to_mysql(row.Cells[5].Value != null ? row.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                 "'" + refresh_strings_to_mysql(row.Cells[6].Value != null ? row.Cells[6].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                 "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                 "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                                                   "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                                                   "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                                                 "" + (row.Cells[15].Value != null ? "'" + DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") + "'" : DateTime.Now.ToString("yyyy-MM-dd")) + ", " +
                                                                                 "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                                                 "'" + (ot_kogo_1) + "'," +
                                                                                 "'" + (ot_kogo_2) + "'," +
                                                                                 "'" + (gruppa_naim) + "'," +
                                                                                 "'" + (number_pereosenka) + "'," +
                                                                                 "'" + ((pereosenka_foiz).ToString()).Replace(',', '.') + "'" +
                                                                                 ")";
                                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();


                                                        break;
                                                    }
                                                case "6":
                                                    {
                                                        var number = " SELECT id,first FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                                        sql5.myReader = sql5.return_MySqlCommand(number).ExecuteReader();
                                                        while (sql5.myReader.Read())
                                                        {
                                                            pereosenka_foiz = (sql5.myReader["first"] != DBNull.Value ? sql5.myReader.GetDouble("first") : 1);

                                                            var update_sklad = "update products set " +
                                                                               "sena = '" + refresh_strings_to_mysql(((row.Cells[6].Value != null ? (Convert.ToDouble(row.Cells[6].Value.ToString().Replace(".", ","))) : 0) * pereosenka_foiz).ToString()) + "'," +
                                                                               "summa = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'" +
                                                                               " where product_id = " + row.Cells[1].Value + " and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and komu_1='" + ot_kogo_1 + "' and komu_2='" + ot_kogo_2 + "' ";
                                                            sql.return_MySqlCommand(update_sklad).ExecuteNonQuery();
                                                        }
                                                        sql5.myReader.Close();

                                                        var insert_product = "insert into pereosenka (user,year,month,product_id,deb_sch,naim_tov,edin,kol,sena,summa,gruppa,seria_num,inventar_num,date_pr,deb_sch_2,kre_sch,kre_sch_2,podraz_1,podraz_2,gruppa0_naim,num_period,pereosenka_foiz) values(" +
                                                                                "'" + string_for_otdels + "'," +
                                                                                "'" + (year_global) + "'," +
                                                                                "'" + (month_global) + "'," +
                                                                                "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[5].Value != null ? row.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[6].Value != null ? row.Cells[6].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                                                "" + (row.Cells[15].Value != null ? "'" + DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") + "'" : DateTime.Now.ToString("yyyy-MM-dd")) + ", " +
                                                                                "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                                                "'" + (ot_kogo_1) + "'," +
                                                                                "'" + (ot_kogo_2) + "'," +
                                                                                "'" + (gruppa_naim) + "'," +
                                                                                "'" + (number_pereosenka) + "'," +
                                                                                "'" + ((pereosenka_foiz).ToString()).Replace(',', '.') + "'" +
                                                                                ")";
                                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();


                                                        break;
                                                    }
                                                case "7":
                                                    {
                                                        var number = " SELECT id,first FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                                        sql5.myReader = sql5.return_MySqlCommand(number).ExecuteReader();
                                                        while (sql5.myReader.Read())
                                                        {
                                                            pereosenka_foiz = (sql5.myReader["first"] != DBNull.Value ? sql5.myReader.GetDouble("first") : 1);

                                                            var update_sklad = "update products set " +
                                                                               "sena = '" + refresh_strings_to_mysql(((row.Cells[6].Value != null ? (Convert.ToDouble(row.Cells[6].Value.ToString().Replace(".", ","))) : 0) * pereosenka_foiz).ToString()) + "'," +
                                                                               "summa = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'" +
                                                                               " where product_id = " + row.Cells[1].Value + " and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and komu_1='" + ot_kogo_1 + "' and komu_2='" + ot_kogo_2 + "' ";
                                                            sql.return_MySqlCommand(update_sklad).ExecuteNonQuery();
                                                        }
                                                        sql5.myReader.Close();

                                                        var insert_product = "insert into pereosenka (user,year,month,product_id,deb_sch,naim_tov,edin,kol,sena,summa,gruppa,seria_num,inventar_num,date_pr,deb_sch_2,kre_sch,kre_sch_2,podraz_1,podraz_2,gruppa0_naim,num_period,pereosenka_foiz) values(" +
                                                                                 "'" + string_for_otdels + "'," +
                                                                                 "'" + (year_global) + "'," +
                                                                                 "'" + (month_global) + "'," +
                                                                                 "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                                                   "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                                                 "'" + refresh_strings_to_mysql(row.Cells[5].Value != null ? row.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                 "'" + refresh_strings_to_mysql(row.Cells[6].Value != null ? row.Cells[6].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                 "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                 "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                                                   "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                                                   "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                                                 "" + (row.Cells[15].Value != null ? "'" + DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") + "'" : DateTime.Now.ToString("yyyy-MM-dd")) + ", " +
                                                                                 "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                                                 "'" + (ot_kogo_1) + "'," +
                                                                                 "'" + (ot_kogo_2) + "'," +
                                                                                 "'" + (gruppa_naim) + "'," +
                                                                                 "'" + (number_pereosenka) + "'," +
                                                                                 "'" + ((pereosenka_foiz).ToString()).Replace(',', '.') + "'" +
                                                                                 ")";
                                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();


                                                        break;
                                                    }
                                                case "8":
                                                    {
                                                        var number = " SELECT id,first FROM gruppa0 where gruppa='" + gruppa_naim + "' ";
                                                        sql5.myReader = sql5.return_MySqlCommand(number).ExecuteReader();
                                                        while (sql5.myReader.Read())
                                                        {
                                                            pereosenka_foiz = (sql5.myReader["first"] != DBNull.Value ? sql5.myReader.GetDouble("first") : 1);

                                                            var update_sklad = "update products set " +
                                                                               "sena = '" + refresh_strings_to_mysql(((row.Cells[6].Value != null ? (Convert.ToDouble(row.Cells[6].Value.ToString().Replace(".", ","))) : 0) * pereosenka_foiz).ToString()) + "'," +
                                                                               "summa = '" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(",", ".") : "0") + "'" +
                                                                               " where product_id = " + row.Cells[1].Value + " and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and komu_1='" + ot_kogo_1 + "' and komu_2='" + ot_kogo_2 + "' ";
                                                            sql.return_MySqlCommand(update_sklad).ExecuteNonQuery();
                                                        }
                                                        sql5.myReader.Close();

                                                        var insert_product = "insert into pereosenka (user,year,month,product_id,deb_sch,naim_tov,edin,kol,sena,summa,gruppa,seria_num,inventar_num,date_pr,deb_sch_2,kre_sch,kre_sch_2,podraz_1,podraz_2,gruppa0_naim,num_period,pereosenka_foiz) values(" +
                                                                                "'" + string_for_otdels + "'," +
                                                                                "'" + (year_global) + "'," +
                                                                                "'" + (month_global) + "'," +
                                                                                "'" + (row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "") + "'," +
                                                                                 "'" + (row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[4].Value != null ? row.Cells[4].Value.ToString() : "") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[5].Value != null ? row.Cells[5].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[6].Value != null ? row.Cells[6].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + refresh_strings_to_mysql(row.Cells[7].Value != null ? row.Cells[7].Value.ToString().Replace(',', '.') : "0") + "'," +
                                                                                "'" + (row.Cells[11].Value != null ? row.Cells[11].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[13].Value != null ? row.Cells[13].Value.ToString() : "") + "'," +
                                                                                  "'" + (row.Cells[14].Value != null ? row.Cells[14].Value.ToString() : "") + "'," +
                                                                                "" + (row.Cells[15].Value != null ? "'" + DateTime.Parse(row.Cells[15].Value.ToString()).ToString("yyyy-MM-dd") + "'" : DateTime.Now.ToString("yyyy-MM-dd")) + ", " +
                                                                                "'" + (row.Cells[16].Value != null ? row.Cells[16].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[17].Value != null ? row.Cells[17].Value.ToString() : "") + "'," +
                                                                                "'" + (row.Cells[18].Value != null ? row.Cells[18].Value.ToString() : "") + "'," +
                                                                                "'" + (ot_kogo_1) + "'," +
                                                                                "'" + (ot_kogo_2) + "'," +
                                                                                "'" + (gruppa_naim) + "'," +
                                                                                "'" + (number_pereosenka) + "'," +
                                                                                "'" + ((pereosenka_foiz).ToString()).Replace(',', '.') + "'" +
                                                                                ")";
                                                        sql.return_MySqlCommand(insert_product).ExecuteNonQuery();



                                                        break;
                                                    }
                                            }


                                        }
                                        sql3.myReader.Close();

                                    }
                                    sql2.myReader.Close();


                                }
                                sql4.myReader.Close();


                                sklad_load();
                                MessageBox.Show("Добавлено ", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                            
                            }
                        }
                }
                
            }
            else
            {

            }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("prixod_dataGridView_UserDeletingRow " + ex.Message);
            //}
        }

        private void search_comboBox_TextChanged(object sender, EventArgs e)
        {
            this.sklad_dataGridView.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.sklad_dataGridView_CellValueChanged);

            sklad_dataGridView.Rows.Clear();


            double pri_kol = 0;
            double ras_kol = 0;
            double vnut_ras_kol = 0;

            var products = " select t.id,t.vid_doc,t.kod_doc,t.product_id,t.gruppa,t.naim_tov,t.edin,t.inventar_num,t.seria_num,sum(t.kol) as kol,t.sena,sum(t.summa) as summa,t.deb_sch,t.deb_sch_2,t.kre_sch,t.kre_sch_2,t.provodka_iznos,t.provodka_iznos_2,t.summa_iznos,t.date_pr" +
                            " from(SELECT id, vid_doc, kod_doc, product_id, gruppa, naim_tov, edin, inventar_num, seria_num, sum(kol) as kol, sena, sum(summa) as summa, deb_sch, deb_sch_2, kre_sch, kre_sch_2, provodka_iznos, provodka_iznos_2, summa_iznos, date_pr FROM products" +
                            " where vid_doc = '1' and user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' and kol > 0 and komu_1 = '" + ot_kogo_1 + "' and komu_2 = '" + ot_kogo_2 + "' and naim_tov like '%" + search_comboBox.Text + "%' group by product_id" +
                            " union all" +
                            " SELECT id, vid_doc, kod_doc, product_id, gruppa, naim_tov, edin, inventar_num, seria_num, sum(kol) as kol, sena, summa, deb_sch, deb_sch_2, kre_sch, kre_sch_2, provodka_iznos, provodka_iznos_2, summa_iznos, date_pr FROM products" +
                            " where vid_doc = '3' and user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' and kol > 0 and komu_1 = '" + ot_kogo_1 + "' and komu_2 = '" + ot_kogo_2 + "' and naim_tov like '%" + search_comboBox.Text + "%' group by product_id) as t group by t.product_id ";
            sql.myReader = sql.return_MySqlCommand(products).ExecuteReader();

            while (sql.myReader.Read())
            {
                int index = sklad_dataGridView.Rows.Add();

                sklad_dataGridView.Rows[index].Cells[0].Value = (sql.myReader["id"] != DBNull.Value ? sql.myReader.GetString("id") : "");
                sklad_dataGridView.Rows[index].Cells[1].Value = (sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetString("product_id") : "");
                sklad_dataGridView.Rows[index].Cells[2].Value = (sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch") : "");
                sklad_dataGridView.Rows[index].Cells[3].Value = (sql.myReader["naim_tov"] != DBNull.Value ? sql.myReader.GetString("naim_tov") : "");
                sklad_dataGridView.Rows[index].Cells[4].Value = (sql.myReader["edin"] != DBNull.Value ? sql.myReader.GetString("edin") : "");


                pri_kol = (sql.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("kol").Replace(".", ","))) : 0);

                var products_pri = " SELECT sum(kol) as kol FROM products where vid_doc='2' and product_id='" + (sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetString("product_id") : "") + "' and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and ot_kogo='" + ot_kogo_1 + "' and ot_kogo_2='" + ot_kogo_2 + "' and kol > 0 ";

                sql2.myReader = sql2.return_MySqlCommand(products_pri).ExecuteReader();

                while (sql2.myReader.Read())
                {
                    ras_kol = (sql2.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql2.myReader.GetString("kol").Replace(".", ","))) : 0);
                }
                sql2.myReader.Close();

                var products_vnut = " SELECT id,sum(kol) as kol FROM products where vid_doc = '3' and product_id='" + (sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetString("product_id") : "") + "' and user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' and kol > 0 and ot_kogo = '" + ot_kogo_1 + "' and ot_kogo_2 = '" + ot_kogo_2 + "' group by product_id";

                sql2.myReader = sql2.return_MySqlCommand(products_vnut).ExecuteReader();

                while (sql2.myReader.Read())
                {
                    vnut_ras_kol = (sql2.myReader["kol"] != DBNull.Value ? sql2.myReader.GetDouble("kol") : 0);
                }
                sql2.myReader.Close();


                string kols = (pri_kol - ras_kol - vnut_ras_kol).ToString();//sql3.myReader["kol"] != DBNull.Value ? sql3.myReader.GetString("kol") : "";
                //sklad_dataGridView.Rows[index].Cells[5].Style.BackColor = Color.Yellow;
                if (kols.Length <= 3)
                {
                    sklad_dataGridView.Rows[index].Cells[5].Value = string.Format("{0:#0.00}", (pri_kol - ras_kol - vnut_ras_kol));
                }
                if (kols.Length > 3)
                {
                    sklad_dataGridView.Rows[index].Cells[5].Value = string.Format("{0:#,###.00}", (pri_kol - ras_kol - vnut_ras_kol));
                }

                string sena = sql.myReader["sena"] != DBNull.Value ? sql.myReader.GetString("sena") : "";

                if (sena.Length <= 3)
                {
                    sklad_dataGridView.Rows[index].Cells[6].Value = string.Format("{0:#0.00}", (sql.myReader["sena"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("sena").Replace(".", ","))) : 0));
                }
                if (sena.Length > 3)
                {
                    sklad_dataGridView.Rows[index].Cells[6].Value = string.Format("{0:#,###.00}", (sql.myReader["sena"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("sena").Replace(".", ","))) : 0));
                }

                string summa = ((pri_kol - ras_kol - vnut_ras_kol) * (sql.myReader["sena"] != DBNull.Value ? sql.myReader.GetDouble("sena") : 0)).ToString(); //sql3.myReader["summa"] != DBNull.Value ? sql3.myReader.GetString("summa") : "";

                if (summa.Length <= 3)
                {
                    sklad_dataGridView.Rows[index].Cells[7].Value = string.Format("{0:#0.00}", ((pri_kol - ras_kol - vnut_ras_kol) * (sql.myReader["sena"] != DBNull.Value ? sql.myReader.GetDouble("sena") : 0)));
                }
                if (summa.Length > 3)
                {
                    sklad_dataGridView.Rows[index].Cells[7].Value = string.Format("{0:#,###.00}", ((pri_kol - ras_kol - vnut_ras_kol) * (sql.myReader["sena"] != DBNull.Value ? sql.myReader.GetDouble("sena") : 0)));
                }

                string sum_iznos = sql.myReader["summa_iznos"] != DBNull.Value ? sql.myReader.GetString("summa_iznos") : "";

                if (sum_iznos.Length <= 3)
                {
                    sklad_dataGridView.Rows[index].Cells[8].Value = string.Format("{0:#0.00}", (sql.myReader["summa_iznos"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa_iznos").Replace(".", ","))) : 0));
                }
                if (sum_iznos.Length > 3)
                {
                    sklad_dataGridView.Rows[index].Cells[8].Value = string.Format("{0:#,###.00}", (sql.myReader["summa_iznos"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa_iznos").Replace(".", ","))) : 0));
                }


                //deb_sch,deb_sch_2,kre_sch,kre_sch_2,provodka_iznos,provodka_iznos_2,summa_iznos,date_pr,id_products
                sklad_dataGridView.Rows[index].Cells[9].Value = ("0");
                sklad_dataGridView.Rows[index].Cells[9].Style.BackColor = Color.GreenYellow;
                sklad_dataGridView.Rows[index].Cells[10].Value = ("0");
                sklad_dataGridView.Rows[index].Cells[11].Value = (sql.myReader["gruppa"] != DBNull.Value ? sql.myReader.GetString("gruppa") : "");
                sklad_dataGridView.Rows[index].Cells[13].Value = (sql.myReader["seria_num"] != DBNull.Value ? sql.myReader.GetString("seria_num") : "");
                sklad_dataGridView.Rows[index].Cells[14].Value = (sql.myReader["inventar_num"] != DBNull.Value ? sql.myReader.GetString("inventar_num") : "");
                sklad_dataGridView.Rows[index].Cells[15].Value = (sql.myReader["date_pr"] != DBNull.Value ? (DateTime.Parse(sql.myReader.GetString("date_pr")).ToString("dd.MM.yyyy")) : null);
                sklad_dataGridView.Rows[index].Cells[16].Value = (sql.myReader["deb_sch_2"] != DBNull.Value ? sql.myReader.GetString("deb_sch_2") : "");
                sklad_dataGridView.Rows[index].Cells[17].Value = (sql.myReader["kre_sch"] != DBNull.Value ? sql.myReader.GetString("kre_sch") : "");
                sklad_dataGridView.Rows[index].Cells[18].Value = (sql.myReader["kre_sch_2"] != DBNull.Value ? sql.myReader.GetString("kre_sch_2") : "");


                //deb_sch_2,kre_sch,kre_sch_2
            }
            sql.myReader.Close();


            this.sklad_dataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.sklad_dataGridView_CellValueChanged);
        }
    }
}
