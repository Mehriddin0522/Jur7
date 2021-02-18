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
    public partial class Ostatok : Form
    {
        Connect sql = new Connect();
        Connect sql1 = new Connect();
        Connect sql2 = new Connect();

        public string string_for_otdels;
        public string month_global;
        public string year_global;
        public Ostatok(string string_for_otdels, string year_global, string month_global)
        {
            InitializeComponent();

            sql.Connection();
            sql1.Connection();
            sql2.Connection();

            this.string_for_otdels = string_for_otdels;
            this.year_global = year_global;
            this.month_global = month_global;

            run_treeview();
        }

        TreeNode tovar = new TreeNode("Общий");

        public void run_treeview()
        {
            try
            {


                var select1 = "SELECT podraz_naim FROM podraz_jur7 where podraz_naim is not null group by podraz_naim";
                sql.myReader = sql.return_MySqlCommand(select1).ExecuteReader();
                while (sql.myReader.Read())
                {
                    // tovar = new TreeNode(sql.myReader.GetString("otdel_group"));
                    // treeView1.Nodes[0].Nodes.Add(tovar);

                    treeView.Nodes.Add(new TreeNode(sql.myReader.GetString("podraz_naim")));


                    var select = "SELECT fio FROM podraz_jur7 where fio is not null and podraz_naim='" + (sql.myReader.GetString("podraz_naim")) + "' group by fio";
                    sql2.myReader = sql2.return_MySqlCommand(select).ExecuteReader();
                    while (sql2.myReader.Read())
                    {
                        TreeNode tovar_type = new TreeNode(sql2.myReader.GetString("fio"));
                        treeView.Nodes[treeView.Nodes.Count - 1].Nodes.Add(tovar_type);

                        //tovar.Nodes[tovar.Nodes.Count - 1].ImageKey = "2";
                        //.Nodes[tovar.Nodes.Count - 1].SelectedImageKey = "2";
                    }
                    sql2.myReader.Close();

                }
                sql.myReader.Close();

                treeView.Nodes.Add(tovar);


                //treeView.Nodes[0].Expand();
            }
            catch (Exception ex)
            {
                MessageBox.Show("run_treeview " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private bool HasCheckedNode(TreeNode node)
        {
            return node.Nodes.Cast<TreeNode>().Any(n => n.Checked);
        }

        private void SelectParents(TreeNode node, Boolean isChecked)
        {

            try
            {
                var parent = node.Parent;

                if (parent == null)
                    return;

                if (!isChecked && HasCheckedNode(parent))
                    return;

                parent.Checked = isChecked;
                SelectParents(parent, isChecked);
            }
            catch (Exception ex)
            {
                MessageBox.Show("SelectParents " + ex.Message);
            }
        }

        public TreeNode previousSelectedNode = null;
        private void treeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            try
            {
                if (treeView.SelectedNode.Parent != null)
                {
                    sklad_dataGridView.Rows.Clear();

                    this.sklad_dataGridView.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.sklad_dataGridView_CellValueChanged);


                    sklad_dataGridView.Rows.Clear();




                    var products = " select t.id,t.vid_doc,t.kod_doc,t.product_id,t.gruppa,t.naim_tov,t.edin,t.inventar_num,t.seria_num,sum(t.kol) as kol,t.sena,sum(t.summa) as summa,t.deb_sch,t.deb_sch_2,t.kre_sch,t.kre_sch_2,t.provodka_iznos,t.provodka_iznos_2,t.summa_iznos,t.date_pr" +
                                    " from(SELECT id, vid_doc, kod_doc, product_id, gruppa, naim_tov, edin, inventar_num, seria_num, sum(kol) as kol, sena, sum(summa) as summa, deb_sch, deb_sch_2, kre_sch, kre_sch_2, provodka_iznos, provodka_iznos_2, summa_iznos, date_pr FROM products_jur7" +
                                    " where vid_doc = '1' and user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' and kol > 0 and komu_1 = '" + treeView.SelectedNode.Parent.Text.ToString() + "' and komu_2 = '" + treeView.SelectedNode.Text.ToString() + "' group by product_id" +
                                    " union all" +
                                    " SELECT id, vid_doc, kod_doc, product_id, gruppa, naim_tov, edin, inventar_num, seria_num, sum(kol) as kol, sena, summa, deb_sch, deb_sch_2, kre_sch, kre_sch_2, provodka_iznos, provodka_iznos_2, summa_iznos, date_pr FROM products_jur7" +
                                    " where vid_doc = '3' and user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' and kol > 0 and komu_1 = '" + treeView.SelectedNode.Parent.Text.ToString() + "' and komu_2 = '" + treeView.SelectedNode.Text.ToString() + "' group by product_id) as t where t.kol > 0 group by t.product_id " +
                                    " union all " +
                                    " select id, '' as vid_doc,'' as kod_doc,product_id,gruppa, naim_tov, edin, inventar_num, seria_num,kol,sena,summa,deb_sch, deb_sch_2, kre_sch, kre_sch_2, '' as provodka_iznos, '' as provodka_iznos_2,summa_iznos, " +
                                    " data_pr from saldo_jur7 where kol > 0 and user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' and podraz_1 = '" + treeView.SelectedNode.Parent.Text.ToString() + "' and podraz_2 = '" + treeView.SelectedNode.Text.ToString() + "' ";

                    sql.myReader = sql.return_MySqlCommand(products).ExecuteReader();

                    while (sql.myReader.Read())
                    {
                        double pri_kol = 0;
                        double ras_kol = 0;
                        double vnut_ras_kol = 0;

                        int index = sklad_dataGridView.Rows.Add();

                        sklad_dataGridView.Rows[index].Cells[0].Value = (sql.myReader["id"] != DBNull.Value ? sql.myReader.GetString("id") : "");
                        sklad_dataGridView.Rows[index].Cells[1].Value = (sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetString("product_id") : "");

                        string gruppa_name = "";
                        var gruppa_naim = " SELECT naim FROM gruppa_jur7 where kod_gruppa = '" + (sql.myReader["gruppa"] != DBNull.Value ? sql.myReader.GetString("gruppa") : "") + "' group by naim ";

                        sql2.myReader = sql2.return_MySqlCommand(gruppa_naim).ExecuteReader();

                        while (sql2.myReader.Read())
                        {
                            gruppa_name = (sql2.myReader["naim"] != DBNull.Value ? sql2.myReader.GetString("naim") : "");
                        }
                        sql2.myReader.Close();

                        sklad_dataGridView.Rows[index].Cells[2].Value = gruppa_name;
                        sklad_dataGridView.Rows[index].Cells[3].Value = (sql.myReader["edin"] != DBNull.Value ? sql.myReader.GetString("edin") : "");
                        sklad_dataGridView.Rows[index].Cells[4].Value = (sql.myReader["date_pr"] != DBNull.Value ? (DateTime.Parse(sql.myReader.GetString("date_pr")).ToString("dd.MM.yyyy")) : null);
                        sklad_dataGridView.Rows[index].Cells[5].Value = (sql.myReader["naim_tov"] != DBNull.Value ? sql.myReader.GetString("naim_tov") : "");





                        pri_kol = (sql.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("kol").Replace(".", ","))) : 0);

                        var products_pri = " SELECT sum(kol) as kol FROM products_jur7 where vid_doc='2' and product_id='" + (sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetString("product_id") : "") + "' and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and ot_kogo='" + treeView.SelectedNode.Parent.Text.ToString() + "' and ot_kogo_2='" + treeView.SelectedNode.Text.ToString() + "' and kol > 0 ";

                        sql2.myReader = sql2.return_MySqlCommand(products_pri).ExecuteReader();

                        while (sql2.myReader.Read())
                        {
                            ras_kol = (sql2.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql2.myReader.GetString("kol").Replace(".", ","))) : 0);
                        }
                        sql2.myReader.Close();

                        var products_vnut = " SELECT sum(kol) as kol FROM products_jur7 where vid_doc = '3' and product_id='" + (sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetString("product_id") : "") + "' and user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' and kol > 0 and ot_kogo = '" + treeView.SelectedNode.Parent.Text.ToString() + "' and ot_kogo_2 = '" + treeView.SelectedNode.Text.ToString() + "' group by product_id";

                        sql2.myReader = sql2.return_MySqlCommand(products_vnut).ExecuteReader();

                        while (sql2.myReader.Read())
                        {
                            vnut_ras_kol = (sql2.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql2.myReader.GetString("kol").Replace(".", ","))) : 0);
                        }
                        sql2.myReader.Close();

                        string kols = (pri_kol - ras_kol - vnut_ras_kol).ToString();//sql3.myReader["kol"] != DBNull.Value ? sql3.myReader.GetString("kol") : "";
                                                                                    //sklad_dataGridView.Rows[index].Cells[5].Style.BackColor = Color.Yellow;
                        if (kols.Length <= 3)
                        {
                            sklad_dataGridView.Rows[index].Cells[6].Value = string.Format("{0:#0.00}", (pri_kol - ras_kol - vnut_ras_kol));
                        }
                        if (kols.Length > 3)
                        {
                            sklad_dataGridView.Rows[index].Cells[6].Value = string.Format("{0:#,###.00}", (pri_kol - ras_kol - vnut_ras_kol));
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

                        sklad_dataGridView.Rows[index].Cells[8].Value = (sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch") : "");


                        string sum_iznos = sql.myReader["summa_iznos"] != DBNull.Value ? sql.myReader.GetString("summa_iznos") : "";

                        if (sum_iznos.Length <= 3)
                        {
                            sklad_dataGridView.Rows[index].Cells[9].Value = string.Format("{0:#0.00}", (sql.myReader["summa_iznos"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa_iznos").Replace(".", ","))) : 0));
                        }
                        if (sum_iznos.Length > 3)
                        {
                            sklad_dataGridView.Rows[index].Cells[9].Value = string.Format("{0:#,###.00}", (sql.myReader["summa_iznos"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa_iznos").Replace(".", ","))) : 0));
                        }


                    }
                    sql.myReader.Close();

                    label_update_prixod();

                    this.sklad_dataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.sklad_dataGridView_CellValueChanged);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("run_treeview " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }


        public void label_update_prixod()
        {
            double kol = 0;
            double summa = 0;


            foreach (DataGridViewRow row in sklad_dataGridView.Rows)
            {
                kol = kol + (row.Cells[6].Value != null ? Double.Parse(row.Cells[6].Value.ToString()) : 0);

                summa = summa + (row.Cells[7].Value != null ? Double.Parse(row.Cells[7].Value.ToString()) : 0);

            }
            if (kol.ToString().Length <= 3)
            {
                label2.Text = string.Format("{0:#0.00}", kol);
            }
            if (kol.ToString().Length > 3)
            {
                label2.Text = string.Format("{0:#0,000.00}", kol);
            }

            if (summa.ToString().Length <= 3)
            {
                label4.Text = string.Format("{0:#0.00}", summa);
            }
            if (summa.ToString().Length > 3)
            {
                label4.Text = string.Format("{0:#0,000.00}", summa);
            }

        }
        private void sklad_dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Ostatok_Load(object sender, EventArgs e)
        {
            this.sklad_dataGridView.RowsDefaultCellStyle.BackColor = Color.White;
            this.sklad_dataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(233, 233, 234);



            sklad_dataGridView.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            sklad_dataGridView.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            sklad_dataGridView.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            sklad_dataGridView.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            sklad_dataGridView.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            sklad_dataGridView.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            sklad_dataGridView.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            sklad_dataGridView.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {

                sklad_dataGridView.Rows.Clear();

                this.sklad_dataGridView.CellValueChanged -= new System.Windows.Forms.DataGridViewCellEventHandler(this.sklad_dataGridView_CellValueChanged);

                var products = " select t.id,t.vid_doc,t.kod_doc,t.product_id,t.gruppa,t.naim_tov,t.edin,t.inventar_num,t.seria_num,sum(t.kol) as kol,t.sena,sum(t.summa) as summa,t.deb_sch,t.deb_sch_2,t.kre_sch,t.kre_sch_2,t.provodka_iznos,t.provodka_iznos_2,t.summa_iznos,t.date_pr" +
                                 " from(SELECT id, vid_doc, kod_doc, product_id, gruppa, naim_tov, edin, inventar_num, seria_num, sum(kol) as kol, sena, sum(summa) as summa, deb_sch, deb_sch_2, kre_sch, kre_sch_2, provodka_iznos, provodka_iznos_2, summa_iznos, date_pr FROM products_jur7" +
                                 " where vid_doc = '1' and user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' and kol > 0 and naim_tov like '%" + textBox1.Text + "%' group by product_id" +
                                 " union all" +
                                 " SELECT id, vid_doc, kod_doc, product_id, gruppa, naim_tov, edin, inventar_num, seria_num, sum(kol) as kol, sena, summa, deb_sch, deb_sch_2, kre_sch, kre_sch_2, provodka_iznos, provodka_iznos_2, summa_iznos, date_pr FROM products_jur7" +
                                 " where vid_doc = '3' and user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' and kol > 0 and naim_tov like '%" + textBox1.Text + "%' group by product_id) as t group by t.product_id " +
                                 " union all " +
                                 " select id, '' as vid_doc,'' as kod_doc,product_id,gruppa, naim_tov, edin, inventar_num, seria_num,kol,sena,summa,deb_sch, deb_sch_2, kre_sch, kre_sch_2, '' as provodka_iznos, '' as provodka_iznos_2,summa_iznos, " +
                                 " data_pr from saldo_jur7 where user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' and naim_tov like '%" + textBox1.Text + "%' ";

                sql.myReader = sql.return_MySqlCommand(products).ExecuteReader();

                while (sql.myReader.Read())
                {
                    double pri_kol = 0;
                    double ras_kol = 0;
                    double vnut_ras_kol = 0;

                    int index = sklad_dataGridView.Rows.Add();

                    sklad_dataGridView.Rows[index].Cells[0].Value = (sql.myReader["id"] != DBNull.Value ? sql.myReader.GetString("id") : "");
                    sklad_dataGridView.Rows[index].Cells[1].Value = (sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetString("product_id") : "");

                    string gruppa_name = "";
                    var gruppa_naim = " SELECT naim FROM gruppa_jur7 where kod_gruppa = '" + (sql.myReader["gruppa"] != DBNull.Value ? sql.myReader.GetString("gruppa") : "") + "' group by naim ";

                    sql2.myReader = sql2.return_MySqlCommand(gruppa_naim).ExecuteReader();

                    while (sql2.myReader.Read())
                    {
                        gruppa_name = (sql2.myReader["naim"] != DBNull.Value ? sql2.myReader.GetString("naim") : "");
                    }
                    sql2.myReader.Close();

                    sklad_dataGridView.Rows[index].Cells[2].Value = gruppa_name;
                    sklad_dataGridView.Rows[index].Cells[3].Value = (sql.myReader["edin"] != DBNull.Value ? sql.myReader.GetString("edin") : "");
                    sklad_dataGridView.Rows[index].Cells[4].Value = (sql.myReader["date_pr"] != DBNull.Value ? (DateTime.Parse(sql.myReader.GetString("date_pr")).ToString("dd.MM.yyyy")) : null);
                    sklad_dataGridView.Rows[index].Cells[5].Value = (sql.myReader["naim_tov"] != DBNull.Value ? sql.myReader.GetString("naim_tov") : "");


                    pri_kol = (sql.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("kol").Replace(".", ","))) : 0);

                    var products_pri = " SELECT sum(kol) as kol FROM products_jur7 where vid_doc='2' and product_id='" + (sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetString("product_id") : "") + "' and user='" + string_for_otdels + "' and year='" + year_global + "' and month='" + month_global + "' and naim_tov like '%" + textBox1.Text + "%' and kol > 0 ";

                    sql2.myReader = sql2.return_MySqlCommand(products_pri).ExecuteReader();

                    while (sql2.myReader.Read())
                    {
                        ras_kol = (sql2.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql2.myReader.GetString("kol").Replace(".", ","))) : 0);
                    }
                    sql2.myReader.Close();

                    var products_vnut = " SELECT sum(kol) as kol FROM products_jur7 where vid_doc = '3' and product_id='" + (sql.myReader["product_id"] != DBNull.Value ? sql.myReader.GetString("product_id") : "") + "' and user = '" + string_for_otdels + "' and year = '" + year_global + "' and month = '" + month_global + "' and kol > 0 and naim_tov like '%" + textBox1.Text + "%' group by product_id";

                    sql2.myReader = sql2.return_MySqlCommand(products_vnut).ExecuteReader();

                    while (sql2.myReader.Read())
                    {
                        vnut_ras_kol = (sql2.myReader["kol"] != DBNull.Value ? (Convert.ToDouble(sql2.myReader.GetString("kol").Replace(".", ","))) : 0);
                    }
                    sql2.myReader.Close();

                    string kols = (pri_kol - ras_kol - vnut_ras_kol).ToString();//sql3.myReader["kol"] != DBNull.Value ? sql3.myReader.GetString("kol") : "";
                                                                                //sklad_dataGridView.Rows[index].Cells[5].Style.BackColor = Color.Yellow;
                    if (kols.Length <= 3)
                    {
                        sklad_dataGridView.Rows[index].Cells[6].Value = string.Format("{0:#0.00}", (pri_kol - ras_kol - vnut_ras_kol));
                    }
                    if (kols.Length > 3)
                    {
                        sklad_dataGridView.Rows[index].Cells[6].Value = string.Format("{0:#,###.00}", (pri_kol - ras_kol - vnut_ras_kol));
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

                    sklad_dataGridView.Rows[index].Cells[8].Value = (sql.myReader["deb_sch"] != DBNull.Value ? sql.myReader.GetString("deb_sch") : "");


                    string sum_iznos = sql.myReader["summa_iznos"] != DBNull.Value ? sql.myReader.GetString("summa_iznos") : "";

                    if (sum_iznos.Length <= 3)
                    {
                        sklad_dataGridView.Rows[index].Cells[9].Value = string.Format("{0:#0.00}", (sql.myReader["summa_iznos"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa_iznos").Replace(".", ","))) : 0));
                    }
                    if (sum_iznos.Length > 3)
                    {
                        sklad_dataGridView.Rows[index].Cells[9].Value = string.Format("{0:#,###.00}", (sql.myReader["summa_iznos"] != DBNull.Value ? (Convert.ToDouble(sql.myReader.GetString("summa_iznos").Replace(".", ","))) : 0));
                    }


                }
                sql.myReader.Close();


                label_update_prixod();

                this.sklad_dataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.sklad_dataGridView_CellValueChanged);
            }
            catch (Exception ex)
            {
                sql.myReader.Close();
                MessageBox.Show("poisk_Click" + ex.Message);
            }
        }

        private void sklad_dataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.Exception.Message == "DataGridViewComboBoxCell value is not valid.")
            {
                object value = sklad_dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (!((DataGridViewComboBoxColumn)sklad_dataGridView.Columns[e.ColumnIndex]).Items.Contains(value))
                {
                    ((DataGridViewComboBoxColumn)sklad_dataGridView.Columns[e.ColumnIndex]).Items.Add(value);
                    e.ThrowException = false;
                }
            }
        }
    }
}
