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
    public partial class qrcode : Form
    {
        Connect sql = new Connect();
        Connect sql1 = new Connect();
        Connect sql2 = new Connect();

        public string string_for_otdels;
        public string month_global;
        public string year_global;
        public qrcode(string string_for_otdels, string year_global, string month_global)
        {
            InitializeComponent();

            sql.Connection();
            sql1.Connection();
            sql2.Connection();

            this.string_for_otdels = string_for_otdels;
            this.month_global = month_global;
            this.year_global = year_global;
        }

        private void qrcode_Load(object sender, EventArgs e)
        {
            this.vnut_per_dataGridView.RowsDefaultCellStyle.BackColor = Color.White;
            this.vnut_per_dataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(233, 233, 234);

            vnut_per_dataGridView.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            vnut_per_dataGridView.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            vnut_per_dataGridView.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            vnut_per_dataGridView.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            vnut_per_dataGridView.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            vnut_per_dataGridView.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            vnut_per_dataGridView.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            vnut_per_dataGridView.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            vnut_per_dataGridView.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            vnut_per_dataGridView.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            vnut_per_dataGridView.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            vnut_per_dataGridView.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            vnut_per_dataGridView.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            vnut_per_dataGridView.Columns[15].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            vnut_per_dataGridView.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            vnut_per_dataGridView.Columns[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }
    }
}
