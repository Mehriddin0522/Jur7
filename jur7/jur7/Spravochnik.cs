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
    public partial class Spravochnik : Form
    {
        public Spravochnik()
        {
            InitializeComponent();
        }

        private void prixod_btn_Click(object sender, EventArgs e)
        {

        }

        private void prixod_btn_Click_1(object sender, EventArgs e)
        {
            //try
            //{

            Podrazdelenie podraz = new Podrazdelenie();

            podraz.Show();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("prixod_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }
    }
}
