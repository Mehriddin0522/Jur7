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
    public partial class LoginForm : Form
    {
        Connect sql = new Connect();
        public LoginForm()
        {
            InitializeComponent();
            sql.Connection();
            run_main();



            login_comboBox.SelectedIndex = 0;
            pasword_textBox.Text = "1";
        }

        public void run_main()
        {
            //try
            //{
                sql.myReader = sql.return_MySqlCommand("select user from users_jur7").ExecuteReader();
                while (sql.myReader.Read())
                {
                    login_comboBox.Items.Add(sql.myReader.GetString("user"));
                }
                sql.myReader.Close();

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("run_main " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        public string string_for_otdels;
        private void login_btn_Click(object sender, EventArgs e)
        {
            //try
            //{
                if (login_comboBox.Text != "" && pasword_textBox.Text != "")
                {

                    sql.myReader = sql.return_MySqlCommand("select password from users_jur7 where user = '" + login_comboBox.Text + "' ").ExecuteReader();
                    while (sql.myReader.Read())
                    {
                        if (sql.myReader.GetString("password") == pasword_textBox.Text)
                        {
                            string_for_otdels = login_comboBox.Text;

                            //   main.run_label_main();



                            //  main.otdelNumber_textBox.Text = "Task";

                            this.Hide();
                            MainForm main = new MainForm(string_for_otdels);

                            //main.WindowState = FormWindowState.Maximized;
                            main.ShowDialog();
                            this.Show();


                        }
                    }
                    sql.myReader.Close();
                }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("login_btn_Click " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void login_comboBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                pasword_textBox.Focus();
            }
        }

        private void pasword_textBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                login_btn.Focus();
            }
        }
    }
}
