using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace jur7
{
    public class Connect
    {
        public MySqlCommand SelectCommand;
        public MySqlDataReader myReader;
        public MySqlConnection myConn;
        public MySqlDataAdapter mydataAdapter;
        public void Connection()
        {
            try
            {
                string ip = System.IO.File.ReadAllText("docs\\ip.txt");
                string database = System.IO.File.ReadAllText("docs\\database.txt");
                string myConnection = "datasource=" + ip + ";port=3306;username=root;password=1101jamshid;database=" + database + ";charset=utf8";
                myConn = new MySqlConnection(myConnection);
                myConn.Open();
                Console.WriteLine("Connection opened");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        public MySqlCommand return_MySqlCommand(string mysqlcommand)
        {
            SelectCommand = new MySqlCommand(mysqlcommand, myConn);
            return SelectCommand;
        }

        public MySqlDataReader select_return_MySqlDataReader()
        {
            return myReader;
        }
        public void insert_MySqlDataReader()
        {
            myReader = SelectCommand.ExecuteReader();
        }
    }
}
