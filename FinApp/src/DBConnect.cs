using System.Collections.Generic;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace FinApp.src
{
    class DBConnect
    {
        private static MySqlConnection m_Connection;

        public static string Server   { get; set; }
        public static string DataBase { get; set; }
        public static string Login    { get; set; }
        public static string Password { get; set; }

        public static string Table    { get; set; }

        private DBConnect()
        {
            //we are not using it
        }

        public DBConnect(
            string server = "localhost", string db = "goods",
            string login = "root", string password = "root")
        {
            Server = server;
            DataBase = db;
            Login = login;
            Password = password;

            string ConnectionString =
                "SERVER=" + Server + ";" +
                "DATABASE=" + DataBase + ";" +
                "UID=" + Login + ";" +
                "PASSWORD=" + Password + ";";
            m_Connection = new MySqlConnection(ConnectionString); ;
        }
        
        private bool OpenConnection()
        {
            try
            {
                m_Connection.Open();
                return true;
            }
            catch (MySqlException ex)
            {
                switch (ex.Number)
                {
                    case 0:
                        MessageBox.Show("Cannot connect to server.");
                        break;
                    case 1045:
                        MessageBox.Show("Invalid username/password, " +
                            "please try again.");
                        break;
                }
                return false;
            }

        }
        private bool CloseConnection()
        {
            try
            {
                m_Connection.Close();
                return true;
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }
        
        public void RunQuery(string query)
        {
            if (OpenConnection())
            {
                try
                {
                    new MySqlCommand(query, m_Connection).ExecuteNonQuery();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }
                CloseConnection();
            }
        }
        public string RunQueryW(string query)
        {
            string result = "";
            if (OpenConnection())
            {
                result = new MySqlCommand(query, m_Connection).ExecuteScalar().ToString();
                CloseConnection();
            }
            return result;
        }
        public List<string> RunQueryW2(string query)
        {
            List<string> list = new List<string>(0);

            if (OpenConnection())
            {
                MySqlDataReader reader
                    = new MySqlCommand(query, m_Connection).ExecuteReader();

                while (reader.Read())
                {
                    list.Add(reader.GetString(0));
                }

                CloseConnection();
            }
            return list;
        }
        public Table RunQueryWT(string query)
        {
            Table table = new Table();

            if (OpenConnection())
            {
                MySqlDataReader reader
                    = new MySqlCommand(query, m_Connection).ExecuteReader();

                while (reader.Read())
                {
                    table.Add(new Product(
                        reader.GetString(0),
                        reader.GetString(1),
                        reader.GetString(2),
                        reader.GetString(3)));
                }

                CloseConnection();
            }
            return table;
        }

        public static string SelectAll(string tableName)
        {
            return $"SELECT * FROM `{DataBase}`.`{tableName}`";
        }

        public static string CreateTable(string tableName)
        {
            return $"CREATE TABLE `{DataBase}`.`{tableName}` (`Товар` VARCHAR(255) NOT NULL,`Количество` VARCHAR(255) NULL, `Цена` VARCHAR(255) NULL, `Стоимость` VARCHAR(255) NULL, PRIMARY KEY(`Товар`));";
        }

        public static string CreateTable(string tableName, string db = "goods")
        {
            return $"CREATE TABLE `{db}`.`{tableName}` (`Товар` VARCHAR(255) NOT NULL,`Количество` VARCHAR(255) NULL, `Цена` VARCHAR(255) NULL, `Стоимость` VARCHAR(255) NULL, PRIMARY KEY(`Товар`));";
        }

        public static string DeleteTable(string tableName)
        {
            return $"DROP TABLE `{DataBase}`.`{tableName}`";
        }

        public static string RenameTable(string tableName, string newName)
        {
            return $"ALTER TABLE `{DataBase}`.`{tableName}`"+
                   $"RENAME TO  `{DataBase}`.`{newName}` ; ";
        }
        
        public static string InsertProduct(string tableName, string product, string number, string nprice, string price)
        {
            return $"INSERT INTO `{DataBase}`.`{tableName}` (`Товар`, `Количество`, `Цена`, `Стоимость`) VALUES('{product}', '{number}', '{nprice}', '{price}');";
        }

        public static string DeleteProduct(string tableName, string product)
        {
            return $"DELETE FROM `{DataBase}`.`{tableName}` WHERE `Товар`= '{product}';";
        }

        public static string DeleteAllProducts(string tableName)
        {
            return $"DELETE FROM `{DataBase}`.`{tableName}`;";
        }

        public static string ShowAllTables()
        {
            return $"SHOW TABLES FROM `{DataBase}`;";
        }

        public static string NumberOfTables(string dbName)
        {
            return $"SELECT COUNT(*) FROM information_schema.TABLES WHERE TABLE_SCHEMA = '{dbName}'; ";
        }

        public static string IsTableExist(string tableName)
        {
            return $"select count(*) from information_schema.tables where table_type = 'BASE TABLE' and table_name = '{tableName}';";
        }

        //TODO: show databases; 
        public static string ShowDB()
        {
            return "Show databases;";
        }

        public static string CreateDB(string dbName)
        {
            return $"create schema if not exists `{dbName}`;";
        }

    }
}
