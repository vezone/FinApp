using System;
using System.Windows.Forms;

using FinApp.src;

namespace FinApp.Forms
{
    public partial class CreationForm : MetroFramework.Forms.MetroForm
    {
        //Remake it
        DBConnect m_DB = new DBConnect();
        
        public CreationForm()
        {
            InitializeComponent();
        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            metroTextBox1.Text = DateTime.Now.ToShortDateString();
        }

        private void CreationForm_Load(object sender, EventArgs e)
        {
            Resizable = false;

            metroComboBox1.Items.Add("Техника");
            metroComboBox1.Items.Add("Развлечения");
            metroComboBox1.Items.Add("Недвижимость");
            metroComboBox1.Items.Add("Семейные траты");
            metroComboBox1.Items.Add("Лента и продукты");
            metroComboBox1.Items.Add("Кошачьи продукты");
            metroComboBox1.Items.Add("Движимость и передвижения");
        }

        private void metroLabel3_Click(object sender, EventArgs e)
        {

        }

        private void metroLabel1_Click(object sender, EventArgs e)
        {

        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            string[] dateTimeStuff = metroTextBox1.Text.Split('.');
            string DataBase = $"{dateTimeStuff[1]}.{dateTimeStuff[2]}";
            DBConnect.DataBase = DataBase;
            string result = 
                m_DB
                .RunQueryW(DBConnect.IsTableExist(metroTextBox2.Text));

            if (result != "1")
            {
                MessageBox.Show("Creating db!");
                m_DB.RunQuery(DBConnect.CreateDB(DataBase));
                m_DB.RunQuery(DBConnect.CreateTable(metroTextBox2.Text));
                MessageBox
                    .Show($"Table {metroTextBox2.Text} was created![{DataBase}]");
            }
            else MessageBox.Show($"Table {metroTextBox2.Text} is already exist");
            Close();
        }

        private void metroComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (metroComboBox1.SelectedIndex)
            {
                case 0:
                    {
                        metroTextBox2.Text = metroTextBox1.Text + "_Техника";
                    }
                    break;
                case 1:
                    {
                        metroTextBox2.Text = metroTextBox1.Text + "_Развлечения";
                    }
                    break;
                case 2:
                    {
                        metroTextBox2.Text = metroTextBox1.Text + "_Недвижимость";
                    }
                    break;
                case 3:
                    {
                        metroTextBox2.Text = metroTextBox1.Text + "_Семейные траты";
                    }
                    break;
                case 4:
                    {
                        metroTextBox2.Text = metroTextBox1.Text + "_Лента и продукты";
                    }
                    break;
                case 5:
                    {
                        metroTextBox2.Text = metroTextBox1.Text + "_Кошачьи продукты";
                    }
                    break;
                case 6:
                    {
                        metroTextBox2.Text = metroTextBox1.Text + "_Движимость и передвижения";
                    }
                    break;
            }
        }

        private void metroTextBox2_Click(object sender, EventArgs e)
        {

        }

        //dont work correctly
        private void metroTextBox1_Click(object sender, EventArgs e)
        {
            if (metroComboBox1.Text != null)
            {
                metroTextBox2.Text = 
                    $"{metroTextBox1.Text}_{metroComboBox1.Text}";
            }
            else
            {
                metroTextBox2.Text = metroTextBox1.Text;
            }
        }
    }
}
