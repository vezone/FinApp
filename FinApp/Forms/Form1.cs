using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

using FinApp.Forms;
using FinApp.src;
using MetroFramework;

namespace FinApp
{

    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        //UI part
        List<Panel> panelsList = new List<Panel>();

        //Logic var
        private int       m_Index       = 0, 
                          m_TablesCount = 0;
        private bool      f_IsCreated   = false;

        private Tables    m_Tables;
        private DBConnect m_DB;
        private XL        m_XL;

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            switch (keyData)
            {
                case Keys.Control | Keys.S:
                    сохранитьToolStripMenuItem.PerformClick(); //имитируем нажатие
                    return true;
                case Keys.Control | Keys.T:
                    поДнямToolStripMenuItem.PerformClick();
                    break;
                case Keys.Control | Keys.Y:
                    финальныйОтчетToolStripMenuItem.PerformClick();
                    break;
                case Keys.Control | Keys.G:
                    дляТекущейТаблицыToolStripMenuItem.PerformClick();
                    break;
                case Keys.Control | Keys.H:
                    дляВсехТаблицToolStripMenuItem.PerformClick();
                    break;
                default:
                    break;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        public Form1()
        {
            InitializeComponent();
            Text = "Бухгалтерия";

            menuStrip1.Visible = false;
        }
        
        private void Init(string dbName)
        {
            menuStrip1.Visible = true;
            m_DB = new DBConnect(db: dbName);

            FillTables();
            m_XL = new XL(m_Tables);

            metroGrid1.Columns[0].DefaultCellStyle.Font =
                new Font("Arial", 16);
            metroGrid1.Columns[1].DefaultCellStyle.Font =
                new Font("Arial", 16);
            metroGrid1.Columns[2].DefaultCellStyle.Font =
                new Font("Arial", 16);
            metroGrid1.Columns[3].DefaultCellStyle.Font =
                new Font("Arial", 16);

            FillGrid();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Resizable = false;

            panelsList.Add(panel1);
            panelsList.Add(panel2);
            panelsList[0].BringToFront();

            metroGrid1.ColumnHeadersDefaultCellStyle.Font =
                new Font("Arial", 16);
            metroGrid1.Columns[0].DefaultCellStyle.Font =
                new Font("Arial", 16);
            metroGrid1.Columns[1].DefaultCellStyle.Font =
                new Font("Arial", 16);
            metroGrid1.Columns[2].DefaultCellStyle.Font =
                new Font("Arial", 16);
            metroGrid1.Columns[3].DefaultCellStyle.Font =
                new Font("Arial", 16);
            metroGrid1.AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.AllCells;
            metroGrid1.AutoSizeRowsMode =
                DataGridViewAutoSizeRowsMode.AllCells;
            metroGrid1.RowHeadersWidthSizeMode =
                DataGridViewRowHeadersWidthSizeMode.EnableResizing;
            metroGrid1.ColumnHeadersHeightSizeMode =
                DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            metroGrid1.GridColor = Color.Black;

            metroButton3.IsAccessible = false;
            metroButton4.IsAccessible = false;
            metroButton5.IsAccessible = false;
            metroButton6.IsAccessible = false;

            //Text = "Бухгалтерия";
        }
        private void panel1_Paint(object sender, PaintEventArgs e)
        {
        }
        private void panel2_Paint(object sender, PaintEventArgs e)
        {
        }
        
        private void MetroButton1_Click(object sender, EventArgs e)
        {
            panelsList[1].BringToFront();
            Init("06.2018");
        }
        private void metroButton2_Click(object sender, EventArgs e)
        {
            panelsList[1].BringToFront();
            Init("07.2018");
        }
        private void metroButton3_Click(object sender, EventArgs e)
        {
            //panelsList[1].BringToFront();
            //metroLabel1.Text = $"{metroButton3.Text}";
        }
        private void metroButton4_Click(object sender, EventArgs e)
        {
            //panelsList[1].BringToFront();
            //metroLabel1.Text = $"{metroButton4.Text}";
        }
        private void metroButton5_Click(object sender, EventArgs e)
        {
            //panelsList[1].BringToFront();
            //metroLabel1.Text = $"{metroButton5.Text}";
        }
        private void metroButton6_Click(object sender, EventArgs e)
        {
            //panelsList[1].BringToFront();
            //metroLabel1.Text = $"{metroButton6.Text}";
        }

        //menu strip stuff
        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
        }

        private void создатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreationForm m_CreationForm = new CreationForm();
            m_CreationForm.Activate();
            m_CreationForm.Visible = true;
            f_IsCreated = true;
        }
        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_DB.RunQuery(DBConnect.DeleteAllProducts(m_Tables[m_Index].Name));

            for (int r = 0; r < metroGrid1.RowCount - 1; r++)
            {
                m_DB.RunQuery(DBConnect.InsertProduct(
                    m_Tables[m_Index].Name,
                    metroGrid1.Rows[r].Cells[0].Value.ToString(),
                    metroGrid1.Rows[r].Cells[1].Value.ToString(),
                    metroGrid1.Rows[r].Cells[2].Value.ToString(),
                    metroGrid1.Rows[r].Cells[3].Value.ToString()));
            }
            MetroMessageBox.Show(this, 
                "Table was saved in db!", "Message", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_DB.RunQuery(DBConnect.DeleteTable(m_Tables[m_Index].Name));
            FillTables();
            FillGrid();
        }
        private void вернутьсяВМенюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panelsList[0].BringToFront();
        }
        
        private void metroButton7_Click(object sender, EventArgs e)
        {
            if (f_IsCreated)
            {
                FillTables();
                f_IsCreated = false;
            }
            if (m_Index > 0)
            {
                --m_Index;
                FillGrid();
            }
        }
        private void metroButton8_Click(object sender, EventArgs e)
        {
            if (f_IsCreated)
            {
                FillTables();
                f_IsCreated = false;
            }
            if (m_Index < (m_TablesCount - 1))
            {
                ++m_Index;
                FillGrid();
            }
        }
        private void metroGrid1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            for (int r = 0; r < metroGrid1.Rows.Count; r++)
            {
                if (metroGrid1.Rows[r].Cells[1].Value != null &&
                    metroGrid1.Rows[r].Cells[2].Value != null)
                {
                    try
                    {
                        metroGrid1.Rows[r].Cells[3].Value =
                        (Double.Parse(metroGrid1.Rows[r].Cells[1].Value.ToString())
                        *
                        Double.Parse(metroGrid1.Rows[r].Cells[2].Value.ToString()));
                    }
                    catch (FormatException ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

            }
        }

        private void MetroLabel1_Click(object sender, EventArgs e)
        {
        }

        private async void ПоДнямToolStripMenuItem_ClickAsync(object sender, EventArgs e)
        {
            await Task.Run(() => m_XL.GenerateDaysReportAsync());
        }
        private async void ФинальныйОтчетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await Task.Run(()=>m_XL.GenerateFinalReport());
        }
        private void ДляТекущейТаблицыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CurrentMessageReport();
        }
        private void ДляВсехТаблицToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TablesMessageReport();
        }

        private void КалькуляторToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process
                .Start(new System.Diagnostics.ProcessStartInfo()
                .FileName = "calc");
        }

        //FUNCS
        private void FillTables()
        {
            List<string> allTables  = m_DB.RunQueryW2(DBConnect.ShowAllTables());
            if (m_TablesCount != allTables.Count)
            {
                m_TablesCount  = allTables.Count;
                if (m_Index   >= m_TablesCount)
                    m_Index    = (m_TablesCount - 1);
            }   m_Tables       = new Tables(allTables.Count);

            for (int i = 0; i < allTables.Count; i++)
            {
                m_Tables[i] =
                    m_DB.RunQueryWT(DBConnect.SelectAll(allTables[i]));
                m_Tables[i].Name = allTables[i];
            }
        }
        private void FillGrid()
        {
            metroLabel1.Text = $"{m_Tables[m_Index].Name}";
            FillTables();
            metroGrid1.Rows.Clear();

            foreach (Product product in m_Tables[m_Index].Content)
            {
                metroGrid1.Rows.Add(product.Name,
                    product.Number, product.Price, product.Cost);
            }

        }
        
        private void CurrentMessageReport()
        {
            double Summary = 0.0;
            List<Product> list = m_Tables[m_Index].Content;
            for (int i = 0; i < list.Count; i++)
            {
                Summary += Double.Parse(list[i].Cost);
            }
            MetroMessageBox.Show(this,
                $"{m_Tables[m_Index].Name}\nСуммарные траты: {Summary.ToString()}", "Отчет",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void TablesMessageReport()
        {
            List<ParsedTables> m_ParsedTables = m_Tables.ParseTablesByDates();
            string str = "";
            str += $"Nuber of table (by date): {m_ParsedTables.Count}\n";
            double globalSummary = 0.0;
            int csdIndex = 0;
            double[] cacheSpendingDate = new double[m_ParsedTables.Count];

            for (int i = 0; i < m_ParsedTables.Count; i++)
            {
                cacheSpendingDate[csdIndex++] =
                    m_ParsedTables[i].GetCacheSpending();
                globalSummary += m_ParsedTables[i].GetCacheSpending();
            }

            for (int i = 0; i < m_ParsedTables.Count; i++)
            {
                str +=
                    $"\n{m_ParsedTables[i].Name}:\n{m_ParsedTables[i].GetStringCacheSpending()}";
                str += $"Global: { m_ParsedTables[i].GetCacheSpending()}\n";
            }

            str += $"\n{globalSummary} \n";
            MessageBox.Show($"{str}");
        }

        //Legacy
        public void RecreateTables()
        {
            for (int i = 0; i < m_Tables.TablesCount; i++)
            {
                m_DB.RunQuery(DBConnect
                    .CreateTable(m_Tables[i].Name, m_Tables[i].LowName));
                DBConnect.DataBase = m_Tables[i].LowName;
                for (int j = 0; j < m_Tables[i].Content.Count; j++)
                {
                    m_DB.RunQuery(DBConnect
                        .InsertProduct(m_Tables[i].Name,
                        m_Tables[i].Content[j].Name,
                        m_Tables[i].Content[j].Number,
                        m_Tables[i].Content[j].Price,
                        m_Tables[i].Content[j].Cost));
                }
            }
        }
        public void RenameTables(string dbName, string tableName, string newName)
        {
            DBConnect.DataBase = dbName;
            for (int i = 0; i < m_Tables.TablesCount; i++)
            {
                if (m_Tables[i].Name.Contains(tableName))
                {
                    m_DB.RunQuery(DBConnect
                        .RenameTable($"{m_Tables[i].FirstName}_{tableName}",
                                     $"{m_Tables[i].FirstName}_{newName}"));
                }
            }
        }

    }
}
