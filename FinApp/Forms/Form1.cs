using FinApp.Forms;
using FinApp.src;
using MetroFramework;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace FinApp
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        //UI part
        List<Panel> panelsList = new List<Panel>();

        //Logic var
        private int  m_Index       = 0, 
                     m_TablesCount = 0;
        private bool f_IsCreated   = false;
        private Table[] m_TableList;
        private List<ParsedTables> m_ParsedTables;
        private DBConnect m_DB;
        
        Excel.Application m_ExcelApplication;
        Excel.Workbook m_WorkBook;
        Excel.Worksheet[] m_WorkSheetCollection;
        int m_WSCollectionIndex;

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

            //RecreateTables();
            //int numberOfTables = Int32.Parse(m_DB.RunQueryW(DBConnect.NumberOfTables("goods")));
            //MessageBox.Show(numberOfTables.ToString());
            //DBConnect.DataBase = "goods";
            //List<string> tablesList = m_DB.RunQueryW2(DBConnect.ShowAllTables());
            //MessageBox.Show(tablesList[19]);
        }
        
        private void Init(string dbName)
        {
            menuStrip1.Visible = true;
            m_DB = new DBConnect(db: dbName);

            FillTables();

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
            m_DB.RunQuery(DBConnect.DeleteAllProducts(m_TableList[m_Index].Name));

            for (int r = 0; r < metroGrid1.RowCount - 1; r++)
            {
                m_DB.RunQuery(DBConnect.InsertProduct(
                    m_TableList[m_Index].Name,
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
            m_DB.RunQuery(DBConnect.DeleteTable(m_TableList[m_Index].Name));
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
        private void metroLabel1_Click(object sender, EventArgs e)
        {
        }

        private async void ПоДнямToolStripMenuItem_ClickAsync(object sender, EventArgs e)
        {
            await GenerateDaysReportAsync();
        }
        private async void ФинальныйОтчетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await Task.Run(()=>GenerateFinalReport());
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
            List<string> text  = m_DB.RunQueryW2(DBConnect.ShowAllTables());
            if (m_TablesCount != text.Count)
            {
                m_TablesCount  = text.Count;
                if (m_Index   >= m_TablesCount)
                    m_Index    = (m_TablesCount - 1);
            } m_TableList      = new Table[text.Count];

            for (int i = 0; i < text.Count; i++)
            {
                m_TableList[i] =
                    m_DB.RunQueryWT(DBConnect.SelectAll(text[i]));
                m_TableList[i].Name = text[i];
            }
        }
        private void FillGrid()
        {
            metroLabel1.Text = $"{m_TableList[m_Index].Name}";
            FillTables();
            metroGrid1.Rows.Clear();

            foreach (Product product in m_TableList[m_Index].Content)
            {
                metroGrid1.Rows.Add(product.Name,
                    product.Number, product.Price, product.Cost);
            }

        }
        private int NumberOfIndividDates()
        {
            int length = 0;
            List<Table> tList = new List<Table>(m_TableList);

            for (int i = 0; i < tList.Count; i++)
            {
                Table table = tList[i];
                length++;
                for (int j = i + 1; j < tList.Count; j++)
                {
                    if (table.FirstName == tList[j].FirstName)
                    {
                        tList.RemoveAt(j);
                    }
                }
                tList.RemoveAt(i);
            }

            return length;
        }
        
        //______EXCEL STUFF
        private async Task ParseTables()
        {
            List<Table> list = new List<Table>(1);
            List<Table> tList = new List<Table>(m_TableList);

            for (int i = 0; i < tList.Count; i++)
            {
                Table table = tList[i];
                list.Add(table);
                for (int j = i + 1; j < tList.Count; j++)
                {
                    if (table.FirstName == tList[j].FirstName)
                    {
                        list.Add(tList[j]);
                        tList.RemoveAt(j);
                        j--;
                    }
                }
                await Task.Run(() => GenerateExcelDatePage(table.FirstName, list));
                list.Clear();
            }
        }
        private async Task GenerateDaysReportAsync()
        {
            //Excel stuff
            m_ExcelApplication = new Excel.Application();
            m_WorkBook = m_ExcelApplication.Workbooks.Add();
            m_WorkSheetCollection = new Excel.Worksheet[NumberOfIndividDates() + 1];
            m_WSCollectionIndex = 0;

            await Task.Run(() => ParseTables());
        }

        private void GenerateFinalReport()
        {
            m_ExcelApplication = new Excel.Application();
            m_WorkBook = m_ExcelApplication.Workbooks.Add();
            m_WorkSheetCollection = new Excel.Worksheet[1];
            m_WorkSheetCollection[0] =
                (Excel.Worksheet)m_WorkBook.Worksheets.get_Item(1);
            m_WorkSheetCollection[0] = m_WorkBook.Worksheets.Add();

            m_WorkSheetCollection[0].StandardWidth = 20;
            m_WSCollectionIndex = 0;
            m_WorkSheetCollection[0].Name = "FinalReport";

            //setting
            m_WorkSheetCollection[0].Cells.HorizontalAlignment = Excel.Constants.xlCenter;
            m_WorkSheetCollection[0].Cells.Font.Size = 14;
            m_WorkSheetCollection[0].Cells.Font.Name = "Times New Roman";

            ParseTablesByDates();

            m_WorkSheetCollection[0].Cells[2, 1].Value = "Кол-во таблиц";
            m_WorkSheetCollection[0].Cells[2, 2].Value = m_ParsedTables.Count;


            for (int i = 3; i < m_ParsedTables.Count + 3; i++)
            {
                m_WorkSheetCollection[0].Cells[i, 1].Value = m_ParsedTables[i - 3].Date;
                m_WorkSheetCollection[0].Cells[i, 2].Value =
                    (int)m_ParsedTables[i - 3].GetCacheSpending();
            }

            m_WorkSheetCollection[0].Cells[m_ParsedTables.Count + 3, 1].Value =
                "Общая сумма";
            m_WorkSheetCollection[0].Cells[m_ParsedTables.Count + 3, 2].FormulaLocal =
                $"=СУММ(B{3}:B{m_ParsedTables.Count + 2})";

            try
            {
                var xlCharts =
                    m_WorkSheetCollection[0].ChartObjects() as Excel.ChartObjects;
                Excel.ChartObject myChart = xlCharts.Add(230, 20, 450, 200);
                Excel.Chart chartPage = myChart.Chart;
                object ob1 = $"B{3}";
                object ob2 = $"B{m_ParsedTables.Count + 2}";
                var chartRange = m_WorkSheetCollection[0].get_Range(ob1, ob2);
                object ob3 = $"A{3}";
                var chartRange2 = m_WorkSheetCollection[0].get_Range(ob3, ob2);

                chartPage.SetSourceData(chartRange);
                chartPage.ChartType = Excel.XlChartType.xlLine;//xlColumnClustered
                chartPage.ChartWizard(
                    Source: chartRange2,
                    Title: "График расходов по дням",
                    CategoryTitle: "Дата",
                    ValueTitle: "Деньги");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show(ex.Message);
            }

            // Открываем созданный excel-файл
            m_ExcelApplication.Visible = true;
            m_ExcelApplication.UserControl = true;
        }
        private void CurrentMessageReport()
        {
            double Summary = 0.0;
            List<Product> list = m_TableList[m_Index].Content;
            for (int i = 0; i < list.Count; i++)
            {
                Summary += Double.Parse(list[i].Cost);
            }
            MetroMessageBox.Show(this,
                $"{m_TableList[m_Index].Name}\nСуммарные траты: {Summary.ToString()}", "Отчет",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void TablesMessageReport()
        {
            ParseTablesByDates();
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
                    $"\n{m_ParsedTables[i].Date}:\n{m_ParsedTables[i].GetStringCacheSpending()}";
                str += $"Global: { m_ParsedTables[i].GetCacheSpending()}\n";
            }

            str += $"\n{globalSummary} \n";
            MessageBox.Show($"{str}");
        }
        private void GenerateExcelDatePage(string date, List<Table> list)
        {
            //Basic
            m_WorkSheetCollection[m_WSCollectionIndex] =
                m_WorkBook.Worksheets.Add();
            m_WorkSheetCollection[m_WSCollectionIndex].Name = date;
            m_WorkSheetCollection[m_WSCollectionIndex].StandardWidth
                = 30;

            int RowToBegin = 2;
            int ColumnToBegin = 1;

            for (int i = 0, j = 0; j < list.Count; j++, i = i + 3)
            {
                //grid
                var cells = m_WorkSheetCollection[m_WSCollectionIndex]
                    .get_Range($"{(char)('A' + RowToBegin - 2)}{1 + ColumnToBegin + i}",
                               $"{(char)('B' + RowToBegin - 2)}{3 + ColumnToBegin + i}");
                cells.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle
                    = Excel.XlLineStyle.xlContinuous; // внутренние вертикальные
                cells.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle
                    = Excel.XlLineStyle.xlContinuous; // внутренние горизонтальные            
                cells.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle
                    = Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                cells.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle
                    = Excel.XlLineStyle.xlContinuous; // правая внешняя
                cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle
                    = Excel.XlLineStyle.xlContinuous; // левая внешняя
                cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle
                    = Excel.XlLineStyle.xlContinuous;

                //individ
                m_WorkSheetCollection[m_WSCollectionIndex].Cells[i + RowToBegin, ColumnToBegin]
                    .Font.Bold = true;
                m_WorkSheetCollection[m_WSCollectionIndex].Range[
                    m_WorkSheetCollection[m_WSCollectionIndex].Cells[i + RowToBegin, ColumnToBegin],
                    m_WorkSheetCollection[m_WSCollectionIndex].Cells[i + RowToBegin, ColumnToBegin + 1]
                    ].Font.Size = 20;
                m_WorkSheetCollection[m_WSCollectionIndex].Range[
                    m_WorkSheetCollection[m_WSCollectionIndex].Cells[i + RowToBegin, ColumnToBegin],
                    m_WorkSheetCollection[m_WSCollectionIndex].Cells[i + RowToBegin, ColumnToBegin + 1]
                    ].Merge();

                //Set Times New Roman to ALL
                m_WorkSheetCollection[m_WSCollectionIndex].Cells.HorizontalAlignment
                   = Excel.Constants.xlCenter;
                m_WorkSheetCollection[m_WSCollectionIndex].Cells.Font.Name
                    = "Times New Roman";
                m_WorkSheetCollection[m_WSCollectionIndex].Range[
                    m_WorkSheetCollection[m_WSCollectionIndex].Cells[i + 3, ColumnToBegin],
                    m_WorkSheetCollection[m_WSCollectionIndex].Cells[i + 4, ColumnToBegin + 1]
                    ].Font.Size = 14;

                //Заголовок
                m_WorkSheetCollection[m_WSCollectionIndex].Cells[i + RowToBegin, ColumnToBegin] =
                    list[j].SecondName;

                //Параметры
                m_WorkSheetCollection[m_WSCollectionIndex].Cells[i + 3, ColumnToBegin] =
                    "кол-во";
                m_WorkSheetCollection[m_WSCollectionIndex].Cells[i + 3, ColumnToBegin + 1] =
                    "стоимость";

                //Значения
                m_WorkSheetCollection[m_WSCollectionIndex].Cells[i + 4, ColumnToBegin] =
                    list[j].Content.Count;
                m_WorkSheetCollection[m_WSCollectionIndex].Cells[i + 4, ColumnToBegin + 1] =
                    list[j].GetSummary();

            }

            m_WorkSheetCollection[m_WSCollectionIndex] = null;
            ++m_WSCollectionIndex;

            // Открываем созданный excel-файл
            m_ExcelApplication.Visible = true;
            m_ExcelApplication.UserControl = true;
        }
        private void ParseTablesByDates()
        {
            List<Table> tList = new List<Table>(m_TableList);
            m_ParsedTables = new List<ParsedTables>(NumberOfIndividDates());

            for (int i = 0; i < tList.Count; i++)
            {
                ParsedTables parsedTables = new ParsedTables(tList[i].FirstName);
                parsedTables.Add(tList[i]);
                for (int j = i + 1; j < tList.Count; j++)
                {
                    if (tList[i].FirstName == tList[j].FirstName)
                    {
                        parsedTables.Add(tList[j]);
                        tList.RemoveAt(j);
                        j--;
                    }
                }
                m_ParsedTables.Add(parsedTables);
            }
        }
        
        //Legacy
        private void RecreateTables()
        {
            for (int i = 0; i < m_TableList.Length; i++)
            {
                m_DB.RunQuery(DBConnect
                    .CreateTable(m_TableList[i].Name, m_TableList[i].LowName));
                DBConnect.DataBase = m_TableList[i].LowName;
                for (int j = 0; j < m_TableList[i].Content.Count; j++)
                {
                    m_DB.RunQuery(DBConnect
                        .InsertProduct(m_TableList[i].Name, 
                        m_TableList[i].Content[j].Name,
                        m_TableList[i].Content[j].Number,
                        m_TableList[i].Content[j].Price,
                        m_TableList[i].Content[j].Cost));
                }
            }
        }
        private void RenameTables(string dbName, string tableName, string newName)
        {
            DBConnect.DataBase = dbName;
            for (int i = 0; i < m_TableList.Length; i++)
            {
                if (m_TableList[i].Name.Contains(tableName))
                {
                    m_DB.RunQuery(DBConnect
                        .RenameTable($"{m_TableList[i].FirstName}_{tableName}",
                                     $"{m_TableList[i].FirstName}_{newName}"));
                }
            }
            FillTables();
            FillGrid();
        }
        
    }
}
