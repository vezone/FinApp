using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace FinApp.src
{
    class XL
    {
        private Tables m_Tables;

        int               m_WSCollectionIndex;
        Excel.Application m_ExcelApplication;
        Excel.Workbook    m_WorkBook;
        Excel.Worksheet[] m_WorkSheetCollection;
        
        public XL(Tables tables)
        {
            m_Tables = new Tables(tables);
        }

        public async Task GenerateDaysReportAsync()
        {
            //Excel stuff
            m_ExcelApplication = new Excel.Application();
            m_WorkBook = m_ExcelApplication.Workbooks.Add();
            m_WorkSheetCollection = new Excel.Worksheet[m_Tables.NumberOfIndividDates()];
            m_WSCollectionIndex = 0;

            List<ParsedTables> parsedTables = m_Tables.ParseTablesFirstName();

            for (int i = 0; i < parsedTables.Count; i++)
            {
                await Task.Run(
                    ()=>GenerateExcelDatePage(
                        parsedTables[i].Name,
                        parsedTables[i]));
            }

        }
        public void GenerateExcelDatePage(string date, ParsedTables list)
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

        public void FillThePage(out Excel.Worksheet page)
        {
            page =
                (Excel.Worksheet)m_WorkBook.Worksheets.get_Item(1);
            page = m_WorkBook.Worksheets.Add();
            
            List<ParsedTables> parsedTables = m_Tables.ParseTablesFirstName();
            
            page.StandardWidth = parsedTables[0].LongestName().Length * 1.29f;
            page.Name = "FinalReport";
            page.Cells.HorizontalAlignment = Excel.Constants.xlCenter;
            page.Cells.Font.Size = 14;
            page.Cells.Font.Name = "Times New Roman";

            //1
            page.Cells[2, 1].Value = "Кол-во таблиц";
            page.Cells[2, 2].Value = parsedTables.Count;

            (var names1, var values1) = m_Tables.SplitList(m_Tables.ParseTablesFirstName);
            ExcelPrototype.SetTable2(
                ref page,
                3, names1.Count + 3,
                names1, values1);

            ExcelChartInfo excelCI =
                new ExcelChartInfo().Left(530).Top(20).Width(450).Heigth(200)
                .NameLRange($"B{3}").NameHRange($"B{values1.Count + 2}")
                .SourceLRange($"A{3}").SourceHRange($"B{values1.Count + 2}")
                .Title("По дням").CategoryTitle("Дата").ValueTitle("Деньги")
                .ChartType(Excel.XlChartType.xlLine);

            ExcelPrototype.CreateChart(ref page, excelCI);

            //2
            (var names2, var values2) = m_Tables.SplitList(m_Tables.ParseTablesSecondName);
            ExcelPrototype.SetTable2(
                ref page,
                (names1.Count + 5), (names1.Count + names2.Count + 5),
                names2, values2);
            
            ExcelChartInfo excelCI2 =
                new ExcelChartInfo().Left(530).Top(240).Width(450).Heigth(350)
                .NameLRange($"B{names1.Count + 5}")
                .NameHRange($"B{(names1.Count + names2.Count + 4)}")
                .SourceLRange($"A{names1.Count + 5}")
                .SourceHRange($"B{(names1.Count + names2.Count + 4)}")
                .Title("По типам трат").CategoryTitle("Дата").ValueTitle("Деньги")
                .ChartType(Excel.XlChartType.xlColumnClustered);
            
            ExcelPrototype.CreateChart(ref page, excelCI2);
        }

        public void GenerateFinalReport()
        { 
            m_ExcelApplication = new Excel.Application();
            m_ExcelApplication.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
            m_WorkBook = m_ExcelApplication.Workbooks.Add();

            m_WorkSheetCollection = new Excel.Worksheet[1];
            
            FillThePage(out m_WorkSheetCollection[m_WSCollectionIndex]);
            
            // Открываем созданный excel-файл
            m_ExcelApplication.Visible = true;
            m_ExcelApplication.UserControl = true;
        }
        
    }


    class ExcelPrototype
    {
        public static void SetTable2<FirstT, SecondT>(
            ref Excel.Worksheet page, 
            int left, int bot, 
            List<FirstT> fValues, List<SecondT> sValues)
        {
            var cells = page
                    .get_Range(
                          $"A{left}", 
                          $"B{bot}");
                //topPosition, botPosition);
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
            
            for (int i = left; i < bot; i++)
            {
                page.Cells[i, 1].Value = fValues[i - left];
                page.Cells[i, 2].Value = sValues[i - left];
            }
            page.Cells[fValues.Count + left, 1].Value =
                "Общая сумма";
            page.Cells[fValues.Count + left, 2].FormulaLocal =
                $"=СУММ(B{left}:B{bot-1})";
        }

        public static void CreateChart(
            ref Excel.Worksheet page,
            ExcelChartInfo excelCI)
        {
            try
            {
                var xlCharts =
                    page.ChartObjects() as Excel.ChartObjects;
                Excel.ChartObject myChart = 
                    xlCharts.Add(excelCI.m_Left, excelCI.m_Top, excelCI.m_Width, excelCI.m_Height);
                Excel.Chart chartPage = myChart.Chart;

                var chartRange = page.get_Range(excelCI.m_NameLRange, excelCI.m_NameHRange);
                var chartRange2 = page.get_Range(excelCI.m_SourceLRange, excelCI.m_SourceHRange);

                chartPage.SetSourceData(chartRange);
                chartPage.ChartType =
                    excelCI.m_ChartType; //xlLine || xlColumnClustered
                chartPage.ChartWizard(
                    Source: chartRange2,
                    Title: excelCI.m_Title,
                    CategoryTitle: excelCI.m_CategoryTitle,
                    ValueTitle: excelCI.m_ValueTitle);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

    }

    class ExcelChartInfo
    {
        public int m_Left;
        public int m_Top;
        public int m_Width;
        public int m_Height;
        
        public string m_NameLRange;
        public string m_NameHRange;
        public string m_SourceLRange;
        public string m_SourceHRange;

        public string m_Title;
        public string m_ValueTitle;
        public string m_CategoryTitle;

        public Excel.XlChartType m_ChartType;

        public ExcelChartInfo Left(int value)
        {
            m_Left = value;
            return this;
        }
        public ExcelChartInfo Top(int value)
        {
            m_Top = value;
            return this;
        }
        public ExcelChartInfo Width(int value)
        {
            m_Width = value;
            return this;
        }
        public ExcelChartInfo Heigth(int value)
        {
            m_Height = value;
            return this;
        }

        public ExcelChartInfo NameLRange(string value)
        {
            m_NameLRange = value;
            return this;
        }
        public ExcelChartInfo NameHRange(string value)
        {
            m_NameHRange = value;
            return this;
        }
        public ExcelChartInfo SourceLRange(string value)
        {
            m_SourceLRange = value;
            return this;
        }
        public ExcelChartInfo SourceHRange(string value)
        {
            m_SourceHRange = value;
            return this;
        }
        
        public ExcelChartInfo Title(string value)
        {
            m_Title = value;
            return this;
        }
        public ExcelChartInfo ValueTitle(string value)
        {
            m_ValueTitle = value;
            return this;
        }
        public ExcelChartInfo CategoryTitle(string value)
        {
            m_CategoryTitle = value;
            return this;
        }

        public ExcelChartInfo ChartType(Excel.XlChartType type)
        {
            m_ChartType = type;
            return this;
        }

    }

}