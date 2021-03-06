﻿#region Namespaces

using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;

using Autodesk.Revit.DB;

using Microsoft.Win32;

using Excel = Microsoft.Office.Interop.Excel;

#endregion

namespace RevitUtils
{
    /// <summary>
    /// Логика взаимодействия для CreateSheets.xaml
    /// </summary>
    public partial class CreateSheets
    {
        private Excel.Application _xlApp;

        private Excel.Workbook _xlWorkBook;

        private Excel.Worksheet _xlWorkSheet;

        private Excel.Range _range;

        private int _rowCount;

        private int _colCount;

        public IEnumerable<FamilySymbol> TitleBlocks { get; }

        private readonly Document _doc;

        public CreateSheets(Document doc)
        {
            InitializeComponent();
            this._doc = doc;

            // get titleblock (family types of viewsheets)
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FamilySymbol));
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
            TitleBlocks = collector.ToElements().Cast<FamilySymbol>();

            FamilyForViewSheetBox.ItemsSource = TitleBlocks;
        }

        private void LoadExcelBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
                FilePathTextBox.Text = openFileDialog.FileName;

            if (FilePathTextBox.Text != string.Empty)
            {
                // open the Excel file and get the specified worksheet
                int rCnt;

                _xlApp = new Excel.Application();

                // open the excel
                _xlWorkBook = _xlApp.Workbooks.Open(FilePathTextBox.Text);

                // get the first sheet of the excel
                _xlWorkSheet = (Excel.Worksheet)_xlWorkBook.Worksheets.Item[1];

                _range = _xlWorkSheet.UsedRange;

                // get the total row count
                _rowCount = _range.Rows.Count;

                // get the total column count
                _colCount = _range.Columns.Count;

                var myRows = new List<MyRow>();

                if (_colCount < 2)
                {
                    // traverse all the row in the excel
                    for (rCnt = 1; rCnt <= _rowCount; rCnt++)
                    {
                        var myRow = new MyRow { Col1 = (string)(_range.Cells[rCnt, 1] as Excel.Range)?.Value2.ToString(), };

                        myRows.Add(myRow);
                    }
                }
                else
                {
                    // traverse all the row in the excel
                    for (rCnt = 1; rCnt <= _rowCount; rCnt++)
                    {
                        // traverse columns (the first column is not included)
                        for (int col = 2; col <= _colCount; col++)
                        {
                            var myRow = new MyRow
                                            {
                                                Col1 = (string)(_range.Cells[rCnt, 1] as Excel.Range)?.Value2.ToString(),
                                                Col2 = (string)(_range.Cells[rCnt, col] as Excel.Range)?.Value2.ToString()
                                            };

                            myRows.Add(myRow);
                        }
                    }
                }

                GridView1.ItemsSource = myRows;

                // release the resources
                _xlWorkBook.Close(true);
                _xlApp.Quit();
                Marshal.ReleaseComObject(_xlWorkSheet);
                Marshal.ReleaseComObject(_xlWorkBook);
                Marshal.ReleaseComObject(_xlApp);
            }
        }

        private void CreateSheetsBtn_Click(object sender, RoutedEventArgs e)
        {
            Transaction tran = new Transaction(_doc);
            tran.Start("Create Sheets");

            FamilySymbol fs = null;
            bool isSheetCreated = false;

            try
            {
                fs = (FamilySymbol)FamilyForViewSheetBox.SelectionBoxItem;
            }
            catch
            {
                MessageBox.Show("Выберите семейство листа");
            }

            if (fs != null)
            {
                if (GridView1.Columns[0].GetCellContent(GridView1.Items[0]) is TextBlock textBlock && !string.IsNullOrEmpty(textBlock.Text))
                {
                    foreach (var gridView1Item in GridView1.Items)
                    {
                        if (GridView1.Columns[0].GetCellContent(gridView1Item) is TextBlock x && !string.IsNullOrEmpty(x.Text))
                            ViewSheet.Create(_doc, fs.Id);
                    }

                    isSheetCreated = true;
                }
                else
                {
                    MessageBox.Show(
                        "Загрузите наименования для элементов, которые вы хотите создать.\n\n"
                        + "Для загрузки элементов по-умолчанию нажмите \"Загрузить стандартные\"\n. Для загрузки файла Excel нажмите \"Загрузить Excel\"");
                }
            }

            tran.Commit();

            if (isSheetCreated)
            {
                tran.Start("Rename Sheets");

                FilteredElementCollector collector = new FilteredElementCollector(_doc);
                collector.OfClass(typeof(ViewSheet));
                var allSheets = collector.ToElements().Cast<ViewSheet>().ToList();

                // set sheet numbers
                for (int i = 0; i < allSheets.Count; i++)
                {
                    if (GridView1.Columns[0].GetCellContent(GridView1.Items[i]) is TextBlock x && !string.IsNullOrEmpty(x.Text))
                        allSheets[i].SheetNumber = x.Text;

                    if (GridView1.Columns[1].GetCellContent(GridView1.Items[i]) is TextBlock y && !string.IsNullOrEmpty(y.Text))
                        allSheets[i].Name = y.Text;
                }

                tran.Commit();
                MessageBox.Show("Листы созданы");
            }
        }

        private void CreateWorkSetsBtn_Click(object sender, RoutedEventArgs e)
        {
            if (GridView1.Columns[0].GetCellContent(GridView1.Items[0]) is TextBlock y && !string.IsNullOrEmpty(y.Text))
            {
                if (_doc.IsWorkshared)
                {
                    // Worksets can only be created in a document with worksharing enabled
                    using (Transaction worksetTransaction = new Transaction(_doc, "Создание рабочих наборов"))
                    {
                        worksetTransaction.Start();
                        foreach (var gridViewItem in GridView1.Items)
                        {
                            // Workset name must not be in use by another workset
                            if (GridView1.Columns[0].GetCellContent(gridViewItem) is TextBlock x && !string.IsNullOrEmpty(x.Text) && WorksetTable.IsWorksetNameUnique(_doc, x.Text))
                            {
                                Workset.Create(_doc, x.Text);
                            }
                        }

                        worksetTransaction.Commit();
                    }

                    MessageBox.Show("Рабочие наборы созданы");
                }
                else
                {
                    MessageBox.Show(
                        "Worksets can only be created in a document with worksharing enabled\n\n"
                        + "(Рабочие наборы могут быть созданы только в документе с включенной совместной работой)");
                }
            }
            else
            {
                MessageBox.Show(
                    "Загрузите наименования для элементов, которые вы хотите создать.\n\n"
                    + "Для загрузки элементов по-умолчанию нажмите \"Загрузить стандартные\"\n. Для загрузки файла Excel нажмите \"Загрузить Excel\"");
            }
        }

        private void ViewTypesCreateBtn_Click(object sender, RoutedEventArgs e)
        {
            ViewFamily viewType = ViewFamily.StructuralPlan;
            string transName = "Создание типов планов несущих конструкций";
            string completeMessage = "Типы планов несущих конструкций созданы";

            if (PlansRadioBtn.IsChecked == false)
            {
                viewType = ViewFamily.Section;
                transName = "Создание типов разрезов";
                completeMessage = "Типы разрезов созданы";
            }

            if (GridView1.Columns[0].GetCellContent(GridView1.Items[0]) is TextBlock y && !string.IsNullOrEmpty(y.Text))
            {
                IEnumerable<ViewFamilyType> viewFamilyTypes = from elem in new FilteredElementCollector(_doc).OfClass(typeof(ViewFamilyType))
                                                              let type = elem as ViewFamilyType
                                                              where type.ViewFamily == viewType
                                                              select type;

                using (Transaction viewFamilyTypesTransaction = new Transaction(_doc, transName))
                {
                    viewFamilyTypesTransaction.Start();

                    foreach (var gridViewItem in GridView1.Items)
                    {
                        if (GridView1.Columns[0].GetCellContent(gridViewItem) is TextBlock x && !string.IsNullOrEmpty(x.Text))
                        {
                            viewFamilyTypes.FirstOrDefault()?.Duplicate(x.Text);
                        }
                    }

                    viewFamilyTypesTransaction.Commit();
                }

                MessageBox.Show(completeMessage);
            }
            else
            {
                MessageBox.Show(
                    "Загрузите наименования для элементов, которые вы хотите создать.\n\n"
                    + "Для загрузки элементов по-умолчанию нажмите \"Загрузить стандартные\"\n. Для загрузки файла Excel нажмите \"Загрузить Excel\"");
            }
        }

        private void FilterDeleteBtn_Click(object sender, RoutedEventArgs e)
        {
            using (Transaction t = new Transaction(_doc, "Удаление неиспользуемых фильтров"))
            {
                t.Start();

                _doc.Delete(ViewFilters.GetUnUsedFilterIds(_doc));

                t.Commit();
                MessageBox.Show("Неиспользуемые фильтры удалены");
            }

            var docFiltersNames = ViewFilters.GetDocFilters(_doc).Select(f => new MyRow { Col1 = f.Name, Col2 = f.Id.ToString() }).ToList();
            GridView1.ItemsSource = docFiltersNames;
        }

        private void ShowFiltersBtn_Click(object sender, RoutedEventArgs e)
        {
            var docFiltersNames = ViewFilters.GetDocFilters(_doc).Select(f => new MyRow { Col1 = f.Name, Col2 = f.Id.ToString() }).ToList();
            GridView1.ItemsSource = docFiltersNames;
            GridView1.Columns[0].Header = "Фильтры";
            GridView1.Columns[1].Header = "Id";

            DataGridTextColumn textColumn = new DataGridTextColumn { Header = "Статус", Binding = new System.Windows.Data.Binding("State") };
            GridView1.Columns.Add(textColumn);

            var unUsedFilterIds = ViewFilters.GetUnUsedFilterIds(_doc);

            foreach (var item in GridView1.ItemsSource)
            {
                MyRow row = (MyRow)item;
                foreach (var unUsedFilterId in unUsedFilterIds)
                {
                    if (row.Col2.Equals(unUsedFilterId.ToString()))
                    {
                        row.State = State.UnUsedFilter;
                    }
                }
            }
        }
    }
}