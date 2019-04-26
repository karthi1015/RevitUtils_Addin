#region Namespaces

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
#endregion

namespace RevitUtils
{
    /// <summary>
    /// My customized model
    /// </summary>
    public class MyRow
    {
        public string Col1 { get; set; }
        public string Col2 { get; set; }
    }

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

            //get titleblock (family types of viewsheets)
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
                //open the Excel file and get the specified worksheet
                int rCnt;

                _xlApp = new Excel.Application();
                //open the excel
                _xlWorkBook = _xlApp.Workbooks.Open(FilePathTextBox.Text);
                //get the first sheet of the excel
                _xlWorkSheet = (Excel.Worksheet)_xlWorkBook.Worksheets.Item[1];

                _range = _xlWorkSheet.UsedRange;
                // get the total row count
                _rowCount = _range.Rows.Count;
                //get the total column count
                _colCount = _range.Columns.Count;

                var myRows = new List<MyRow>();

                if (_colCount < 2)
                {
                    // traverse all the row in the excel
                    for (rCnt = 1; rCnt <= _rowCount; rCnt++)
                    {
                        var myRow = new MyRow
                        {
                            Col1 = (string)(_range.Cells[rCnt, 1] as Excel.Range)?.Value2.ToString(),
                        };

                        myRows.Add(myRow);
                    }
                }
                else
                {
                    // traverse all the row in the excel
                    for (rCnt = 1; rCnt <= _rowCount; rCnt++)
                    {
                        //traverse columns (the first column is not included)
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

                //release the resources
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
                foreach (var gridView1Item in GridView1.Items)
                {
                    if (GridView1.Columns[0].GetCellContent(gridView1Item) is TextBlock x && !string.IsNullOrEmpty(x.Text))
                        ViewSheet.Create(_doc, fs.Id);
                }

                isSheetCreated = true;
            }

            tran.Commit();

            if (isSheetCreated)
            {
                tran.Start("Rename Sheets");

                FilteredElementCollector collector = new FilteredElementCollector(_doc);
                collector.OfClass(typeof(ViewSheet));
                var allSheets = collector.ToElements().Cast<ViewSheet>().ToList();

                //set sheet numbers
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
            // Worksets can only be created in a document with worksharing enabled
            if (_doc.IsWorkshared)
            {
                using (Transaction worksetTransaction = new Transaction(_doc, "Создание рабочих наборов"))
                {
                    worksetTransaction.Start();
                    foreach (var gridViewItem in GridView1.Items)
                    {
                        // Workset name must not be in use by another workset
                        if (GridView1.Columns[0].GetCellContent(gridViewItem) is TextBlock x &&
                            !string.IsNullOrEmpty(x.Text) && WorksetTable.IsWorksetNameUnique(_doc, x.Text))
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
                MessageBox.Show("Worksets can only be created in a document with worksharing enabled\n\n" +
                                "(Рабочие наборы могут быть созданы только в документе с включенной совместной работой)");
            }
        }

        private void ViewTypesCreateBtn_Click(object sender, RoutedEventArgs e)
        {
            IEnumerable<ViewFamilyType> viewFamilyTypes =
                from elem in new FilteredElementCollector(_doc).OfClass(typeof(ViewFamilyType))
                let type = elem as ViewFamilyType
                where type.ViewFamily == ViewFamily.StructuralPlan
                select type;

            using (Transaction viewFamilyTypesTransaction = new Transaction(_doc, "Создание типов планов несущих конструкций"))
            {
                viewFamilyTypesTransaction.Start();

                viewFamilyTypes.FirstOrDefault()?.Duplicate("new 1");

                viewFamilyTypesTransaction.Commit();
            }
            MessageBox.Show("Типы планов несущих конструкций созданы");
        }
    }
}
