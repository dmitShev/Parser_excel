using System;
using System.Collections.Generic;
using System.Linq;
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
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Data;
using System.Diagnostics;

namespace Parser_excel
{
    public partial class MainWindow : Window
    {
        public static OpenFileDialog openfile;
        public static bool? browsefile;
        public static bool? browsefile1;
        public static bool? browsefile2;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void buttonDownload_Click(object sender, RoutedEventArgs e)
        {
            openfile = new OpenFileDialog
            {
                DefaultExt = ".xlsx",
                Filter = "(.xlsx)|*.xlsx"
            };
            browsefile = openfile.ShowDialog();
            MessageBox.Show("Загрузка завершена!");
        }

        private void buttonClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Up_Click(object sender, RoutedEventArgs e)
        {
            scroll.LineUp();
        }

        private void Down_Click(object sender, RoutedEventArgs e)
        {
            scroll.LineDown();
        }

        private void buttonShortOutput_Click(object sender, RoutedEventArgs e)
        {
            txtFilePath.Text = openfile.FileName;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelBook = excelApp.Workbooks.Open(txtFilePath.Text.ToString(), 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet excelSheet = (Excel.Worksheet)excelBook.Worksheets.get_Item(1);
            Excel.Range excelRange = excelSheet.UsedRange;
            string strCellData = "";
            double douCellData;
            int rowCnt = 0;
            int colCnt = 0;
            DataTable dt = new DataTable();
            for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
            {
                if(colCnt > 2)
                {
                    continue;
                }
                string strColumn = "";
                strColumn = (string)(excelRange.Cells[1, colCnt] as Excel.Range).Value2;
                dt.Columns.Add(strColumn, typeof(string));
            }
            for (rowCnt = 2; rowCnt <= excelRange.Rows.Count; rowCnt++)
            {
                string strData = "";
                for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                {
                    try
                    {
                        if(colCnt <= 2)
                        {
                            strCellData = (string)(excelRange.Cells[rowCnt, colCnt] as Excel.Range).Value2;
                            strData = "Уби." + strData + strCellData + "|";
                        }
                    }
                    catch 
                    {
                        douCellData = (excelRange.Cells[rowCnt, colCnt] as Excel.Range).Value2;
                        strData += douCellData.ToString() + "|";
                    }
                }
                if(rowCnt == 2)
                {
                    strData = strData.Remove(0, 8);
                }
                
                strData = strData.Remove(strData.Length - 1, 1);
                dt.Rows.Add(strData.Split('|'));
            }
            
            dtGrid.ItemsSource = dt.DefaultView;

            excelBook.Close(true, null, null);
            excelApp.Quit();
            MessageBox.Show("Загрузка кратких сведений выполнена!");
        }

        private void buttonFullOutput_Click(object sender, RoutedEventArgs e)
        {
            if (browsefile == true)
            {
                txtFilePath.Text = openfile.FileName;
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelBook = excelApp.Workbooks.Open(txtFilePath.Text.ToString(), 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet excelSheet = (Excel.Worksheet)excelBook.Worksheets.get_Item(1);
                Excel.Range excelRange = excelSheet.UsedRange;

                string strCellData = "";
                double douCellData;
                int rowCnt = 0;
                int colCnt = 0;

                DataTable dt = new DataTable();
                for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                {
                    string strColumn = "";
                    strColumn = (string)(excelRange.Cells[1, colCnt] as Excel.Range).Value2;
                    dt.Columns.Add(strColumn, typeof(string));
                }

                for (rowCnt = 2; rowCnt <= excelRange.Rows.Count; rowCnt++)
                {
                    string strData = "";
                    for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                    {
                        try
                        {
                            strCellData = (string)(excelRange.Cells[rowCnt, colCnt] as Excel.Range).Value2;
                            strData += strCellData + "|";
                        }
                        catch 
                        {
                            douCellData = (excelRange.Cells[rowCnt, colCnt] as Excel.Range).Value2;
                            strData += douCellData.ToString() + "|";
                        }
                    }
                    strData = strData.Remove(strData.Length - 1, 1);
                    dt.Rows.Add(strData.Split('|'));
                }

                dtGrid.ItemsSource = dt.DefaultView;

                excelBook.Close(true, null, null);
                excelApp.Quit();
                MessageBox.Show("Загрузка полных сведений выполнена!");
            }
        }

        private void buttonResfresh_Click(object sender, RoutedEventArgs e)
        {
            //string url = @"https://bdu.fstec.ru/files/documents/thrlist.xlsx";


            txtFilePath.Text = openfile.FileName;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelBook = excelApp.Workbooks.Open(txtFilePath.Text.ToString(), 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet excelSheet = (Excel.Worksheet)excelBook.Worksheets.get_Item(1);
            Excel.Range excelRange = excelSheet.UsedRange;
            Excel.Application excelApp1;
            Excel.Workbook excelBook1;
            Excel.Worksheet excelSheet1;
            Excel.Range excelRange1;
            OpenFileDialog openfile1;
            bool? browsefile1;
            openfile1 = new OpenFileDialog
            {
                DefaultExt = ".xlsx",
                Filter = "(.xlsx)|*.xlsx"
            };
            browsefile1 = openfile1.ShowDialog();
            txtFilePath.Text = openfile1.FileName;
            excelApp1 = new Excel.Application();
            excelBook1 = excelApp1.Workbooks.Open(txtFilePath.Text.ToString(), 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            excelSheet1 = (Excel.Worksheet)excelBook1.Worksheets.get_Item(1);
            excelRange1 = excelSheet1.UsedRange;
            MessageBox.Show("Загрузка файла прошла успешно!");
            if (browsefile == true)
            {
                string strCellData;
                double douCellData;
                int rowCnt;
                int colCnt;
                string strCellData1;
                double douCellData1;
                int colCnt1;
                DataTable dt = new DataTable();
                for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                {
                    string strColumn = "";
                    strColumn = (string)(excelRange.Cells[1, colCnt] as Excel.Range).Value2;
                    dt.Columns.Add(strColumn, typeof(string));
                }
                for (rowCnt = 2; rowCnt <= excelRange.Rows.Count; rowCnt++)
                {
                    string strData = "";
                    for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                    {
                        try
                        {
                            strCellData = (string)(excelRange.Cells[rowCnt, colCnt] as Excel.Range).Value2;
                            strData += strCellData + "|";
                        }
                        catch
                        {
                            douCellData = (excelRange.Cells[rowCnt, colCnt] as Excel.Range).Value2;
                            strData += douCellData.ToString() + "|";
                        }
                    }
                    strData = strData.Remove(strData.Length - 1, 1);

                    string strData1 = "";
                    for (colCnt1 = 1; colCnt1 <= excelRange1.Columns.Count; colCnt1++)
                    {
                        try
                        {
                            strCellData1 = (string)(excelRange1.Cells[rowCnt, colCnt1] as Excel.Range).Value2;
                            strData1 += strCellData1 + "|";
                        }
                        catch
                        {
                            douCellData1 = (excelRange1.Cells[rowCnt, colCnt1] as Excel.Range).Value2;
                            strData1 += douCellData1.ToString() + "|";
                        }
                    }
                    strData1 = strData1.Remove(strData1.Length - 1, 1);

                    if (strData != strData1)
                    {
                        strData = "Было:" + strData;
                        strData1 = "Стало:" + strData1;
                        dt.Rows.Add(strData.Split('|'));
                        dt.Rows.Add(strData1.Split('|'));
                    }
                }
                dtGrid.ItemsSource = dt.DefaultView;
                excelBook.Close(true, null, null);
                excelApp.Quit();
                excelBook1.Close(true, null, null);
                excelApp1.Quit();
                MessageBox.Show("Сравнение файлов прошло успешно!");
            }
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            dtGrid.SelectAllCells();
            dtGrid.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;
            ApplicationCommands.Copy.Execute(null, dtGrid);
            String resultat = (string)Clipboard.GetData(DataFormats.CommaSeparatedValue);
            String result = (string)Clipboard.GetData(DataFormats.Text);
            dtGrid.UnselectAllCells();
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel file (*.xls)|*.xls|Text file (*.txt)|*.txt";
            if (saveFileDialog.ShowDialog() == true)
            {
                StreamWriter file = new StreamWriter(saveFileDialog.FileName, true, Encoding.UTF8);
                file.WriteLine(result.Replace(',', ' '));
                file.Close();
            }
            MessageBox.Show(@"Экспорт прошёл, файл можете найти здесь: " + saveFileDialog.FileName);
        }


    }
}
