using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Windows;
using Microsoft.Win32;
using System.IO;
using System.Windows.Forms;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using MessageBox = System.Windows.MessageBox;

namespace Dashboard1
{
    public partial class MainWindow : Window
    {
        string templateWord, dataExcel;
        string folderPath;

        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void executeButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var wordApp = new Word.Application();
                wordApp.Visible = false;
                for (int i = 0; i < 200; i++)
                {
                    string[] tmp = new string[6];
                    tmp = readExcel(i);

                    var wordDocument = wordApp.Documents.Add(templateWord);
                    if (tmp[0] == "")
                    {
                        break;
                    }
                    else
                    {
                        ReplaceWordStub("<first>", tmp[0], wordDocument);
                        ReplaceWordStub("<second>", tmp[1], wordDocument);
                        ReplaceWordStub("<three>", tmp[2], wordDocument);
                        ReplaceWordStub("<four>", tmp[3], wordDocument);
                        ReplaceWordStub("<five>", tmp[4], wordDocument);
                        ReplaceWordStub("<seven>", tmp[5], wordDocument);

                        wordDocument.SaveAs($"{templateWord}{wordNameTemplate.Text}{tmp[0]}{i}.docx");
                        wordDocument.Close();
                    }
                }
                MessageBox.Show("Files are created!");

                wordApp.Quit();
            }
            catch
            {
                MessageBox.Show("Fill in all the fields!");
            }
        }

        private void ReplaceWordStub(string stubToReplace, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
        }

        private string[] readExcel(int index)
        {
            index = index + 2;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(dataExcel, 0, true, 5, "", "", true);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            string[] data = new string[7];
            data[0] = xlWorkSheet.get_Range("A" + index.ToString()).Text;
            data[1] = xlWorkSheet.get_Range("B" + index.ToString()).Text;
            data[2] = xlWorkSheet.get_Range("C" + index.ToString()).Text;
            data[3] = xlWorkSheet.get_Range("D" + index.ToString()).Text;
            data[4] = xlWorkSheet.get_Range("E" + index.ToString()).Text;
            data[5] = xlWorkSheet.get_Range("G" + index.ToString()).Text;

            data[6] = xlWorkSheet.get_Range("H" + index.ToString()).Text;

            xlWorkBook.Close(false);
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            return data;
        }

        private void dataButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = " (*.xls)|*.xls";
            if (openFileDialog1.ShowDialog() == true)
            {
                dataExcel = openFileDialog1.FileName;
            }
            else
            {
                MessageBox.Show("Select data!");
            }
        }

        private void exitButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void templateButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = " (*.docx)|*.docx";
            if (openFileDialog2.ShowDialog() == true)
            {
                templateWord = openFileDialog2.FileName;
            }
            else
            {
                MessageBox.Show("Select template!");
            }
        }
    }
}
