using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
// Проект -> добавит ссылку на модель СОМ -> Microsoft Excel Object Library , Microsoft Word Object Library
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using static System.Net.Mime.MediaTypeNames;

// копирование excel таблицы в word
namespace ExcelToWord
{
    class Program
    {
        static void Main(string[] args)
        {

            string wbkName = AppDomain.CurrentDomain.BaseDirectory + @"пример.xlsx"; //можно просто указать полный путь до файла

            Excel._Application xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            Excel.Workbook workbook = xlApp.Workbooks.Open(wbkName);
            Excel.Sheets sheets = workbook.Worksheets;

            Word._Application wdApp = new Word.Application();
            wdApp.Visible = false;
            wdApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            Word.Document document = wdApp.Documents.Add();
            Word.Paragraph paragraph;
           

            // Get pages count
            int xlPagesCount = workbook.Sheets.Count;
            
            // итерируюсь по всем эксель листам и вставляю в ворд в начале название экслеь листа а потом саму таблицу с этого листа
            for (int i = 0; i < xlPagesCount; i++)
            {

                Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(i);
                string strWorksheetName = worksheet.Name;
                worksheet = (Excel.Worksheet)workbook.Sheets[i];

                paragraph = document.Paragraphs.Add();
                paragraph.Range.Text = strWorksheetName;

                worksheet.UsedRange.Copy();
                paragraph = document.Paragraphs.Add();
                paragraph.Range.PasteSpecial();
            }
            

            workbook.Close();
            xlApp.Quit();
            document.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "Result.docx");
            document.Close();
            wdApp.Quit();
        }
    }
}