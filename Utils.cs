using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.Win32;
using System.Windows;
using System.IO;
using Microsoft.WindowsAPICodePack.Dialogs;
using DocumentFormat.OpenXml.Spreadsheet;

namespace GenTemplateBJ
{
    internal static class Utils
    {
        public static XLWorkbook? OpenAnExcelFile()
        {
            var fd = new OpenFileDialog()
            {
                Multiselect = false,
                DefaultExt = ".xlsx",
                Filter = "Excel files (*.xlsx)|*.xlsx"
            };
            if (fd.ShowDialog() is not null and true)
            {
                try
                {
                    XLWorkbook? workbook = new(fd.OpenFile());
                    return workbook;
                }
                catch(IOException e)
                {
                    MessageBox.Show(e.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            return null;
        }

        public static string? OpenAFolder()
        {
            CommonOpenFileDialog dialog = new()
            {
                InitialDirectory = Environment.CurrentDirectory,
                IsFolderPicker = true
            };
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                return dialog.FileName;
            }
            return null;
        }

        public static IEnumerable<IXLCell> GetCellsUntilLastCellUsed<T>(T data) where T: IXLRangeBase
        {
            return data.Cells($"1:{data.LastCellUsed().Address.RowNumber}");
        }

        public static XLWorkbook GetTemplateExcel(string templateType, string name)
        {
            var folderPath = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName;
            return new XLWorkbook(Path.Combine(folderPath, "Templates", templateType, name));
        }

        public static int ColumnLetterToNumber(this IXLWorksheet sheet, string letter)
        {
            return sheet.Column(letter).ColumnNumber();
        }

        public static string ColumnNumberToLetter(this IXLWorksheet sheet, int number)
        {
            return sheet.Column(number).ColumnLetter();
        }
    }
}
