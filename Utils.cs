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
    }
}
