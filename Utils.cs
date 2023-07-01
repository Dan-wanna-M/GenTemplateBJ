using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.Win32;
using System.Windows;

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
                XLWorkbook? workbook = new(fd.OpenFile());
                return workbook;
            }
            return null;
        }
    }
}
