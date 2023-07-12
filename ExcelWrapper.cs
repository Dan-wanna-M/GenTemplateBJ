using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace GenTemplateBJ
{
    internal class ExcelWrapper
    {
        public IXLWorkbook Workbook {get;private set; }
        public IXLWorksheet[] ActiveWorkSheets { get;private set; }
        public ExcelWrapper(IXLWorkbook workbook, params IXLWorksheet[] activeWorkSheets)
        {
            Workbook = workbook;
            ActiveWorkSheets = activeWorkSheets;
        }
    }
}
