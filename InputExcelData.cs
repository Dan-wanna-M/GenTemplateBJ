using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ClosedXML.Excel;

namespace GenTemplateBJ
{
    internal class InputExcelData
    {
        public Dictionary<string, string> OneToOneData { get; } = new();
        public Dictionary<string, List<string>> OneToManyData { get;} = new();

        public InputExcelData(XLWorkbook workbook)
        {
            var oneToOneKeys = workbook.FindColumns(x=>x.ColumnNumber()==1).Single().CellsUsed().ToList();
            var oneToOneValues = workbook.FindColumns(x => x.ColumnNumber() == 2).CellsUsed().ToList();
            if (oneToOneKeys.Count != oneToOneValues.Count)
                throw new ArgumentException($"一对一的数据字段和数据内容的数量不匹配({oneToOneKeys.Count}:{oneToOneValues.Count})");
            for (int i = 0; i < oneToOneKeys.Count; i++)
            {
                OneToOneData.Add(oneToOneKeys[i].Value.ToString(), oneToOneValues[i].Value.ToString());
            }
            var oneToManyData = workbook.FindColumns(x => (x.ColumnNumber() is not (1 or 2)));
            foreach(var column in oneToManyData) 
            {
                var cells = column.CellsUsed();
                var first = cells.First();
                OneToManyData.Add(first.Value.ToString(),
                    cells
                    .Where(x => first != x)
                    .Select(x=>x.Value.ToString())
                    .ToList());
            }
        }
    }
}
