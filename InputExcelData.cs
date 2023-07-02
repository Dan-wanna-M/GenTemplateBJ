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
            var worksheet = workbook.Worksheet(1);
            var oneToOneKeys = worksheet.Column(1).CellsUsed().ToList();
            var oneToOneValues = worksheet.Column(2).CellsUsed().ToList();
            if (oneToOneKeys.Count != oneToOneValues.Count)
                throw new ArgumentException($"一对一的数据字段和数据内容的数量不匹配({oneToOneKeys.Count}:{oneToOneValues.Count})");
            for (int i = 0; i < oneToOneKeys.Count; i++)
            {
                OneToOneData.Add(oneToOneKeys[i].Value.ToString().Trim(), oneToOneValues[i].Value.ToString().Trim());
            }
            var oneToManyData = worksheet.Columns().Where(x => (x.ColumnNumber() is not (1 or 2)));
            foreach(var column in oneToManyData) 
            {
                var cells = column.CellsUsed().ToList();
                var first = cells.First();
                OneToManyData.Add(first.Value.ToString().Trim(),
                    cells
                    .Where(x => first.Value.ToString() != x.Value.ToString())
                    .Select(x=>x.Value.ToString().Trim())
                    .ToList());
            }
        }
    }
}
