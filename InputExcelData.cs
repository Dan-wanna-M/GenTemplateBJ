using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
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
        public Dictionary<string, XLCellValue[]> OneToManyData { get;} = new();

        public InputExcelData(XLWorkbook workbook)
        {
            var worksheet = workbook.Worksheet(1);
            var oneToOneKeys = worksheet.Column(1).CellsUsed().ToList();
            var oneToOneValues = worksheet.Column(2).CellsUsed().ToList();
            for (int i = 0; i < oneToOneKeys.Count; i++)
            {
                if (i < oneToOneValues.Count)
                    OneToOneData.Add(oneToOneKeys[i].Value.ToString().Trim(), oneToOneValues[i].Value.ToString().Trim());
                else
                    OneToOneData.Add(oneToOneKeys[i].Value.ToString().Trim(), "");
            }
            var oneToManyData = worksheet.Columns().Where(x => x.ColumnNumber() is not (1 or 2)).Select(x=>x.CellsUsed().ToList());
            foreach(var cells in oneToManyData)
            {
                var values = new XLCellValue[oneToManyData.First(x => x[0].Value.ToString().Trim() == "材料编码/设备位号").Count - 1];
                Array.Fill(values, "");
                var temp = cells.Skip(1).Select(x=>x.Value).ToArray();
                Array.Copy(temp, values, temp.Length);
                OneToManyData.Add(cells.First().Value.ToString().Trim(), values);
            }
        }
    }
}
