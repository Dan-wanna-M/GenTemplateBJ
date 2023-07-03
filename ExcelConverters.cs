using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace GenTemplateBJ
{
    internal class ExcelConverters: INotifyPropertyChanged
    {
        public List<string> TemplateTypes { get; } = new() {"请选择", "川西" };
        public Dictionary<string, Action> TemplateTypeToExcelConverter { get; } = new();

        public ExcelConverters()
        {
            TemplateTypeToExcelConverter["川西"] = () =>
            {
                OutputExcels = new();
                var result = FillTransportList("川西");
                OutputExcels.Add("发货清单.xlsx", result);
            };
        }

        private XLWorkbook FillTransportList(string templateType)
        {
            var folderPath = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName;
            // Console.WriteLine("Folder Path: " + folderPath);
            var transportList = new XLWorkbook(Path.Combine(folderPath, "Templates", templateType, "发货清单.xlsx"));
            var worksheet = transportList.Worksheet(1);
            worksheet.Cell(1, "C").Value = excelData.OneToOneData["工程名字"];
            worksheet.Cell(2, "C").Value = $"材料单({excelData.OneToOneData["工程类别"]})";
            worksheet.Cell(1, "G").Value = $"装箱单号: {excelData.OneToOneData["总箱数量"]}";
            worksheet.Cell(3, "C").Value = excelData.OneToOneData["货物名称"];
            worksheet.Cell(4, "C").Value = excelData.OneToOneData["合同编号"];
            worksheet.Cell(5, "C").Value = excelData.OneToOneData["请购单号"];
            worksheet.Cell(6, "C").Value = excelData.OneToOneData["发货日期"];
            worksheet.Cell(7, "C").Value = excelData.OneToOneData["到货地点"];
            worksheet.Cell(3, "F").Value = excelData.OneToOneData["发货单位"];
            worksheet.Cell(4, "F").Value = excelData.OneToOneData["发货人 电话"];
            worksheet.Cell(5, "F").Value = excelData.OneToOneData["收货人 电话"];
            worksheet.Cell(6, "F").Value = excelData.OneToOneData["承运商"];
            worksheet.Cell(7, "F").Value = excelData.OneToOneData["运输方式"];
            worksheet.Row(9).InsertRowsBelow(excelData.OneToManyData["材料编码/设备位号"].Count);
            var end = 10 + excelData.OneToManyData["材料编码/设备位号"].Count;
            for (int i = 10; i < end; i++)
            {
                worksheet.Cell(i, "A").Value = i - 9;
                int j = i - 10;
                worksheet.Cell(i, "B").Value = excelData.OneToManyData["材料编码/设备位号"][j];
                worksheet.Cell(i, "C").Value = excelData.OneToManyData["产品规格(Size)"][j];
                worksheet.Cell(i, "F").Value = excelData.OneToManyData["单位（Unit）"][j];
                worksheet.Cell(i, "G").Value = excelData.OneToManyData["数量（Quantity）"][j];
                // worksheet.Cell(i, "H").Value = excelData.OneToManyData["箱号"][j];
                // worksheet.Cell(i, "I").Value = excelData.OneToManyData["备注（跟踪号）"][j];
            }
            worksheet.Cell(end, "F").Value = excelData.OneToManyData["数量（Quantity）"].Select(x => int.Parse(x)).Sum();
            return transportList;
        }

        private InputExcelData? excelData;
        public InputExcelData? ExcelData { get=>excelData; set
            {
                excelData = value;
                OnPropertyChanged(nameof(ExcelData));
                OnPropertyChanged(nameof(IsExcelDataNotNull));
            } }

        private Dictionary<string, XLWorkbook>? outputExcels;
        public Dictionary<string, XLWorkbook>? OutputExcels
        {
            get => outputExcels; set
            {
                outputExcels = value;
                OnPropertyChanged(nameof(OutputExcels));
                OnPropertyChanged(nameof(IsOutputExcelsNotNull));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public bool IsExcelDataNotNull { get => ExcelData != null; }

        public bool IsOutputExcelsNotNull { get => OutputExcels != null; }

        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
