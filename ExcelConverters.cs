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
                var result1 = FillTransportList("川西");
                OutputExcels.Add("发货清单.xlsx", result1);
                var result2 = FillQualityList("川西");
                OutputExcels.Add("质检报告.xlsx", result2);

            };
        }

        private XLWorkbook FillTransportList(string templateType)
        {
            var folderPath = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName;
            // Console.WriteLine("Folder Path: " + folderPath);
            var transportList = new XLWorkbook(Path.Combine(folderPath, "Templates", templateType, "发货清单.xlsx"));
            var worksheet = transportList.Worksheet(1);
            worksheet.Cell(1, "C").Value = excelData.OneToOneData["项目名称"];
            //worksheet.Cell(2, "C").Value = $"材料单({excelData.OneToOneData["工程类别"]})";
            worksheet.Cell(1, "G").Value = $"装箱单号: {excelData.OneToOneData["总箱数量"]}";
            worksheet.Cell(3, "C").Value = excelData.OneToOneData["材料名称"];
            worksheet.Cell(4, "C").Value = excelData.OneToOneData["合同号"];
            worksheet.Cell(5, "C").Value = excelData.OneToOneData["请购单号"];
            worksheet.Cell(6, "C").Value = excelData.OneToOneData["发货日期"];
            worksheet.Cell(7, "C").Value = excelData.OneToOneData["到货地点"];
            worksheet.Cell(3, "F").Value = excelData.OneToOneData["公司名称"];
            worksheet.Cell(4, "F").Value = excelData.OneToOneData["发货人 电话"];
            worksheet.Cell(5, "F").Value = excelData.OneToOneData["收货人 电话"];
            worksheet.Cell(6, "F").Value = excelData.OneToOneData["承运商"];
            worksheet.Cell(7, "F").Value = excelData.OneToOneData["运输方式"];
            worksheet.Row(9).InsertRowsBelow(excelData.OneToManyData["材料编码/设备位号"].Length);
            var end = 10 + excelData.OneToManyData["材料编码/设备位号"].Length;
            for (int i = 10; i < end; i++)
            {
                worksheet.Cell(i, "A").Value = i - 9;
                int j = i - 10;
                worksheet.Cell(i, "B").Value = excelData.OneToManyData["材料编码/设备位号"][j];
                worksheet.Cell(i, "C").Value = excelData.OneToManyData["产品规格(Size)"][j];
                worksheet.Cell(i, "F").Value = excelData.OneToManyData["单位（Unit）"][j];
                worksheet.Cell(i, "G").Value = excelData.OneToManyData["数量（Quantity）"][j];
                worksheet.Cell(i, "H").Value = excelData.OneToManyData["箱号"][j];
                worksheet.Cell(i, "I").Value = excelData.OneToManyData["备注（跟踪号）"][j];
            }
            worksheet.Cell(end, "G").Value = excelData.OneToManyData["数量（Quantity）"].Select(x => (int)x.GetUnifiedNumber()).Sum();
            return transportList;
        }

        private XLWorkbook FillQualityList(string templateType)
        {
            var folderPath = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName;
            // Console.WriteLine("Folder Path: " + folderPath);
            var qualityList = new XLWorkbook(Path.Combine(folderPath, "Templates", templateType, "质检文件.xlsx"));
            var worksheet = qualityList.Worksheet(1);
            worksheet.Cell(3, "A").Value = string.Format("报告编号: TJMZLBG-yyyymm-{0}", excelData.OneToOneData["质检报告编号"]);
            worksheet.Cell(4, "B").Value = excelData.OneToOneData["公司名称"];
            worksheet.Cell(5, "B").Value = excelData.OneToOneData["项目名称"];
            worksheet.Cell(6, "B").Value = excelData.OneToOneData["依据标准"];
            worksheet.Cell(7, "B").Value = excelData.OneToOneData["使用部分"];
            worksheet.Cell(9, "B").Value = excelData.OneToOneData["材料名称"];
            worksheet.Row(8).InsertRowsBelow(excelData.OneToManyData["材料编码/设备位号"].Length);
            worksheet.Cell(9, "A").Value = excelData.OneToOneData["材料名称"];
            var end1 = 9 + excelData.OneToManyData["材料编码/设备位号"].Length;
            for (int i = 9; i < end1; i++)
            {
                int j = i - 9;
                worksheet.Cell(i, "B").Value = excelData.OneToManyData["材料编码/设备位号"][j];
                worksheet.Cell(i, "C").Value = excelData.OneToManyData["产品规格(Size)"][j];
                worksheet.Cell(i, "D").Value = excelData.OneToManyData["数量（Quantity）"][j];
                worksheet.Cell(i, "E").Value = excelData.OneToManyData["单位（Unit）"][j];
                worksheet.Cell(i, "F").Value = excelData.OneToManyData["生产负责人"][j];
            }
            worksheet.Row(end1 - 1).InsertRowsBelow(excelData.OneToManyData["试验项目"].Length - 1);
            var end2 = end1 + excelData.OneToManyData["试验项目"].Length;
            for (int i = end1; i < end2; i++)
            {
                int j = i - end1;
                worksheet.Cell(i, "B").Value = excelData.OneToManyData["试验项目"][j];
                worksheet.Cell(i, "C").Value = excelData.OneToManyData["标准值"][j];
            }

                return qualityList;

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
