using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using ClosedXML.Utils;

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
                OutputExcels = new()
                {
                    { "发货清单.xlsx", FillTransportList("川西") },
                    { "质检报告.xlsx", FillQualityList("川西") },
                    {"产品合格证.xlsx",FillProductionCertificate("川西") }
                };
            };
        }

        private XLWorkbook FillTransportList(string templateType)
        {
            var transportList=Utils.GetTemplateExcel(templateType, "2送货单模版.xlsx");
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
            var height = worksheet.Row(10).Height;
            foreach (var i in worksheet.Rows(10, 11))
                i.Delete();
            worksheet.Row(9).InsertRowsBelow(excelData.OneToManyData["材料编码/设备位号"].Length);

            var end = 9 + excelData.OneToManyData["材料编码/设备位号"].Length+1;
            for (int i = 10; i < end; i++)
            {
                worksheet.Row(i).Height = height;
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
            var qualityList = Utils.GetTemplateExcel(templateType, "质检文件.xlsx");
            var worksheet = qualityList.Worksheet(1);
            worksheet.Cell(3, "A").Value = $"报告编号: TJMZLBG-yyyymm-{excelData.OneToOneData["质检报告编号"]}";
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
        enum CertificateRowStatus
        {
            Empty,
            LeftFull,
            Full,
        }
        private XLWorkbook FillProductionCertificate(string templateType)
        {
            var productionCertificate = Utils.GetTemplateExcel(templateType, "3产品合格证模版.xlsx");
            var worksheet = productionCertificate.Worksheet(1);
            int certificateWidth = worksheet.ColumnLetterToNumber("W") - worksheet.ColumnLetterToNumber("E")+1;
            int certicateHeight = 28 - 3+1;
            int marginW = 2;
            int marginH = 2;
            var layoutState = CertificateRowStatus.Empty;
            int currentLeft = worksheet.ColumnLetterToNumber("E");
            int currentTop = 3;
            int initialTop = currentTop;
            var logo = worksheet.Pictures.Single();
            void AdjustWidth(int initialLeft, int current, int size)
            {
                for (int i = current; i < current+size; i++)
                {
                    worksheet.Column(i).Width = worksheet.Column(initialLeft + i - current).Width;
                }
            }
            void AdjustHeight(int initialTop, int current, int size)
            {
                for (int i = current; i < current + size; i++)
                {
                    worksheet.Row(i).Height = worksheet.Row(initialTop + i - current).Height;
                }
            }
            AdjustWidth(currentLeft, worksheet.ColumnLetterToNumber("Z"), certificateWidth);
            void AddOneCertificate(XLCellValue productSize, XLCellValue materialCode, XLCellValue quantity)
            {
                int horizontalShift = certificateWidth + marginW;
                int verticalShift = certicateHeight + marginH;
                switch (layoutState)
                {
                    case CertificateRowStatus.Empty:
                        layoutState = CertificateRowStatus.LeftFull;
                        break;
                    case CertificateRowStatus.LeftFull:
                        worksheet.Range(
                            $"{worksheet.ColumnNumberToLetter(currentLeft)}{currentTop}" +
                            $":{worksheet.ColumnNumberToLetter(currentLeft + certificateWidth)}{currentTop + certicateHeight}")
                            .CopyTo(worksheet.Cell(currentTop, currentLeft + horizontalShift));
                        currentLeft += horizontalShift;
                        layoutState = CertificateRowStatus.Full;
                        break;
                    case CertificateRowStatus.Full:
                        worksheet.Range(
                             $"{worksheet.ColumnNumberToLetter(currentLeft)}{currentTop}" +
                             $":{worksheet.ColumnNumberToLetter(currentLeft + certificateWidth)}{currentTop + certicateHeight}")
                             .CopyTo(worksheet.Cell(currentTop+verticalShift, currentLeft - horizontalShift));
                        currentTop += verticalShift;
                        currentLeft -= horizontalShift;
                        layoutState = CertificateRowStatus.LeftFull;
                        break;
                }
                AdjustHeight(currentTop, initialTop, certicateHeight);
                int firstCellVerticalOffset = 10 - 3;
                int firstCellHorizontalOffset = worksheet.ColumnLetterToNumber("L") - worksheet.ColumnLetterToNumber("E");
                logo.Duplicate().MoveTo(worksheet.Cell(currentTop + 1, currentLeft + worksheet.ColumnLetterToNumber("P") - worksheet.ColumnLetterToNumber("E")));
                worksheet.Cell(currentTop + firstCellVerticalOffset, currentLeft+firstCellHorizontalOffset).Value = excelData.OneToOneData["材料名称"];
                worksheet.Cell(currentTop + firstCellVerticalOffset+2, currentLeft + firstCellHorizontalOffset).Value = productSize;
                worksheet.Cell(currentTop + firstCellVerticalOffset + 2+3, currentLeft + firstCellHorizontalOffset).Value = materialCode;
                worksheet.Cell(currentTop + firstCellVerticalOffset + 2+3+3, currentLeft + firstCellHorizontalOffset).Value = quantity;
                worksheet.Cell(currentTop + firstCellVerticalOffset + 2 + 3 + 3+2, currentLeft + firstCellHorizontalOffset).Value = excelData.OneToOneData["出厂日期"];
            }
            for (int i = 0; i < excelData.OneToManyData["材料编码/设备位号"].Length; i++)
            {
                AddOneCertificate(excelData.OneToManyData["产品规格(Size)"][i], excelData.OneToManyData["材料编码/设备位号"][i], excelData.OneToManyData["数量（Quantity）"][i]);
            }
            return productionCertificate;
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
