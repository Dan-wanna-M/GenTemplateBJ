﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ClosedXML.Excel;
using NPOI.XWPF.UserModel;
using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.WindowsAPICodePack.Shell.Interop;
using DocumentFormat.OpenXml.Wordprocessing;

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
                    { "各厂家自查表.xlsx", FillSelfCheckTable("川西") },
                    {"产品合格证.xlsx",FillProductionCertificate("川西") },
                    {"放行报告.xlsx", FillReleaseReport("川西") }
                };
                OutputDocxs = new()
                {
                    {"封面.docx", FillCoverPage("川西") }
                };
            };
        }

        private XLWorkbook FillTransportList(string templateType)
        {
            var transportList=Utils.GetTemplateExcel(templateType, "2送货单模版.xlsx");
            var worksheet = transportList.Worksheet(1);
            worksheet.Cell(1, "C").Value = excelData.OneToOneData["项目名称"]+ExcelData.OneToOneData["使用部分"];
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
            var qualityList = Utils.GetTemplateExcel(templateType, "4质检报告模版.xlsx");
            var worksheet = qualityList.Worksheet("检验报告-02804-01-4000-MP-R-M-8050");
            worksheet.Cell(3, "A").Value = $"报告编号: TJMZLBG-yyyymm-{excelData.OneToOneData["质检报告编号"]}";
            worksheet.Cell(4, "B").Value = excelData.OneToOneData["公司名称"];
            worksheet.Cell(5, "B").Value = excelData.OneToOneData["项目名称"]+excelData.OneToOneData["使用部分"];
            worksheet.Cell(6, "B").Value = excelData.OneToOneData["依据标准"];
            worksheet.Cell(7, "B").Value = excelData.OneToOneData["使用部分"];

            worksheet.Row(8).InsertRowsBelow(excelData.OneToManyData["材料编码/设备位号"].Length - 2);
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
            worksheet.Row(end1 + 2).InsertRowsBelow(excelData.OneToManyData["试验项目"].Length);
            var end2 = end1 + excelData.OneToManyData["试验项目"].Length + 3;
            for (int i = end1 + 3; i < end2; i++)
            {
                int j = i - end1 - 3;
                worksheet.Cell(i, "B").Value = excelData.OneToManyData["试验项目"][j];
                worksheet.Cell(i, "C").Value = excelData.OneToManyData["标准值"][j];
                worksheet.Range(string.Format("C{0}", i), string.Format("D{0}", i)).Merge();
            }

            worksheet.Range("A9", string.Format("A{0}", end1 - 1)).Merge();


            return qualityList;
        }


        private XLWorkbook FillSelfCheckTable(string templateType)
        {
            var selfCheckTable = Utils.GetTemplateExcel(templateType, "各厂家自查表模版.xlsx");
            var worksheet = selfCheckTable.Worksheet(1);
            string usedMaterial = "";
            int materialLength = 0;
            for (int i = 0; i < excelData.OneToManyData["所用材料"].Length; i++)
            {
                if (excelData.OneToManyData["所用材料"][i].ToString() != "")
                {
                    materialLength += 1;
                }
            }
            for (int i = 0; i < materialLength; i++)
            {
                usedMaterial += excelData.OneToManyData["所用材料"][i];
                if (i != materialLength - 1)
                {
                    usedMaterial += "、";
                }
                
            }

                for (int i = 4; i < excelData.OneToManyData["材料编码/设备位号"].Length + 4; i++)
            {
                int j = i - 4;
                worksheet.Cell(i, "B").Value = excelData.OneToOneData["站号"]+ "站";
                worksheet.Cell(i, "C").Value = "MP";
                worksheet.Cell(i, "D").Value = excelData.OneToOneData["发货日期"];
                worksheet.Cell(i, "E").Value = excelData.OneToOneData["公司名称"];
                worksheet.Cell(i, "F").Value = excelData.OneToOneData["合同号"];
                worksheet.Cell(i, "G").Value = excelData.OneToOneData["请购单号"];
                worksheet.Cell(i, "H").Value = excelData.OneToManyData["材料编码/设备位号"][j];
                worksheet.Cell(i, "I").Value = excelData.OneToOneData["材料名称"];
                worksheet.Cell(i, "J").Value = excelData.OneToManyData["产品规格(Size)"][j];
                worksheet.Cell(i, "K").Value = usedMaterial;
                worksheet.Cell(i, "L").Value = $"TJMZLBG-yyyymm-{excelData.OneToOneData["质检报告编号"]}";
                worksheet.Cell(i, "M").Value = excelData.OneToOneData["依据标准"];
                worksheet.Cell(i, "N").Value = excelData.OneToManyData["单位"][j];
                worksheet.Cell(i, "O").Value = excelData.OneToManyData["数量"][j];
                worksheet.Cell(i, "T").Value = excelData.OneToOneData["批次"];
                worksheet.Cell(i, "AB").Value = "产品质量证明文件";
            }
            return selfCheckTable;
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
        private XLWorkbook FillReleaseReport(string templateType)
        {
            var releaseReport = Utils.GetTemplateExcel(templateType, "放行报告模版.xlsx");
            var worksheet = releaseReport.Worksheet(1);
            var tickbox = worksheet.Picture("图片 7");
            int horizontalTickBoxOffset = tickbox.GetOffset(ClosedXML.Excel.Drawings.XLMarkerPosition.TopLeft).X;
            tickbox = tickbox.Duplicate();
            foreach (var i in worksheet.Pictures.ToList())
            {
                if (i == tickbox||i.Name=="图片 0")
                    continue;
                i.Delete();
            }
            worksheet.Cell(3, 'C').Value = excelData.OneToOneData["项目名称"]+excelData.OneToOneData["使用部分"];
            worksheet.Cell(5, 'C').Value = excelData.OneToOneData["业主"];
            worksheet.Cell(7, 'C').Value = excelData.OneToOneData["材料名称"];
            worksheet.Cell(9, 'C').Value = excelData.OneToOneData["公司名称"];
            worksheet.Cell(11, 'C').Value = excelData.OneToOneData["供方地点"];
            worksheet.Cell(3, 'G').Value = excelData.OneToOneData["使用部分"];
            worksheet.Cell(5, 'G').Value = excelData.OneToOneData["请购单号"];
            worksheet.Cell(7, 'G').Value = excelData.OneToOneData["合同号"];
            worksheet.Cell(9, 'G').Value = excelData.OneToOneData["使用部分"];
            worksheet.Cell(10, 'G').Value = excelData.OneToOneData["放行联系人"];
            worksheet.Cell(11, 'G').Value = excelData.OneToOneData["放行联系人电话"];
            var height = worksheet.Row(16).Height;
            var heights = worksheet.Rows(17, 33).Select(x => x.Height);
            worksheet.Row(16).Delete();
            worksheet.Row(15).InsertRowsBelow(excelData.OneToManyData["材料编码/设备位号"].Length);
            var end = 15 + excelData.OneToManyData["材料编码/设备位号"].Length+1;
            for (int i = 16; i < end; i++)
            {
                worksheet.Row(i).Height = height;
                worksheet.Range($"C{i}:D{i}").Merge();
                int j = i - 16;
                worksheet.Cell(i, "A").Value = i - 14;
                worksheet.Cell(i, "B").Value = excelData.OneToOneData["材料名称"];
                worksheet.Cell(i, "C").Value = excelData.OneToManyData["材料编码/设备位号"][j];
                worksheet.Cell(i, "E").Value = excelData.OneToManyData["材质"][j];
                worksheet.Cell(i, "F").Value = excelData.OneToManyData["数量（Quantity）"][j];
                worksheet.Cell(i, "G").Value = excelData.OneToManyData["重量"][j];
                worksheet.Cell(i, "H").Value = excelData.OneToManyData["备注-或物明细"][j];
            }
            var temp = new int[]{21,22,25,26,27 };
            foreach(var i in temp)
            {
                tickbox = tickbox.MoveTo(worksheet.Cell(i - 17 + end, "A"), horizontalTickBoxOffset, 0);
                tickbox = tickbox.Duplicate();
            }
            tickbox.Delete();
            return releaseReport;
        }

        private XWPFDocument FillCoverPage(string templateType)
        {
            var document = Utils.GetTemplateDocument(templateType, "1封面模版.docx");
            MessageBox.Show(document.Tables.Count.ToString());
            foreach(var i in Utils.RecursiveParagraphsIterator(document))
            {
                i.ReplaceText("{业主}", excelData.OneToOneData["业主"]);
                i.ReplaceText("{项目名称}", excelData.OneToOneData["项目名称"]);
                i.ReplaceText("{站号}", excelData.OneToOneData["站号"]+"站");
                i.ReplaceText("{使用部分}", excelData.OneToOneData["使用部分"]);
                i.ReplaceText("{材料名称}", excelData.OneToOneData["材料名称"]);
                i.ReplaceText("{合同号}", excelData.OneToOneData["合同号"]);
                i.ReplaceText("{请购单号}", excelData.OneToOneData["请购单号"]);
                i.ReplaceText("{版次}", excelData.OneToOneData["版次"]);
                i.ReplaceText("{批次}", excelData.OneToOneData["批次"]);
                i.ReplaceText("{供货商名称}", excelData.OneToOneData["公司名称"]);
                i.ReplaceText("{地址}", excelData.OneToOneData["地址"]);
                i.ReplaceText("{电话}", excelData.OneToOneData["电话"]);
                i.ReplaceText("{传真}", excelData.OneToOneData["传真"]);
                i.ReplaceText("{联系人}", excelData.OneToOneData["联系人"]);
                i.ReplaceText("{公司名称}", excelData.OneToOneData["公司名称"]);
                i.ReplaceText("{合同编号}", excelData.OneToOneData["合同号"]);
                i.ReplaceText("{材料名称}", excelData.OneToOneData["材料名称"]);
            }
            return document;
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
                OnPropertyChanged(nameof(IsOutputsNotNull));
            }
        }
        private Dictionary<string, XWPFDocument>? outputDocxs;
        public Dictionary<string, XWPFDocument>? OutputDocxs
        {
            get => outputDocxs; set
            {
                outputDocxs = value;
                OnPropertyChanged(nameof(OutputDocxs));
                OnPropertyChanged(nameof(IsOutputsNotNull));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public bool IsExcelDataNotNull { get => ExcelData != null; }

        public bool IsOutputsNotNull { get => OutputExcels != null && OutputDocxs!=null; }

        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
