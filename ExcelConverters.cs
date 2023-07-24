using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using NPOI.XWPF.UserModel;
using DocumentFormat.OpenXml.Wordprocessing;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp;
using NPOI.Util;
using DocumentFormat.OpenXml.Spreadsheet;

namespace GenTemplateBJ
{
    internal class ExcelConverters: INotifyPropertyChanged
    {
        public List<string> TemplateTypes { get; } = new() {"请选择", "川西" };
        public Dictionary<string, Action> TemplateTypeToExcelConverter { get; } = new();

        public Image<Rgba32> Seal { get; private set; }
        public Image<Rgba32> Logo { get; private set; }
        public ExcelConverters()
        {
            Seal = Image.Load<Rgba32>(Utils.GetResourcePath("qualityseal.png"));
            Logo = Image.Load<Rgba32>(Utils.GetResourcePath("logo.jpg"));
            TemplateTypeToExcelConverter["川西"] = () =>
            {
                OutputExcels = new()
                {
                    { "发货清单.xlsx", FillTransportList("川西") },
                    { "质检报告.xlsx", FillQualityList("川西") },
                    { "各厂家自查表.xlsx", FillSelfCheckTable("川西") },
                    {"产品合格证.xlsx",FillProductionCertificate("川西") },
                    {"放行报告.xlsx", FillReleaseReport("川西") },
                    {"装箱单.xlsx", FillPackingList("川西") },
                    {"装箱单2.xlsx", FillPackingList2("川西") }
                };
                OutputDocxs = new()
                {
                    {"封面.docx", FillCoverPage("川西") }
                };
                InitializeExcelsPrintSetting();
                AddCertificateSealToExcels();
            };
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        public static XLWorkbook ConvertFactoryData(XLWorkbook dataExample, XLWorkbook mapping, XLWorkbook originalTest)
        {
            var worksheet1 = dataExample.Worksheet(1);

            var worksheet2 = mapping.Worksheet(1);

            var worksheet3 = originalTest.Worksheet(1);

            Dictionary<string, string> deviceCodeMapping = new();

            foreach (var row in worksheet2.RowsUsed().Skip(3))
            {
                var deviceCode1 = row.Cell("F").GetValue<string>();
                var deviceCode2 = row.Cell("V").CachedValue.ToString();
                deviceCodeMapping[deviceCode1] = deviceCode2;
                //Console.WriteLine(deviceCode1 + " " + deviceCode2);
            }

            int i = 2;
            foreach (var row in worksheet1.RowsUsed().Skip(1))
            {

                var deviceDesignator1 = row.Cell(1).GetValue<string>();
                var caseNumber = row.Cell(2).GetValue<string>();
                Console.WriteLine(1);


                if (deviceCodeMapping.ContainsKey(deviceDesignator1))
                {
                    var deviceCode2 = deviceCodeMapping[deviceDesignator1];
                    Console.WriteLine($"Code2: {deviceCode2}, Case Number: {caseNumber}");
                    worksheet3.Cell(i, "C").Value = deviceCode2;
                    worksheet3.Cell(i, "M").Value = caseNumber;
                    i++;
                }
            }
            return originalTest;
        }
        private void AddCertificateSealToExcels()
        {

            void AddSeal(IXLWorksheet worksheet, int lastDataRow, int headerRowEnd, int horizontalOffsetFromRight, int verticalOffsetFromBottom)
            {
                var first = 1;
                var halfwidth = worksheet.ColumnsUsed().Select(x=>x.Width).Sum()/2;
                worksheet.PageSetup.RowBreaks.Sort();
                if(worksheet.PageSetup.RowBreaks.Count > 0) 
                {
                    Utils.AddPictureToExcel(worksheet, Seal.Clone(), worksheet.Cell(worksheet.PageSetup.RowBreaks[0]-verticalOffsetFromBottom,
                        worksheet.LastColumnUsed().ColumnNumber() - horizontalOffsetFromRight), 280);
                    first = worksheet.PageSetup.RowBreaks[0];
                    foreach (var i in worksheet.PageSetup.RowBreaks.Skip(1))
                    {
                        Utils.AddPictureToExcel(worksheet, Seal.Clone(), worksheet.Cell(i-verticalOffsetFromBottom,
                            worksheet.LastColumnUsed().ColumnNumber() - horizontalOffsetFromRight), 280);
                        first = i;
                    }
                }
                else
                {
                    Utils.AddPictureToExcel(worksheet, Seal.Clone(), worksheet.Cell(worksheet.LastRowUsed().RowNumber() - verticalOffsetFromBottom, 
                        worksheet.LastColumnUsed().ColumnNumber() - horizontalOffsetFromRight), 280);
                    if (headerRowEnd + excelData.OneToManyData["材料编码/设备位号"].Length > lastDataRow)
                    {
                        Utils.AddPictureToExcel(worksheet, Seal.Clone(), worksheet.Cell(worksheet.LastRowUsed().RowNumber(),
                            worksheet.LastColumnUsed().ColumnNumber() - horizontalOffsetFromRight), 280);
                    }
                }

            }
            AddSeal(OutputExcels["质检报告.xlsx"].ActiveWorkSheets.Single(),32, 8, 3, 8);
            AddSeal(OutputExcels["发货清单.xlsx"].ActiveWorkSheets.Single(), 28, 9, 4, 5);
            AddSeal(OutputExcels["放行报告.xlsx"].ActiveWorkSheets.Single(), 24, 15, 2, 5);
        }

        private void InitializeExcelsPrintSetting()
        {
            GeneratePreprintExcels(OutputExcels["质检报告.xlsx"], 8, 8, 0.88, (worksheet, x, y) => 
            {
                var startRow = worksheet.Row(x.begin);
                var endRow = worksheet.Row(y.begin);
                var end = startRow.LastCellUsed().Address.ColumnNumber;
                startRow.Row(2, end).CopyTo(endRow.Cell(2));
            });
            GeneratePreprintExcels(OutputExcels["发货清单.xlsx"], 9, 9, 0.81, (worksheet, x, y) =>
            {
                var startRow = worksheet.Row(x.begin);
                var endRow = worksheet.Row(y.begin);
                var end = startRow.LastCellUsed().Address.ColumnNumber;
                startRow.Row(1, end).CopyTo(endRow.Cell(1));
            });
            GeneratePreprintExcels(OutputExcels["放行报告.xlsx"], 14, 15, 0.77, (worksheet, x, y) =>
            {
                var startRow = worksheet.Row(x.begin);
                var endRow = worksheet.Row(y.begin);
                var end = startRow.LastCellUsed().Address.ColumnNumber;
                startRow.CopyTo(endRow.Cell(1));
                worksheet.Row(x.end).CopyTo(worksheet.Cell(y.end, 1));
                worksheet.Range($"C{y.begin}:D{y.end}").Merge();
                for (int i = 1; i < end+1; i++)
                {
                    if(i ==3||i== 4)
                    {
                        continue;
                    }
                    worksheet.Range(worksheet.Cell(y.begin, i), worksheet.Cell(y.end, i)).Merge();
                }
            });
            OutputExcels["放行报告.xlsx"].ActiveWorkSheets.Single().Range("BE:BP").Delete(XLShiftDeletedCells.ShiftCellsLeft); // 由于某些未知原因，去掉之后结果会很诡异。
        }

        private void GeneratePreprintExcels(ExcelWrapper wrapper, int headerRowStart, int headerRowEnd, double percentage,
            Action<IXLWorksheet,(int begin, int end), (int begin, int end)> applyHeader)
        {
            var worksheet = wrapper.ActiveWorkSheets.Single();
            var margins = OutputDocxs["封面.docx"].Document.body.sectPr.pgMar;
            (double w, double h) = (11906, 16838);
            var contentHeight = worksheet.PageSetup.PageOrientation switch  
                {
                    XLPageOrientation.Portrait => (h - margins.top - margins.bottom) / 20,
                    XLPageOrientation.Landscape=> (w - margins.top - margins.bottom) / 20,
                    _ => (h - margins.top - margins.bottom) / 20
                };
            contentHeight /= percentage;
            var temp = contentHeight;
            int count = headerRowEnd+excelData.OneToManyData["材料编码/设备位号"].Length;
            var new_rows_count = headerRowEnd - headerRowStart + 1;
            for (int j = 1; j < count + 1; j++)
            {
                temp -= worksheet.Row(j).Height;
                if (temp - worksheet.Row(j + 1).Height < 0&&count!=j)
                {
                    worksheet.Row(j).InsertRowsBelow(new_rows_count);
                    worksheet.PageSetup.AddHorizontalPageBreak(j);
                    applyHeader(worksheet, (headerRowStart, headerRowEnd), (j+1, j+new_rows_count));
                    for (int i = j+1; i < j+1+new_rows_count; i++)
                    {
                        worksheet.Row(i).Height = worksheet.Row(headerRowStart + i - j - 1).Height;
                    }
                    count += new_rows_count;
                    temp = contentHeight;
                }
            }
        }


        private ExcelWrapper FillTransportList(string templateType)
        {
            var transportList=Utils.GetTemplateExcel(templateType, "2送货单模版.xlsx");
            var worksheet = transportList.Worksheet(1);
            var result = new ExcelWrapper(transportList, transportList.Worksheet(1));
            worksheet.Cell(1, "C").Value = excelData.OneToOneData["项目名称"];
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
                worksheet.Cell(i, "C").Value = excelData.OneToOneData["产品名称"];
                worksheet.Cell(i, "D").Value = excelData.OneToManyData["产品规格(Size)"][j];
                worksheet.Cell(i, "E").Value = excelData.OneToManyData["材质＆标准（描述）"][j];
                worksheet.Cell(i, "F").Value = excelData.OneToManyData["单位（Unit）"][j];
                worksheet.Cell(i, "G").Value = excelData.OneToManyData["数量（Quantity）"][j];
                worksheet.Cell(i, "H").Value = excelData.OneToManyData["箱号"][j];
                worksheet.Cell(i, "I").Value = excelData.OneToManyData["备注（跟踪号）"][j];
            }
            worksheet.Cell(end, "G").Value = excelData.OneToManyData["数量（Quantity）"].Select(x => (int)x.GetUnifiedNumber()).Sum();
            return result;
        }


// 
        private ExcelWrapper FillPackingList(string templateType)
        {
            var packingList = Utils.GetTemplateExcel(templateType, "8装箱单模版.xlsx");
            var worksheet = packingList.Worksheet(1);
            var result = new ExcelWrapper(packingList, packingList.Worksheet(1));
            Utils.AddPictureToExcel(worksheet, Logo.Clone(), worksheet.Cell("A1"), 230, 115);
            worksheet.Cell(1, "I").Value = excelData.OneToOneData["项目名称"] + "  " + excelData.OneToOneData["使用部分"];
            worksheet.Cell(3, "D").Value = excelData.OneToOneData["材料名称"];
            worksheet.Cell(1, "AG").Value = $"装箱单号: {excelData.OneToOneData["请购单号"]}-{excelData.OneToOneData["批次"]}";
            worksheet.Cell(4, "D").Value = excelData.OneToOneData["合同号"];
            worksheet.Cell(5, "D").Value = excelData.OneToOneData["请购单号"];
            worksheet.Cell(6, "D").Value = excelData.OneToOneData["发货日期"];
            worksheet.Cell(7, "D").Value = excelData.OneToOneData["预计到达日期"];
            worksheet.Cell(3, "W").Value = $"{excelData.OneToOneData["公司名称"]} {excelData.OneToOneData["发货人 电话"]}";
            worksheet.Cell(4, "W").Value = excelData.OneToOneData["收货人 电话"];
            worksheet.Cell(5, "W").Value = excelData.OneToOneData["承运商"];
            worksheet.Cell(6, "W").Value = excelData.OneToOneData["运输方式"];
            worksheet.Cell(7, "W").Value = excelData.OneToOneData["到货地点"];


            List<string> packNumList = new List<string>();
            for (int i = 0; i < excelData.OneToManyData["材料编码/设备位号"].Length; i++)
            {
                packNumList.Add($"{excelData.OneToManyData["箱号"][i]}");
            }
            IEnumerable<string> distinctValues = packNumList.Distinct();
            int packNum = distinctValues.Count();
            for (int i = 1; i < packNum; i++)
            {
                worksheet.Range("A1", "AQ11").CopyTo(worksheet.Cell($"A{1 + 12 * i}"));
                Utils.AdjustHeight(worksheet, 1, 1 + 12 * i, 11);
            }
            List<int> sortedList = packNumList
                .Select(int.Parse)
                .Distinct()
                .OrderBy(n => n)
                .ToList();
            int flag2 = 0;
            for (int i = 0; i < sortedList.Count(); i++)
            {
                int flag = 0;
                worksheet.Cell(2 + 12 * i + flag2, "AG").Value = $"共{packNum}箱   第{i + 1}箱";
                for (int j = 0; j < excelData.OneToManyData["材料编码/设备位号"].Length; j++)
                {
                    if (sortedList[i].ToString() == excelData.OneToManyData["箱号"][j].ToString())
                    {
                        int workline = 10 + 12 * i + flag + flag2;
                        worksheet.Cell(workline, "A").Value = flag + 1;
                        worksheet.Cell(workline, "B").Value = excelData.OneToManyData["材料编码/设备位号"][j];
                        worksheet.Cell(workline, "E").Value = excelData.OneToOneData["材料名称"];
                        worksheet.Cell(workline, "K").Value = excelData.OneToManyData["产品规格(Size)"][j];
                        worksheet.Cell(workline, "W").Value = excelData.OneToManyData["单位（Unit）"][j];
                        worksheet.Cell(workline, "AA").Value = excelData.OneToManyData["数量（Quantity）"][j];


                        worksheet.Cell(workline, "AP").Value = excelData.OneToManyData["箱号"][j];
                        worksheet.Row(workline).InsertRowsBelow(1);
                        worksheet.Range($"B{workline + 1}:D{workline + 1}").Merge();
                        worksheet.Range($"E{workline + 1}:J{workline + 1}").Merge();
                        worksheet.Range($"K{workline + 1}:R{workline + 1}").Merge();
                        worksheet.Range($"S{workline + 1}:V{workline + 1}").Merge();
                        worksheet.Range($"W{workline + 1}:Z{workline + 1}").Merge();
                        worksheet.Range($"AA{workline + 1}:AC{workline + 1}").Merge();
                        worksheet.Range($"AD{workline + 1}:AH{workline + 1}").Merge();
                        worksheet.Range($"AP{workline + 1}:AQ{workline + 1}").Merge();
                        flag++; 
                    }


                }
                worksheet.Row(10 + 12 * i + flag + flag2).Delete();
                flag2 += flag - 1;
            }
            return result;
        }



        private ExcelWrapper FillPackingList2(string templateType)
        {
            var packingList2 = Utils.GetTemplateExcel(templateType, "8装箱单模版.xlsx");
            var worksheet1 = packingList2.Worksheet(1);
            
            List<string> packNumList = new List<string>();
            for (int i = 0; i < excelData.OneToManyData["材料编码/设备位号"].Length; i++)
            {
                packNumList.Add($"{excelData.OneToManyData["箱号"][i]}");
            }
            IEnumerable<string> distinctValues = packNumList.Distinct();
            List<int> sortedList = packNumList
                .Select(int.Parse)
                .Distinct()
                .OrderBy(n => n)
                .ToList();
            var packNum = sortedList.Count;
            for (int i = 0; i < sortedList.Count; i++)
            {
                var worksheet = worksheet1.CopyTo($"Sheet{i}");
                Utils.AddPictureToExcel(worksheet, Logo.Clone(), worksheet.Cell("A1"), 230, 115);
                worksheet.Cell(1, "I").Value = excelData.OneToOneData["项目名称"] + "  " + excelData.OneToOneData["使用部分"];
                worksheet.Cell(3, "D").Value = excelData.OneToOneData["材料名称"];
                worksheet.Cell(1, "AG").Value = $"装箱单号: {excelData.OneToOneData["请购单号"]}-{excelData.OneToOneData["批次"]}";
                worksheet.Cell(4, "D").Value = excelData.OneToOneData["合同号"];
                worksheet.Cell(5, "D").Value = excelData.OneToOneData["请购单号"];
                worksheet.Cell(6, "D").Value = excelData.OneToOneData["发货日期"];
                worksheet.Cell(7, "D").Value = excelData.OneToOneData["预计到达日期"];
                worksheet.Cell(3, "W").Value = $"{excelData.OneToOneData["公司名称"]} {excelData.OneToOneData["发货人 电话"]}";
                worksheet.Cell(4, "W").Value = excelData.OneToOneData["收货人 电话"];
                worksheet.Cell(5, "W").Value = excelData.OneToOneData["承运商"];
                worksheet.Cell(6, "W").Value = excelData.OneToOneData["运输方式"];
                worksheet.Cell(7, "W").Value = excelData.OneToOneData["到货地点"];
                worksheet.Cell(2, "AG").Value = $"共{packNum}箱   第{i + 1}箱";
                for (int j = 0; j < excelData.OneToManyData["材料编码/设备位号"].Length; j++)
                {
                    var workline = 10 + j;
                    if (sortedList[i].ToString() == excelData.OneToManyData["箱号"][j].ToString())
                    {
                        worksheet.Cell(workline, "A").Value = j + 1;
                        worksheet.Cell(workline, "B").Value = excelData.OneToManyData["材料编码/设备位号"][j];
                        worksheet.Cell(workline, "E").Value = excelData.OneToOneData["材料名称"];
                        worksheet.Cell(workline, "K").Value = excelData.OneToManyData["产品规格(Size)"][j];
                        worksheet.Cell(workline, "W").Value = excelData.OneToManyData["单位（Unit）"][j];
                        worksheet.Cell(workline, "AA").Value = excelData.OneToManyData["数量（Quantity）"][j];
                        worksheet.Cell(workline, "AP").Value = excelData.OneToManyData["箱号"][j];
                        worksheet.Row(workline).InsertRowsBelow(1);
                        worksheet.Range($"B{workline + 1}:D{workline + 1}").Merge();
                        worksheet.Range($"E{workline + 1}:J{workline + 1}").Merge();
                        worksheet.Range($"K{workline + 1}:R{workline + 1}").Merge();
                        worksheet.Range($"S{workline + 1}:V{workline + 1}").Merge();
                        worksheet.Range($"W{workline + 1}:Z{workline + 1}").Merge();
                        worksheet.Range($"AA{workline + 1}:AC{workline + 1}").Merge();
                        worksheet.Range($"AD{workline + 1}:AH{workline + 1}").Merge();
                        worksheet.Range($"AP{workline + 1}:AQ{workline + 1}").Merge();

                    }
                }
                worksheet.Row(10 + excelData.OneToManyData["材料编码/设备位号"].Length).Delete();

            }
            
            var worksheetToDelete = packingList2.Worksheet(1);
            worksheetToDelete.Delete();
            var worksheets = new IXLWorksheet[packingList2.Worksheets.Count];
            {
                var i = 0;
                foreach (IXLWorksheet worksheet in packingList2.Worksheets)
                {
                    worksheets[i] = worksheet;
                    i++;
                }
            }
            var result = new ExcelWrapper(packingList2, worksheets);
            return result;  

        }

        private ExcelWrapper FillQualityList(string templateType)
        {
            var qualityList = Utils.GetTemplateExcel(templateType, "4质检报告模版.xlsx");
            var worksheet = qualityList.Worksheet("检验报告-02804-01-4000-MP-R-M-8050");
            var result = new ExcelWrapper(qualityList, worksheet);
            worksheet.Cell(3, "A").Value = $"报告编号: TJMZLBG-{DateTime.Now.Year}{DateTime.Now.Month.ToString("D2")}-{excelData.OneToOneData["质检报告编号"]}";
            worksheet.Cell(4, "B").Value = excelData.OneToOneData["公司名称"];
            worksheet.Cell(5, "B").Value = excelData.OneToOneData["项目名称"];
            worksheet.Cell(6, "B").Value = excelData.OneToOneData["依据标准"];
            worksheet.Cell(7, "B").Value = excelData.OneToOneData["使用部分"];

            worksheet.Row(10).InsertRowsBelow(excelData.OneToManyData["材料编码/设备位号"].Length - 2);
            worksheet.Cell(9, "A").Value = excelData.OneToOneData["材料名称"];
            var end1 = 9 + excelData.OneToManyData["材料编码/设备位号"].Length;
            for (int i = 9; i < end1; i++)
            {
                int j = i - 9;
                worksheet.Row(i).Height = worksheet.Row(9).Height;
                worksheet.Cell(i, "B").Value = excelData.OneToManyData["材料编码/设备位号"][j];
                worksheet.Cell(i, "C").Value = excelData.OneToManyData["产品规格(Size)"][j];
                worksheet.Cell(i, "D").Value = excelData.OneToManyData["数量（Quantity）"][j];
                worksheet.Cell(i, "E").Value = excelData.OneToManyData["单位（Unit）"][j];
                worksheet.Cell(i, "F").Value = excelData.OneToManyData["生产负责人"][j];
            }
            var end2 = end1 + excelData.OneToManyData["试验项目"].Length + 3;
            int flag = 0;
            for (int i = end1 + 3; i < end2; i++)
            {
                int j = i - end1 - 3;
                if (excelData.OneToManyData["试验项目"][j].ToString() != string.Empty)
                {
                    worksheet.Row(end1 + 2 + flag).InsertRowsBelow(1);
                    worksheet.Cell(i, "A").Value = excelData.OneToManyData["试验项目"][j];
                    worksheet.Cell(i, "C").Value = excelData.OneToManyData["标准值"][j];
                    worksheet.Range(string.Format("A{0}", i), string.Format("B{0}", i)).Merge();
                    worksheet.Range(string.Format("C{0}", i), string.Format("D{0}", i)).Merge();
                    flag ++;
                }
   
            }

            worksheet.Range("A9", string.Format("A{0}", end1 - 1)).Merge();
            return result;
        }


        private ExcelWrapper FillSelfCheckTable(string templateType)
        {
            var selfCheckTable = Utils.GetTemplateExcel(templateType, "各厂家自查表模版.xlsx");
            var worksheet = selfCheckTable.Worksheet(1);
            var result = new ExcelWrapper(selfCheckTable, worksheet);
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
                worksheet.Cell(i, "O").Value = excelData.OneToManyData["数量（Quantity）"][j];
                worksheet.Cell(i, "T").Value = excelData.OneToOneData["批次"];
                worksheet.Cell(i, "AB").Value = "产品质量证明文件";
            }
            return result;
        }

        enum CertificateRowStatus
        {
            Empty,
            LeftFull,
            Full,
        }
        private ExcelWrapper FillProductionCertificate(string templateType)
        {
            var productionCertificate = Utils.GetTemplateExcel(templateType, "3产品合格证模版.xlsx");
            var worksheet = productionCertificate.Worksheet(1);
            var result = new ExcelWrapper(productionCertificate, worksheet);
            int currentLeft = worksheet.ColumnLetterToNumber("A");
            int initialLeft = currentLeft;
            int certificateWidth = worksheet.ColumnLetterToNumber("S") - currentLeft+1;
            int certicateHeight = 26 - 1+1;
            int marginW = 1;
            int marginH = 1;
            var layoutState = CertificateRowStatus.Empty;
            int currentTop = 1;
            int initialTop = currentTop;
            var logo = worksheet.Pictures.Single();

            Utils.AdjustWidth(worksheet, currentLeft, currentLeft+certificateWidth+marginW, certificateWidth);
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
                        Utils.AdjustHeight(worksheet, initialTop, currentTop, certicateHeight+marginH);
                        layoutState = CertificateRowStatus.LeftFull;
                        break;
                }
                int firstCellVerticalOffset = 8 - initialTop;
                int firstCellHorizontalOffset = worksheet.ColumnLetterToNumber("H") - initialLeft;
                logo.Duplicate().MoveTo(worksheet.Cell(currentTop + 1, currentLeft + worksheet.ColumnLetterToNumber("J") - initialLeft));
                worksheet.Cell(currentTop + firstCellVerticalOffset, currentLeft+firstCellHorizontalOffset).Value = excelData.OneToOneData["材料名称"];
                worksheet.Cell(currentTop + firstCellVerticalOffset+2, currentLeft + firstCellHorizontalOffset).Value = productSize;
                worksheet.Cell(currentTop + firstCellVerticalOffset + 2+3, currentLeft + firstCellHorizontalOffset).Value = materialCode;
                worksheet.Cell(currentTop + firstCellVerticalOffset + 2+3+3, currentLeft + firstCellHorizontalOffset).Value = quantity;
                worksheet.Cell(currentTop + firstCellVerticalOffset + 2 + 3 + 3+2, currentLeft + firstCellHorizontalOffset).Value = excelData.OneToOneData["出厂日期"];
                Utils.AddPictureToExcel(worksheet, Seal.Clone(), worksheet.Cell(currentTop + 13, currentLeft + 4), 207);
            }
            for (int i = 0; i < excelData.OneToManyData["材料编码/设备位号"].Length; i++)
            {
                AddOneCertificate(excelData.OneToManyData["产品规格(Size)"][i], excelData.OneToManyData["材料编码/设备位号"][i], excelData.OneToManyData["数量（Quantity）"][i]);
            }
            return result;
        }
        private ExcelWrapper FillReleaseReport(string templateType)
        {
            var releaseReport = Utils.GetTemplateExcel(templateType, "9放行报告模版.xlsx");
            var worksheet = releaseReport.Worksheet(1);
            var result = new ExcelWrapper(releaseReport, worksheet);
            
            var tickbox = worksheet.Picture("图片 7");
            int horizontalTickBoxOffset = tickbox.GetOffset(ClosedXML.Excel.Drawings.XLMarkerPosition.TopLeft).X;
            tickbox = tickbox.Duplicate();
            foreach (var i in worksheet.Pictures.ToList())
            {
                if (i == tickbox||i.Name=="图片 0")
                    continue;
                i.Delete();
            }
            worksheet.Cell(3, 3).Value = excelData.OneToOneData["项目名称"]+excelData.OneToOneData["使用部分"];
            worksheet.Cell(5, 3).Value = excelData.OneToOneData["业主"];
            worksheet.Cell(7, 3).Value = excelData.OneToOneData["材料名称"];
            worksheet.Cell(9, 3).Value = excelData.OneToOneData["公司名称"];
            worksheet.Cell(11, 3).Value = excelData.OneToOneData["供方地点"];
            worksheet.Cell(3, 7).Value = excelData.OneToOneData["使用部分"];
            worksheet.Cell(5, 7).Value = excelData.OneToOneData["请购单号"];
            worksheet.Cell(7, 7).Value = excelData.OneToOneData["合同号"];
            worksheet.Cell(9, 7).Value = excelData.OneToOneData["使用部分"];
            worksheet.Cell(11, 7).Value = excelData.OneToOneData["放行联系人"];
            worksheet.Cell(12, 7).Value = excelData.OneToOneData["放行联系人电话"];
            releaseReport.SaveAs("D:\\PML\\GenTemplateBJ\\debug.xlsx");
            var height = worksheet.Row(16).Height;
            var heights = worksheet.Rows(17, 33).Select(x => x.Height);
            worksheet.Row(16).Delete();
            worksheet.Row(15).InsertRowsBelow(excelData.OneToManyData["材料编码/设备位号"].Length);
            var end = 15 + excelData.OneToManyData["材料编码/设备位号"].Length+1;
            releaseReport.SaveAs("D:\\PML\\GenTemplateBJ\\debug.xlsx");
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
            releaseReport.SaveAs("D:\\PML\\GenTemplateBJ\\debug.xlsx");
            var temp = new int[]{21,22,25,26,27 };
            foreach(var i in temp)
            {
                tickbox = tickbox.MoveTo(worksheet.Cell(i - 17 + end, "A"), horizontalTickBoxOffset, 0);
                tickbox = tickbox.Duplicate();
            }
            tickbox.Delete();
            Utils.AddPictureToExcel(worksheet, Logo.Clone(), worksheet.Cell("A1"), 200, 115);
            releaseReport.SaveAs("D:\\PML\\GenTemplateBJ\\debug.xlsx");
            return result;
        }

        private XWPFDocument FillCoverPage(string templateType)
        {
            var document = Utils.GetTemplateDocument(templateType, "1封面模版.docx");
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

        private Dictionary<string, ExcelWrapper>? outputExcels;
        public Dictionary<string, ExcelWrapper>? OutputExcels
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

        public bool IsExcelDataNotNull { get => ExcelData != null; }

        public bool IsOutputsNotNull { get => OutputExcels != null && OutputDocxs!=null; }

        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
