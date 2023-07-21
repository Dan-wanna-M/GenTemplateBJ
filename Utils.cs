using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.Win32;
using System.Windows;
using System.IO;
using Microsoft.WindowsAPICodePack.Dialogs;
using NPOI.XWPF;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Dml;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.SS.Formula.Functions;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp;
using static ICSharpCode.SharpZipLib.Zip.ExtendedUnixData;
using SixLabors.ImageSharp.Processing;

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
                try
                {
                    XLWorkbook? workbook = new(fd.OpenFile());
                    return workbook;
                }
                catch(IOException e)
                {
                    MessageBox.Show(e.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            return null;
        }

        public static string? OpenAFolder()
        {
            CommonOpenFileDialog dialog = new()
            {
                InitialDirectory = Environment.CurrentDirectory,
                IsFolderPicker = true
            };
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                return dialog.FileName;
            }
            return null;
        }

        public static IEnumerable<IXLCell> GetCellsUntilLastCellUsed<T>(T data) where T: IXLRangeBase
        {
            return data.Cells($"1:{data.LastCellUsed().Address.RowNumber}");
        }

        private static string GetTemplatePath(string templateType, string name)
        {
            var folderPath = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName;
            return Path.Combine(folderPath, "Templates", templateType, name);
        }

        public static XLWorkbook GetTemplateExcel(string templateType, string name)
        {
            return new XLWorkbook(GetTemplatePath(templateType, name));
        }

        public static int ColumnLetterToNumber(this IXLWorksheet sheet, string letter)
        {
            return sheet.Column(letter).ColumnNumber();
        }

        public static string ColumnNumberToLetter(this IXLWorksheet sheet, int number)
        {
            return sheet.Column(number).ColumnLetter();
        }

        public static XWPFDocument GetTemplateDocument(string templateType, string name)
        {
            using var stream = new FileStream(GetTemplatePath(templateType, name), FileMode.Open);
            return new XWPFDocument(stream);
        }

        public static IEnumerable<XWPFParagraph> RecursiveParagraphsIterator(IBody body)
        {
            foreach (var i in body.Paragraphs)
            {
                yield return i;
            }
            foreach(var i in body.Tables)
            {
                
                foreach (var j in i.Rows)
                {
                    foreach (var k in j.GetTableCells())
                    {
                        foreach (var p in RecursiveParagraphsIterator(k))
                        {
                            yield return p;
                        }
                    }
                }
            }
        }

        public static string GetResourcePath(string filename)
        {
            var folderPath = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.FullName;
            return Path.Combine(folderPath,"resources", filename);
        }

        public static void AddPictureToExcel(IXLWorksheet worksheet, Image<Rgba32> image, IXLCell cell, int sealWidth)
        {
            var random = new Random();
            float rotationAngle = (float)(random.NextDouble() * 20 - 10);
            image.Mutate(x => x.Rotate(rotationAngle));
            using MemoryStream ms = new();
            image.Save(ms, new SixLabors.ImageSharp.Formats.Png.PngEncoder());
            var picture = worksheet.AddPicture(ms);
            picture.MoveTo(cell);
            picture.WithSize(sealWidth, sealWidth * 390/516);
        }
        public static void AddPictureToExcel(IXLWorksheet worksheet, Image<Rgba32> image, IXLCell cell, int pictureWidth, int pictureHeight)
        {
            using MemoryStream ms = new();
            image.Save(ms, new SixLabors.ImageSharp.Formats.Png.PngEncoder());
            var picture = worksheet.AddPicture(ms);
            picture.MoveTo(cell);
            picture.WithSize(pictureWidth, pictureHeight);
        }

        public static void AdjustWidth(IXLWorksheet worksheet, int initialLeft, int current, int size)
        {
            for (int i = current; i < current + size; i++)
            {
                worksheet.Column(i).Width = worksheet.Column(initialLeft + i - current).Width;
            }
        }
        public static void AdjustHeight(IXLWorksheet worksheet, int initialTop, int current, int size)
        {

            for (int i = current; i < current + size; i++)
            {
                worksheet.Row(i).Height = worksheet.Row(initialTop + i - current).Height;
            }
        }
    }
}