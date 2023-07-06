using ClosedXML.Excel;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.PixelFormats;
using System.IO;

public class CertificateSealPrint
{
    public static void Main()
    {
        action("C:\\Users\\hyang\\Desktop\\产品合格证.xlsx",
            "C:\\Users\\hyang\\Desktop\\seal.png",
            "C:\\Users\\hyang\\Desktop\\aaatest.xlsx");
    }

    public static void action(string workbookPath, string imagePath, string savedPath)
    {
        XLWorkbook workbook = new XLWorkbook(workbookPath);
        IXLWorksheet worksheet = workbook.Worksheet(1);

        Random random = new Random();
        int sealCount = 0;
        int lastRowUsed = worksheet.LastRowUsed().RowNumber();
        IXLCell lastCell = worksheet.Row(lastRowUsed).LastCellUsed();
        int lastColumUsed = lastCell.Address.ColumnNumber;

        if (lastColumUsed == 5)
        {
            sealCount = ((lastRowUsed + 2) / 28) * 2 - 1;
        }
        else
        {
            sealCount = ((lastRowUsed + 2) / 28) * 2;
        }
        Console.WriteLine(lastRowUsed);
        Console.WriteLine(lastColumUsed);
        Console.WriteLine(sealCount);
        int flag1 = 0;
        int flag2 = 0;
        for (int i = 0; i < sealCount; i++)
        {

            using (Image<Rgba32> image = Image.Load<Rgba32>(imagePath))
            {

                float rotationAngle = (float)(random.NextDouble() * 70 - 35);

                image.Mutate(x => x.Rotate(rotationAngle));


                using (MemoryStream ms = new MemoryStream())
                {
                    image.Save(ms, new SixLabors.ImageSharp.Formats.Png.PngEncoder());

                    if (i % 2 == 0)
                    {
                        var picture = worksheet.AddPicture(ms)
                            .MoveTo(worksheet.Cell($"L{15 + flag1 * 28}"))
                            .WithSize(280, 280);
                        flag1 += 1;
                    }
                    else
                    {
                        var picture = worksheet.AddPicture(ms)
                            .MoveTo(worksheet.Cell($"AG{15 + flag2 * 28}"))
                            .WithSize(280, 280);
                        flag2 += 1;
                    }


                }
            }


        }


        workbook.SaveAs(savedPath);
    }



}
