using System;
using System.IO;
using GrapeCity.Documents.Excel;
namespace DioDocs1
{
    class Program
    {
        static void Main(string[] args)
        {
            Workbook workbook = new Workbook();

            workbook.Open(@"Book1.xlsx");

            IWorksheet worksheet = workbook.Worksheets[0];

            float f1 = 2.75F;
            float f2 = 2.76F;
            float f3 = 2.77F;
            float f4 = 2.78F;
            float f5 = 2.79F;
            float f6 = 0.53F;

            worksheet.Cells[0, 0].Value = f1;
            worksheet.Cells[0, 1].Value = f2;
            worksheet.Cells[0, 2].Value = f3;
            worksheet.Cells[0, 3].Value = f4;
            worksheet.Cells[0, 4].Value = f5;
            worksheet.Cells[0, 5].Value = f6;

            Console.WriteLine(f6);

            worksheet.Cells[1, 0].Value = (decimal)f1;
            worksheet.Cells[1, 1].Value = (decimal)f2;
            worksheet.Cells[1, 2].Value = (decimal)f3;
            worksheet.Cells[1, 3].Value = (decimal)f4;
            worksheet.Cells[1, 4].Value = (decimal)f5;
            worksheet.Cells[1, 5].Value = (decimal)f6;

            workbook.Save(@"Book2.xlsx");

        }
    }
}
