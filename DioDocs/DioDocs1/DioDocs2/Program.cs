using System;
using GrapeCity.Documents.Excel;

namespace DioDocs2
{
    class Program
    {
        static void Main(string[] args)
        {
            Workbook workbook = new Workbook();

            workbook.Open(@"Book1.xlsx");

            IWorksheet worksheet = workbook.Worksheets[0];

            // numeric 5.2
            int rowIndex = 0;
            for (decimal dec = 999.99M; dec > -1000M; dec -= 0.01M)
            {
                float f = (float)dec;

                worksheet.Cells[rowIndex, 0].Value = f;
                worksheet.Cells[rowIndex, 4].Value = f.ToString();

                worksheet.Cells[rowIndex, 1].Value = (decimal)f;

                worksheet.Cells[rowIndex, 2].Value = dec;

                rowIndex++;
            }

            workbook.Save(@"Book2.xlsx");
        }
    }
}
