using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Konsol
{
    class Program
    {
        static void Main(string[] args)
        {

            IXLWorkbook workbook = new XLWorkbook(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Excel Örnek\\Verilerim.xlsx");
            IXLWorksheet worksheet = workbook.Worksheet("Veriler");

            IXLRange range = worksheet.RangeUsed();

            IXLRangeRow rangeRow = range.FirstRow();
            IXLRow row = worksheet.Row(5);

            IXLColumn column = worksheet.Column("C");

            double width = column.ColumnRight(5).Width;
            Console.WriteLine("Sütun Genişliği:" + width);

            Console.WriteLine("Value : " + row.FirstCell().Value);

            Console.WriteLine("C5'in Değeri: " + worksheet.Cell("C5").Value);

            IXLAddress address = worksheet.Cell("C5").Address;
            IXLCell cell = worksheet.Cell("C5");

            for (int i = 1; i <= range.RowCount(); i++)
            {
                if (range.Row(i).Cell(2).Value.ToString() == "Yılmaz")
                {
                    IXLAddress address1 = range.Row(i).Cell(2).Address;
                    Console.WriteLine(address1.RowNumber + " / " + address1.ColumnLetter);
                    break;
                }
            }


            Console.Read();
        }
    }
}
