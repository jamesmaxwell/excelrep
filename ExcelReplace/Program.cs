using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;

namespace ExcelReplace
{
    class Program
    {
        const int ColLimit = 5;

        static void Main(string[] args)
        {
            var srcFile = "./files/公安干警名单2.xls";
            var newFile = "./files/公安OK2.xls";

            var srcBook = new Workbook(srcFile);
            var newBook = new Workbook(newFile);

            for (var i = 0; i < newBook.Worksheets.Count; i++)
            {
                var srcSheet = srcBook.Worksheets[i];
                var newSheet = newBook.Worksheets[i];
                for (var row = 1; row < newSheet.Cells.Rows.Count; row++)
                {
                    var nRow = newSheet.Cells.Rows[row];
                    if (nRow.IsBlank)
                        continue;
                    var id = nRow[0].Value.ToString();

                    var findRow = -1;
                    for (var sRow = 1; sRow < srcSheet.Cells.Rows.Count; sRow++)
                    {
                        var sId = srcSheet.Cells.Rows[sRow][0].Value.ToString();
                        if (id == sId)
                        {
                            findRow = sRow;
                            break;
                        }
                    }
                    if (findRow >= 0)
                    {
                        for (var x = 0; x < ColLimit; x++)
                        {
                            srcSheet.Cells.Rows[findRow][x].Value = newSheet.Cells.Rows[row][x].Value;
                        }
                    }
                }
                Console.WriteLine("sheet{0} finished", i);
                //Console.WriteLine("srcSheet{0}:{1} rows", i, srcSheet.Cells.Rows.Count);
                // Console.WriteLine("newSheet{0}:{1} rows", i, newSheet.Cells.Rows.Count);
            }
            srcBook.Save("公安.xls", SaveFormat.Excel97To2003);

            Console.WriteLine("save file ok");

            Console.ReadKey();
        }
    }
}
