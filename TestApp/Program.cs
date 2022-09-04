using DFrameLib;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Tools.Excel;
using System.Globalization;

internal class Program
{
    private static void Main(string[] args)
    {
        Excel.Application exApp = new Excel.Application();
        if (exApp == null)
        {
            Console.WriteLine("Excel is not installed!");
            return;
        }

        string path = Path.GetDirectoryName(typeof(Program).Assembly.Location);
        //Console.WriteLine(path);
        Excel.Workbook wb = exApp.Workbooks.Open(path + "\\Book1.xlsx");
        Excel.Worksheet ws = wb.Sheets[1];

        /*
        var usedRange = ws.
        var lastRow = usedRange.Rows.Count;
        var lastCol = usedRange.Columns.Count;
        */
        /*
        try
        {
            Console.WriteLine(s);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        */
        for( int i =1; i<4; i++)
        {
            for (int j = 1; j < 4; j++)
            {
                string s = ws.Cells[i, j].Value.ToString();
                //string s = range.Text;
                Console.Write(s+"\t");
            }
            Console.WriteLine();
        }
        

        /*
        var df = new DFrame();
        var col = new DataColumn("Id", Type.GetType("System.Int32"));
        */



        Console.ReadLine();
    }
}