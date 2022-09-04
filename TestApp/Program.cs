using DFrameLib;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Tools.Excel;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using System.Reflection;

internal class Program
{
    private static void Main(string[] args)
    {
        var df = new DFrame();

        Application exApp = new Application();
        System.Diagnostics.Process excelProc = System.Diagnostics.Process.GetProcessesByName("EXCEL").Last();

        if (exApp == null)
        {
            Console.WriteLine("Excel is not installed!");
            return;
        }
        // файл находится в одной папке с программой
        string path = Path.GetDirectoryName(typeof(Program).Assembly.Location);
        // настраиваем переменные для работы с Excel
        
            Excel.Workbook wb = exApp.Workbooks.Open(path + "\\Book1.xlsx");
            Excel.Worksheet ws = wb.Sheets[1];
            int maxrow = ws.UsedRange.Rows.Count;
            int maxcol = ws.UsedRange.Columns.Count;
        

        Console.WriteLine("reading column names.");
        //List<string> colnames = new List<string>();
        for (int i = 1; i <= maxcol; i++)
        {
            df.Columns.Add(new DataColumn(ws.Cells[1, i].Value.ToString(), ws.Cells[2, i].Value.GetType()));            
            //df.Columns[i-1].DataType = ws.Cells[i, 2].Value.GetType();
        }

        object[] row = new object[maxcol];
        for (int j = 2; j <= maxrow; j++)
        {
            for (int i=1; i <= maxcol; i++)
            {
                row[i-1]=(ws.Cells[j, i].Value);
            }
            df.Rows.Add(row);
            row.Initialize();
        }
        /*
        foreach(DataColumn c in df.Columns)
        {
            Console.WriteLine(c.DataType.ToString());
        }
        */
        Console.WriteLine("Вывод всех столбцов без заголовков:");
        var items = df.Select();
        foreach (var b in items) Console.WriteLine("{0}\t{1}\t{2}", b["id"],b["Name"], b["Age"]);
        //Console.WriteLine();

        Console.WriteLine("Выбор строк, где Age >= 24:");
        string expr = "Age >= 24";
        DataRow[] foundRows = df.Select(expr);
        for (int i = 0; i < foundRows.Length; i++)
        {
            Console.WriteLine(foundRows[i][0]+"\t"+ foundRows[i][1]);
        }
        Console.WriteLine("Выбор строк, где Name = 'Ann':");
        expr = "Name = 'Ann'";
        DataRow[] foundRows2 = df.Select(expr);
        for (int i = 0; i < foundRows2.Length; i++)
        {
            Console.WriteLine(foundRows2[i][0] + "\t" + foundRows2[i][1] + "\t" + foundRows2[i][2]);
        }

        //Console.WriteLine(df.Columns.Count);

        //foreach (string s in colnames) Console.Write(s + "\t");



        /*
        for ( int i =1; i<=maxrow; i++)
        {
            for (int j = 1; j <=maxcol; j++)
            {
                string s = ws.Cells[i, j].Value.ToString();
                //string s = range.Text;
                Console.Write(s+"\t");
            }
            Console.WriteLine();
        }
        
        var t = ws.Cells[1, 1].Value.GetType();
        Console.WriteLine("Cell[1,1] has '"+t+"' format.");

        t = ws.Cells[2, 2].Value.GetType();
        Console.WriteLine("Cell[2,2] has '" + t + "' format.");

        t = ws.Cells[3, 1].Value.GetType();
        Console.WriteLine("Cell[3,1] has '" + t + "' format.");
        /*
        var df = new DFrame();
        var col = new DataColumn("Id", Type.GetType("System.Int32"));
        */



        Console.WriteLine("end.");
        Console.ReadLine();
        // close Excel process
        {
            ws = null;
            wb.Close(false, Type.Missing, Type.Missing);
            wb = null;
            exApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(exApp);
            GC.Collect();
            exApp = null;
            System.GC.Collect();
            excelProc.Kill();
        }
    }
}