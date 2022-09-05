using DFrameLib;
//using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

internal class Program
{
    private static void Main(string[] args)
    {
        var df = new DFrame();
        // подключаемся к Excel
        Application exApp = new Application();
        System.Diagnostics.Process excelProc = 
            System.Diagnostics.Process.GetProcessesByName("EXCEL").Last();
        if (exApp == null)
        {
            Console.WriteLine("Excel is not installed!");
            return;
        }
        // файл находится в одной папке с программой
        string path = Path.GetDirectoryName(typeof(Program).Assembly.Location);
        // настраиваем переменные для работы с Excel
        Workbook wb = exApp.Workbooks.Open(path + @"\Book1.xlsx");
        Worksheet ws = wb.Sheets[1]; // нумерация листов начинается с 1
        int maxrow = ws.UsedRange.Rows.Count;
        int maxcol = ws.UsedRange.Columns.Count;

        Console.WriteLine("Создаем колонки, присваиваем тип данных.");
        //List<string> colnames = new List<string>();
        for (int i = 1; i <= maxcol; i++) 
            df.AddCol(ws.Cells[1, i].Value.ToString(), 
                ws.Cells[2, i].Value.GetType());
        Console.WriteLine("Читаем данные в объект DFrame.");

        object[] row = new object[maxcol];
        for (int j = 2; j <= maxrow; j++)
        {
            for (int i = 1; i <= maxcol; i++)
            {
                row[i - 1] = ws.Cells[j, i].Value;
            }
            //df.Rows.Add(row);
            df.AddRow(row);
        }
        /*
        Console.WriteLine("Вывод всех столбцов без заголовков:");
        var items = df.Select();
        foreach (var b in items) Console.WriteLine("{0}\t{1}\t{2}", 
            b["id"], b["Age"], b["Name"]);
        */
        Console.WriteLine("Вывод таблицы:");
        //DataView dview = new DataView(df);
        df.PrintTable();

        Console.WriteLine("Выбор строк, где Age >= 24:");
        string expr = "Age >= 24";
        DataRow[] foundRows = df.Select(expr);
        for (int i = 0; i < foundRows.Length; i++) 
            Console.WriteLine($"{foundRows[i][0]}\t{foundRows[i][1]}");
        
        Console.WriteLine("Выбор строк, где Name = 'Ann':");
        expr = "Name = 'Ann'";
        DataRow[] foundRows2 = df.Select(expr);
        for (int i = 0; i < foundRows2.Length; i++)
            Console.WriteLine(foundRows2[i][0] + "\t" + 
                foundRows2[i][1] + "\t" + foundRows2[i][2]);


        

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