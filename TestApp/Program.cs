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

        Workbook wb = exApp.Workbooks.Open(path + @"\Book1.xlsx");
        Worksheet ws = wb.Sheets[1]; // нумерация листов начинается с 1
        
        Console.WriteLine("Вывод всех столбцов без заголовков:");
        var items = df.Select();
        foreach (var b in items) Console.WriteLine("{0}\t{1}\t{2}", 
            b["id"], b["Name"], b["Age"]);
        
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