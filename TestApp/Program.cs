using DFrameLib;
//using Microsoft.Office.Tools.Excel;
using System.Data;

internal class Program
{
    static DFrame GetDataFromExcel(string fname)
    {
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

        var df1 = new DFrame();

        // Console.WriteLine("Создаем колонки, присваиваем тип данных.");
        //List<string> colnames = new List<string>();
        for (int i = 1; i <= maxcol; i++) 
            df1.AddCol(ws.Cells[1, i].Value.ToString(), 
                ws.Cells[2, i].Value.GetType());
        Console.WriteLine("Читаем данные в объект DFrame.");

        // читаем данные из Excel
        object[] row = new object[maxcol];
        for (int j = 2; j <= maxrow; j++)
        {
            for (int i = 1; i <= maxcol; i++)
            {
                row[i - 1] = ws.Cells[j, i].Value;
            }
            //df.Rows.Add(row);
            df1.AddRow(row);
        }
        return df1;

    }
    
    private static void Main(string[] args)
    {
        
        var df = new DFrame();
        

        // заполнение таблицы программным способом
        {
            df.Columns.Add("Id", typeof(int));
            df.Columns.Add("Name", typeof(string));
            df.Columns.Add("DateBirth", typeof(DateTime));
            df.Columns.Add("Pet", typeof(string));

            df.AddRow(new object[] { 1, "Ann", DateTime.Parse("01.01.2002"), "dog" });
            df.AddRow(new object[] { 7, "Mary", DateTime.Parse("25.12.1997"), "cat" });
            df.AddRow(new object[] { 10, "John", DateTime.Parse("14.07.2005"), "dog" });
            df.AddRow(new object[] { 11, "Alex", DateTime.Parse("08.03.1995"), "" });
            df.AddRow(new object[] { 14, "Mary", DateTime.Parse("11.11.1990"), "" });
            df.AddRow(new object[] { 9, "Ann", DateTime.Parse("03.02.1993"), "cat" });
        }
        Console.WriteLine("Вывод таблицы:");
        df.PrintTable();

        //Console.WriteLine("Вывод представления со столбцами 'Name', 'Pet'");
        Console.Write("Введите имена столбцов для отображения через пробел: ");
        string e = Console.ReadLine();
        string[] t = e.Split(" ");
        

        df.SelectColByName(t);
        //df.PrintView(v);

        Console.WriteLine("Исходная таблица не повреждена:");
        df.PrintTable();

        /*
        Console.WriteLine("Вывод всех столбцов без заголовков:");
        var items = df.Select();
        foreach (var b in items) Console.WriteLine("{0}\t{1}\t{2}", 
            b["id"], b["Age"], b["Name"]);
        */

        //DataView dview = new DataView(df);
        //df.PrintTable();
        /*
        Console.WriteLine("Выбор строк, где DateBirth >= 01.01.1999:");
        string expr = "DateBirth >= 01/01/1999";
        DataRow[] foundRows = df.Select(expr);
        for (int i = 0; i < foundRows.Length; i++) 
            Console.WriteLine($"{foundRows[i][0]}\t{foundRows[i][1]}");
        */

        // Выбор строк без метода
        /*
        Console.WriteLine("Выбор строк, где Name = 'Ann':");
        var expr = "Name = 'Ann'";
        DataRow[] foundRows2 = df.Select(expr);
        for (int i = 0; i < foundRows2.Length; i++)
            Console.WriteLine(foundRows2[i][0] + "\t" +
                foundRows2[i][1] + "\t" + foundRows2[i][2]);
        */
        Console.WriteLine();
        


        Console.WriteLine("end.");
        //Console.ReadLine();
        // close Excel process
        /*
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
        */
    }
}