using DFrameLib;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Linq;
using System;

internal class Program
{
    static private void PrintRows(DataRow[] rows)
    {
        if (rows.Length == 0)
        {
            Console.WriteLine("Записи не найдены.");
            return;
        }
        foreach (var r in rows)
        {
            foreach (DataColumn col in r.Table.Columns)
                if (r[col].GetType() != typeof(DateTime))
                    Console.Write($"{r[col]}\t");
                else
                {
                    DateTime d = (DateTime)r[col];
                    Console.Write($"{d.ToShortDateString()}\t");
                }
            Console.WriteLine();
        }
        Console.WriteLine();
    }

    // печать таблицы (синие заголовки)
    static public void PrintTableOrView(DataTable t)
    {
        if (t.Columns.Count == 0)    // если нет данных, не выводим
        {
            Console.WriteLine("This view has no columns. Nothing to display.");
            return;
        }
        Console.ForegroundColor = ConsoleColor.Blue;
        for (int i = 0; i < t.Columns.Count; i++)
            Console.Write($"{t.Columns[i].ColumnName}\t");
        Console.WriteLine();
        Console.ResetColor();
        DataRow[] rows = t.Select();
        PrintRows(rows);
    }
    // печать отображения (красные заголовки)
    static public void PrintTableOrView(DataView t)
    {
        if (t.Table.Columns.Count == 0)     // если нет данных, не выводим
        {
            Console.WriteLine("This view has no columns. Nothing to display.");
            return;
        }
        Console.ForegroundColor = ConsoleColor.Red;
        for (int i = 0; i < t.Table.Columns.Count; i++)
            Console.Write($"{t.Table.Columns[i].ColumnName}\t");
        Console.WriteLine();
        Console.ResetColor();
        DataRow[] rows = t.Table.Select();
        PrintRows(rows);
    }



    static DFrame GetDataFromExcel(string fname)
    {
        // подключаемся к Excel

        
        Excel.Application exApp = new Excel.Application();
        System.Diagnostics.Process excelProc = 
            System.Diagnostics.Process.GetProcessesByName("EXCEL").Last();
        if (exApp == null)
        {
            Console.WriteLine("Excel is not installed!");
            return null;
        }

        // файл находится в одной папке с программой
        string path = Path.GetDirectoryName(typeof(Program).Assembly.Location);

        // настраиваем переменные для работы с Excel
        Excel.Workbook wb = exApp.Workbooks.Open(path + $@"\{fname}");
        Excel.Worksheet ws = wb.Sheets[1]; // нумерация листов начинается с 1
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
    
    private static void Main(string[] args)
    {
        
        var df = new DFrame();
        // заполнение таблицы программным способом
        {
            df.Columns.Add("Id", typeof(int));
            df.Columns.Add("Name", typeof(string));
            df.Columns.Add("DateBirth", typeof(DateTime));
            df.Columns.Add("Pet", typeof(string));
            df.Columns.Add("PetAge", typeof(int));

            df.AddRow(new object[] { 1, "Ann", DateTime.Parse("01.01.2002"), "dog", 2 });
            df.AddRow(new object[] { 7, "Mary", DateTime.Parse("25.12.1997"), "cat", 4 });
            df.AddRow(new object[] { 10, "John", DateTime.Parse("14.07.2005"), "dog", 3 });
            df.AddRow(new object[] { 11, "Alex", DateTime.Parse("08.03.1995"), "dog", 4 });
            df.AddRow(new object[] { 14, "Mary", DateTime.Parse("11.11.1990"), "", 0 });
            df.AddRow(new object[] { 9, "Ann", DateTime.Parse("03.02.1993"), "cat", 2 });
            df.AddRow(new object[] { 14, "Mary", DateTime.Parse("11.11.1990"), "", 0 });
        }
        // вывод таблицы
        Console.WriteLine("Вывод таблицы:");
        PrintTableOrView(df.AsDataView());

        //Console.WriteLine("Вывод представления со столбцами 'Name', 'Pet'");
        //Console.Write("Введите имена столбцов через пробел для отображения: ");
        string e = "Name Pet"; // Console.ReadLine();
        string[] t = e.Split(" ");
        //df.SelectColByName(t);
        df.SelectColumns(t);

        // Выборка строк
        Console.WriteLine("Строки, где Name = 'Mary':");
        PrintRows(df.SelectRows("Name = 'Mary'"));
        Console.WriteLine("Строки, где Pet = 'Cat' или Pet = '', сортировка по убыванию по полю Id:");
        PrintRows(df.SelectRows("Pet = 'Cat' or Pet = ''", "Id DESC"));

        // переименование столбцов
        var dict = new Dictionary<string, string>()
        {
            {"Id", "id" },
            {"Name", "names" },
            {"Pet", "pets" },
        };
        df.RenameColumns(dict);
        Console.WriteLine("Переименование трех столбцов, вывод:");
        PrintTableOrView(df);

        // LINQ.where
        Console.WriteLine("LINQ.where 1 logic parameter:");
        PrintRows(df.SelectRowsByColname("pets", "cat"));

        Console.WriteLine("LINQ.where 2 logic parameters:");
        PrintRows(df.SelectRowsByColname("pets", "cat", 3));

        // LINQ.groupby
        Console.WriteLine("LINQ.groupby 'pets':");
        df.GroupRowsByColname("pets");
        Console.WriteLine("LINQ.groupby 'names':");
        df.GroupRowsByColname("names");
        
        // LINQ.select
        Console.WriteLine("LINQ.select 'names':");
        df.SelectItemsByColname("names");

        // создание списка с добавлением шаблона к элементу
        Console.WriteLine("LINQ.select 'names' + постфикс '_item':");
        var list = df.AppendPostfixToColname("names", "_item");
        foreach (var l in list)
            Console.Write($"{l}\t");
        Console.WriteLine();
        Console.WriteLine();

        // Удаление дубликатов строк
        Console.WriteLine("Удаление дубликатов строк из таблицы:");
        var newtable = df.DeleteDuplicateRows();
        PrintTableOrView(newtable);

        Console.WriteLine("end.");
        //Console.ReadLine();
        
        
    }
}