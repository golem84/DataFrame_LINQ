using DFrameLib;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Linq;
using System;
using System.IO;

internal class Program
{
    // печать массива строк без заголовков таблиц
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

    private static class CSVClass
    {
        public static void ToCSV(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers   
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(";");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {

                    string value = dr[i].ToString();
                    if (value.GetType() == typeof(DateTime))
                    {
                        DateTime d = (DateTime)dr[i];
                        value = d.ToShortDateString();
                        sw.Write(value);
                    }
                    else
                    {
                        sw.Write(dr[i].ToString());
                    }

                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(";");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

    }

    static void GetDataFromExcel(string fname, ref DFrame df1)
    {
        // подключаемся к Excel

        
        Excel.Application exApp = new Excel.Application();
        System.Diagnostics.Process excelProc = 
            System.Diagnostics.Process.GetProcessesByName("EXCEL").Last();
        if (exApp == null)
        {
            Console.WriteLine("Excel is not installed!");
            return ;
        }

        // файл находится в одной папке с программой
        string path = Path.GetDirectoryName(typeof(Program).Assembly.Location);

        // настраиваем переменные для работы с Excel
        Excel.Workbook wb = exApp.Workbooks.Open(path + $@"\{fname}");
        Excel.Worksheet ws = wb.Sheets[1]; // нумерация листов начинается с 1
        int maxrow = ws.UsedRange.Rows.Count;
        int maxcol = ws.UsedRange.Columns.Count;

        df1 = new DFrame();

        // Console.WriteLine("Создаем колонки, присваиваем тип данных.");
        //List<string> colnames = new List<string>();
        for (int i = 1; i <= maxcol; i++) 
            df1.AddCol(ws.Cells[1, i].Value.ToString(), 
                ws.Cells[2, i].Value.GetType());
        // Console.WriteLine("Читаем данные в объект DFrame.");

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
        return ;
    }

    private static DFrame df2 = new DFrame();
    
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
        PrintTableOrView(df);
        
        Console.WriteLine("Создаем и заполняем новую таблицу из Excel");
        GetDataFromExcel("Book1.xlsx", ref df2);
        PrintTableOrView(df2);

        Console.WriteLine("Объединение таблиц при помощи метода 'Merge' невозможно, поскольку типы данных у таблиц различны.");
        Console.WriteLine("Далее работаем c первой таблицей, созданной из программы...");

        //df.Merge(df2);
        //PrintTableOrView(df);

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

        // LINQ.where
        Console.WriteLine("LINQ.where 1 logic parameter:");
        PrintRows(df.SelectRowsByColname("Pet", "cat"));

        Console.WriteLine("LINQ.where 2 logic parameters:");
        PrintRows(df.SelectRowsByColname("Pet", "cat", 3));

        // LINQ.groupby
        Console.WriteLine("LINQ.groupby 'Pet':");
        df.GroupRowsByColname("Pet");
        Console.WriteLine("LINQ.groupby 'Name':");
        df.GroupRowsByColname("Name");
        
        // LINQ.select
        Console.WriteLine("LINQ.select 'Name':");
        df.SelectItemsByColname("Name");

        // создание списка с добавлением шаблона к элементу
        Console.WriteLine("LINQ.select 'Name' + постфикс '_item':");
        var list0 = df.AppendPostfixToColname("Name", "_item");
        foreach (var l in list0)
            Console.Write($"{l}  ");
        Console.WriteLine();
        Console.WriteLine();

        // Удаление дубликатов строк
        Console.WriteLine("Удаление дубликатов строк из таблицы:");
        //var newtable = df.DeleteDuplicateRows();
        //PrintTableOrView(newtable);
        df.DeleteDuplicateRows();
        PrintTableOrView(df);

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

        // 
        // дополнительные задания
        //

        // вывод DataTable в файл CSV
        Console.WriteLine("Выводим таблицу в файл CSV, сортировка по столбцу 'DateBirth'.");
        string path = Path.GetDirectoryName(typeof(Program).Assembly.Location) + @"\datatable.csv";
        DataView view = df.DefaultView;
        view.Sort = "DateBirth ASC";
        CSVClass.ToCSV(view.ToTable(), path);
        Console.WriteLine();

        // работа со списками
        Console.WriteLine("Списки значений");
        var list = new List<int>();
        Random rnd = new Random();
        for (int i = 0; i < 280; i++)
            //list.Add(rnd.Next(0,200));
            list.Add(i);

        //foreach (var i in list1) Console.Write($"{i} ");
        var list1 = new List<int>();
        var list2 = new List<int>();

        for (int i = 99; i < list.Count; i += 100)
        {
            list1.AddRange(list.GetRange(i - 50, 50));
            list2.AddRange(list.GetRange(i + 1, 50));
        }
        /*
        Console.WriteLine("list1:");
        foreach (var i in list1) Console.Write($"{i} ");
        Console.WriteLine();
        Console.WriteLine("list2:");
        foreach (var i in list2) Console.Write($"{i} ");
        Console.WriteLine();
        */
        Console.WriteLine("result:");
        for (int i = 0; i < list1.Count; i++)
            Console.Write(Math.Abs(list1[i] - list2[i]) + "  ");
        Console.WriteLine();

        Console.WriteLine("Конец программы.");
        Console.ReadLine();
    }
}