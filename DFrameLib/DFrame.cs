using System.Data;
using System.Runtime.CompilerServices;
// нельзя в библиотеку добавить ссылку для использования COM объектов.

namespace DFrameLib
{
    public class DFrame : DataTable
    {
        // добавить колонку:
        // * именованная типизированная пустая,
        // * именованная с массивом данных без проверки на тип??? - исправить?
        public void AddCol(string _name, Type _type) => this.Columns.Add(_name, _type);
        public void AddCol(string _name, object[] item)
        {
            this.Columns.Add(_name);
            this.AddItemToCol(_name, item);
        }
        // добавить item в столбец colname
        public void AddItemToCol(string colname, object[] item)
        {
            DataRow newrow;
            for (int i = 0; i < item.Length; i++)
            {
                newrow = this.NewRow();
                try
                {
                    newrow[colname] = item[i];
                    this.Rows.Add(newrow);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ошибка. ", ex.Message);
                    return;
                }
            }
        }
        // Добавляем массив объектов как новую строку в таблицу
        public void AddRow(object[] row) => this.Rows.Add(row);
        // печать таблицы (синие заголовки)
        public void PrintTable()
        {
            if (this.Columns.Count == 0)    // если нет данных, не выводим
            {
                Console.WriteLine("This view has no columns. Nothing to display.");
                return;
            }
            Console.ForegroundColor = ConsoleColor.Blue;
            for (int i = 0; i < this.Columns.Count; i++)
                Console.Write($"{this.Columns[i].ColumnName}\t");
            Console.WriteLine();
            Console.ResetColor();
            DataRow[] rows = this.Select();
            PrintRows(rows);
        }
        // печать отображения (красные заголовки)
        public void PrintView(DataView t)
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

        // выбор столбцов по именам
        // используем копию исходной таблицы
        public void SelectColByName(params string[] names)
        {
            var allnames = "";
            foreach (string name in names) allnames += name;
            DataTable newtable = this.Clone(); // копируем структуру исходной таблицы в новую таблицу
            foreach (DataRow row in this.Rows) // копируем данные из исходной таблицы
            {
                var newrow = newtable.Rows.Add();
                for (int i = 0; i < row.Table.Columns.Count; i++) newrow[i] = row[i];
            }
            DataView v = new DataView(newtable); // для новой таблицы создаем отображение
        newsearch:
            for (int i = 0; i < v.Table.Columns.Count; i++) // ищем по новой таблице стобцы
            {
                if (!allnames.Contains(v.Table.Columns[i].ColumnName))
                {
                    v.Table.Columns.RemoveAt(i);    // удаляем ненужные столбцы
                    goto newsearch;                 // смотрим таблицу сначала, поскольку она изменилась
                }
            }
            PrintView(v);
        }

        // короткая версия метода выбора столбцов
        public void SelectColumns(params string[] names)
        {
            DataView view = new DataView(this);
            DataTable values;
            try
            {
                values = view.ToTable(true, names);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return;
            }
            view = new DataView(values);
            PrintView(view);
        }

        // печать массива строк
        private void PrintRows(DataRow[] rows )
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

        // выборка строк по условию
        public void SelectRows(string expr)
        {
            DataRow[] rows;
            rows = this.Select(expr);
            PrintRows(rows);
        }

        // выборка строк по условию, сортировка по условию
        public void SelectRows(string expr, string sort)
        {
            var rows = this.Select(expr, sort);
            PrintRows(rows);
        }

        // переименование колонок на основе словаря
        public void RenameColumns(Dictionary<string, string> newnames)
        {
            foreach (var n in newnames)
            {
                for (int i = 0; i < this.Columns.Count; i++)
                    if (this.Columns[i].ColumnName == n.Key) this.Columns[i].ColumnName = n.Value;
            }
        }

        // using LINQ.where
        // метод расширения WHERE выбирает данные и передает в виде коллекции DataRow исходной таблицы
        // для удобства вывода коллекцию DataRow преобразуем в массив DataRow, метод для вывода реализован выше.
        // 1 logic parameter
        public void SelectRowsByColname(string colname, string s)
        {
            var query = this.AsEnumerable().Where(x => x.Field<string>(colname) == s);
            PrintRows(query.ToArray());
        }

        // 2 logic parameters
        public void SelectRowsByColname(string colname, string s, int n)
        {
            var query = this.AsEnumerable().Where(x => x.Field<string>(colname) == s && x.Field<int>("PetAge")> n);
            PrintRows(query.ToArray());
        }

        // using LINQ.Groupby
        public void GroupRowsByColname(string colname)
        {
            var query = this.AsEnumerable().GroupBy(x => x.Field<string>(colname));
            foreach (var q in query)
            {
                Console.WriteLine($"{q.Key}");
                PrintRows(q.ToArray());
            }
        }

        // using LINQ.select
        // метод расширения SELECT преобразует результаты выбора в новый формат (здесь - в строковый массив)
        public void SelectItemsByColname(string colname)
        {
            var query = this.AsEnumerable().Select(x => x.Field<string>(colname)).ToArray();
            foreach (var q in query)
                Console.WriteLine($"{q}");
        }
        
    }
}