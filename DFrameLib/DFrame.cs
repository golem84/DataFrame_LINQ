using System.Collections;
using System.Data;
using System.Net.NetworkInformation;
using System.Runtime.CompilerServices;
// нельзя в библиотеку добавить ссылку для использования COM объектов.

namespace DFrameLib
{

    

    public class DFrame:DataTable
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
        
        // печать массива DataRow
        private void PrintRows(DataRow[] rows)
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

        // выбор столбцов по именам
        // используем копию исходной таблицы
        public DataTable SelectColByName(params string[] names)
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
            return v.ToTable();
        }

        // короткая версия метода выбора столбцов
        public DataTable SelectColumns(params string[] names)
        {
            DataView view = new DataView(this);
            DataTable values;
            values = view.ToTable(true, names);
            return values;
        }

        // печать массива строк
        

        // выборка строк по условию
        public DataRow[] SelectRows(string expr)
        {
            DataRow[] t;
            return t = this.Select(expr);
        }

        // выборка строк по условию, сортировка по условию
        public DataRow[] SelectRows(string expr, string sort)
        { 
            DataRow[] t; 
            t = this.Select(expr, sort);
            return t;
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
        public DataRow[] SelectRowsByColname(string colname, string s)
        {
            var query = this.AsEnumerable().Where(x => x.Field<string>(colname) == s);
            DataRow[] t = query.ToArray();
            return t;
        }

        // 2 logic parameters, вторая форма записи запроса LINQ
        public DataRow[] SelectRowsByColname(string colname, string s, int n)
        {
            var query = from r in this.AsEnumerable()
                        where (string)r[colname] == s && (int)r["PetAge"] > n
                        select r;
            DataRow[] t = query.ToArray();
            return t;
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
        // метод расширения SELECT преобразует результаты выбора в новый формат
        // (здесь - в строковый массив, можно преобразовать в свой тип, схожий со структурой)
        public void SelectItemsByColname(string colname)
        {
            var query = this.AsEnumerable().Select(x => x[colname]).ToArray();
            foreach (var q in query)
                Console.WriteLine($"{q}");
            Console.WriteLine();
        }

        public List<string> AppendPostfixToColname(string colname, string fix)
        {
            var query = this.AsEnumerable().Select(x => x[colname]).ToList();
            var newlist = new List<string>();
            foreach (var q in query)
                newlist.Add(q.ToString()+fix);
            return newlist;
        }

        public DataTable DeleteDuplicateRows()
        {
            var UniqueRows = this.AsEnumerable().Distinct(DataRowComparer.Default);
            return UniqueRows.CopyToDataTable();
        }

        public void DeleteDuplicateColumns()
        {
            Comparer<DataColumn> defComp = Comparer<DataColumn>.Default;


            var UniqueColumns = this.AsEnumerable().Distinct(DataRowComparer.Default);
        }
    }
}