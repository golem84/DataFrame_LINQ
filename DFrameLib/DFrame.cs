using System.Data;
using System;
using System.Net.Security;
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
                    return ;
                }
            }
        }
        // Добавляем массив объектов как новую строку в таблицу
        public void AddRow(object[] row ) => this.Rows.Add(row);
        // печать таблицы (синие заголовки)
        public void PrintTable()
        {
            if (this.Columns.Count == 0)    // если нет данных, не выводим
            {
                Console.WriteLine("This view has no columns. Nothing to display.");
                return;
            }

            Console.ForegroundColor=ConsoleColor.Blue;
            for (int i = 0; i<this.Columns.Count; i++)
                Console.Write($"{this.Columns[i].ColumnName}\t");
            Console.WriteLine();
            Console.ResetColor();
            for (int i = 0; i < this.Rows.Count; i++)
            {
                DataRow row = this.Rows[i];
                for (int j = 0; j < row.Table.Columns.Count; j++)
                {
                    if (this.Columns[j].DataType != typeof(DateTime))
                        Console.Write($"{row[j]}\t");
                    else
                    {
                        DateTime d = (DateTime)row[j];
                        Console.Write($"{d.ToShortDateString()}\t");
                    }
                }
                Console.WriteLine();
            }
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
            for (int i = 0; i < t.Table.Rows.Count; i++)
            {
                DataRow row = t.Table.Rows[i];
                for (int j = 0; j < row.Table.Columns.Count; j++)
                {
                    if (t.Table.Columns[j].DataType != typeof(DateTime))
                        Console.Write($"{row[j]}\t");
                    else
                    {
                        DateTime d = (DateTime)row[j];
                        Console.Write($"{d.ToShortDateString()}\t");
                    }
                }
                Console.WriteLine();
            }
            Console.WriteLine();
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
            // можно добавить метод для возврата новой таблицы или ее отображения вместо печати
        }
        public void RenameColumns(Dictionary<string, string> newnames)
        {
            foreach (var n in newnames)
            {
                for (int i = 0; i < this.Columns.Count; i++)
                    if (this.Columns[i].ColumnName == n.Key) this.Columns[i].ColumnName = n.Value;
            }
        }

    }
}