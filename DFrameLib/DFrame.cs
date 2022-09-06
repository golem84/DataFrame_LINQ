using System.Data;
using System;
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
        // печать всей таблицы
        public void PrintTable()
        {
            Console.WriteLine();
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
                // вывод типа столбцов
                /*
                for (int i = 0; i < this.Columns.Count; i++)
                    Console.Write($"{this.Columns[i].DataType}\t");               
                Console.WriteLine();
                */
                Console.WriteLine();
        }

        public void PrintView(DataView t)
        {
            Console.WriteLine();
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





    }
}