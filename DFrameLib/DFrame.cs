using System.Data;
// нельзя в библиотеку добавить ссылку для использования COM объектов.

namespace DFrameLib
{
    public class DFrame : DataTable
    {
        // добавить колонку (именованная типизированная, именованная с массивом данных)
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
            for (int i = 0; i<this.Columns.Count; i++)
                Console.Write($"{this.Columns[i].ColumnName}\t");
            Console.WriteLine();
            for (int i = 0; i < this.Rows.Count; i++)
            {
                DataRow row = this.Rows[i];
                for (int j=0;j<row.Table.Columns.Count;j++)
                    Console.Write($"{row[j]}\t");
                Console.WriteLine();
            }                   
            
        }
        



    }
}