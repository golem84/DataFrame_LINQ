using System.Data;
// нельзя в библиотеку добавить ссылку для использования COM объектов.

namespace DFrameLib
{
    public class DFrame : DataTable
    {
        // добавить колонку (именованная типизированная, именованная с массивом данных)
        //public void AddCol(string _name) => this.Columns.Add(_name);
        public void AddCol(string _name, Type _type) => this.Columns.Add(_name, _type);
        public void AddCol(string _name, object[] item)
        {
            this.Columns.Add(_name);
            this.AddItemToCol(_name, item);
        }
        
        public void AddItemToCol(string colname, object[] item) // добавить item в столбец colname
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
        
        



    }
}