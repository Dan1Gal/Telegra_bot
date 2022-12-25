using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace TelegramBot_Kurs.BD
{
    class ExcelHelper : IDisposable
    {
        private Excel.Application _excel;
        private Excel.Workbook _workbook;
        private string _filePath;

        public ExcelHelper()
        {
            _excel = new Excel.Application();
        }

        internal bool Open(string filePath)//открытие или создание таблицы
        {
            try
            {
                if (File.Exists(filePath))
                {
                    _workbook = _excel.Workbooks.Open(filePath);
                }
                else
                {
                    _workbook = _excel.Workbooks.Add();
                    _filePath = filePath;
                }

                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }

        internal void Save()//сохранение таблицы
        {
            if (!string.IsNullOrEmpty(_filePath))
            {
                _workbook.SaveAs(_filePath);
                _filePath = null;
            }
            else
            {
                _workbook.Save();
                _workbook.Close();
                _excel.Quit();
            }
        }

        internal bool Set(string column, int row, string data)//определение данных
        {
            try
            {
                //var val = ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column].Value2;

                ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column] = data;
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }

        internal object Get(string column, int row)//получение данных
        {
            try
            {
                {
                    return ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column].Value2;
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return null;
        }

        public void Dispose()//закрытие и выход из таблицы
        {
            try
            {
                _excel.Quit();
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        internal bool Read(string column, int row)//чтение ячеек
        {
            throw new NotImplementedException();
        }
       
    }
}
