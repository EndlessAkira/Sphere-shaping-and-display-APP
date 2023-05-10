using Aspose.Cells;
using System;
using System.Windows.Forms;

namespace SaveSystem
{
    class ExcelSave: IDisposable
    {
        private Aspose.Cells.License _lic;
        private Aspose.Cells.Workbook _workbook;
        private Aspose.Cells.Worksheet _worksheet;
        private Aspose.Cells.Style _style;
        private string _filePath;

        public ExcelSave()
        {
            //_lic = new Aspose.Cells.License();
            //_lic.SetLicense("Aspose.Cells.licence");
        }
        internal bool Open(string filePath)
        {
            try
            {
                _workbook = new Aspose.Cells.Workbook(filePath); 
                _filePath = filePath;
                _worksheet = _workbook.Worksheets[0];

                // Очищаем файл
                _worksheet.Cells.Clear();

                // Настройка размера ячеек
                _workbook.Worksheets[0].Cells.StandardWidth = 35;
                _workbook.Worksheets[0].Cells.StandardHeight = 35;

                // Настройка формата ячеек
                _style = _worksheet.Cells.Style;
                _style.VerticalAlignment = TextAlignmentType.Center;
                _style.HorizontalAlignment = TextAlignmentType.Center;
                _style.Font.Name = "Cascadia Code";
                _style.Font.IsBold = true;
                _style.Font.Size = 16;

                // Применяем настроенный стиль ко всем яйчекам
                _worksheet.Cells.Style = _style;
                return true;
            }catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return false;
        }
        internal void SetText(string place, object data)
        {
            try
            {
                _worksheet.Cells[place].Value = data;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        internal void SetTime(string place, DateTime data)
        {
            try
            {
                string hour = CorrectTime(data.Hour);
                string minute = CorrectTime(data.Minute);
                string second = CorrectTime(data.Second);
                _worksheet.Cells[place].Value = hour + ":" + minute + ":" + second;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        internal void SetImage(int column, int row, string path)
        {
            try
            {
                // Добавляем изображение по ссылке в заданный столбец и строку
                _worksheet.Pictures.Add(row, column, path);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        private string CorrectTime(int value)
        {
            if(value < 10)
                return "0" + value.ToString();
            else 
                return value.ToString();
        }
        internal void Save()
        {
            if (!string.IsNullOrEmpty(_filePath))
            {
                // Сохраняем файл
                try
                {
                    _workbook.Save(_filePath, Aspose.Cells.SaveFormat.Xlsx);
                    MessageBox.Show("Файл успешно сохранён");
                    _filePath = null;
                    Dispose();
                }
                catch
                {
                    MessageBox.Show("Произошла ошибка сохранения! Закройте файл и повторите попытку");
                }
            }
        }
        public void Dispose()
        {
            try
            {
                _workbook.Dispose();
                _worksheet.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
