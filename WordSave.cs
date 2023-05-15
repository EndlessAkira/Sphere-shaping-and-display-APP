using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Word.Application;

namespace SaveSystem
{
    public class WordSave
    {
        private string fileName;
        private string path;
        private Application _wordApp;
        private Document _document;
        private Table _table;
        public bool Open(string filePath)
        {
            try
            {
                // Инициализация Word
                _wordApp = new Application();
                _document = _wordApp.Documents.Open(filePath);
                fileName = Path.GetFileName(filePath);
                return true;
            }
            catch
            {
                return false;
            }
        }
        public void AddData(List<List<string>> data)
        {
            _table = _document.Tables.Add(_document.Range(), data.Count, data[0].Count);

            // Заполнение таблицы данными
            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < data[i].Count; j++)
                {
                    _table.Cell(i + 1, j + 1).Range.Text = data[i][j];
                }
            }
        }
        public void AddImage(string imagePath)
        {
            try
            {
                // Вставка изображения
                Range range = _document.Range();
                range = range.GoTo( Microsoft.Office.Interop.Word.WdGoToItem.wdGoToLine , Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToLast);
                range.InsertParagraphAfter();
                range.InlineShapes.AddPicture(imagePath);
                range.InsertParagraphAfter();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка при вставке изображения: " + ex.Message);
            }
        }
        public void Save()
        {
            try
            {
                // Сохранение документа

                _document.Save();
                _document.Close();
                _wordApp.Quit();

                MessageBox.Show($"Таблица успешно сохранена в файле {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Произошла ошибка при сохранении таблицы: " + ex.Message);
            }
        }


    }
}
