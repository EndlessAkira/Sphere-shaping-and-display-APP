using System;
using System.Collections.Generic;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace SaveSystem
{
    class WordSave
    {
        private FileInfo _fileInfo;

        public WordSave(string fileName)
        {
            if (File.Exists(fileName))
            {
                _fileInfo = new FileInfo(fileName);
            } 
            else
            {
                throw new ArgumentException("File not found");
            }
        }

        internal bool SetText(List<List<string>> list)
        {
            Word.Application word = new Word.Application();
            Word.Table table = word.Application.ActiveDocument.Tables.Add(word.Selection.Range, list.Count, 2);

            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle; 

            for(int i = 0; i < list.Count; i++)
            {
                for (int j = 0; j < 2; j++)
                {
                    table.Cell(i, j).Range.Text = list[i][j];
                }
            }
            //
            return false;
        }
    }
}
