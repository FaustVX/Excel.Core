using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace Excel.NET
{
    [DebuggerDisplay("{Name}")]
    public class WorkSheet
    {
        private readonly Worksheet _worksheet;
        private readonly WorkBook _book;

        public WorkSheet(Worksheet worksheet, WorkBook book)
        {
            _worksheet = worksheet;
            _book = book;
        }

        public Cell Range(string start)
            => new Cell(_worksheet.Range[start], this);

        public Range Range(string start, string end)
            => new Range(_worksheet.Range[start, end], this);

        public Range Range(string start, int rowSize, int columnSize)
            => Range(start).Resize(rowSize, columnSize);

        public string Name
        {
            get => _worksheet.Name;
            set
            {
                if (Name == value)
                    return;

                if (_book._worksheets.ContainsKey(value))
                    throw new System.ArgumentException($"\"{value}\" is already used.");

                _book._worksheets.Remove(Name);
                _worksheet.Name = value;
                _book._worksheets.Add(Name, this);
            }
        }

        public void Select()
            => _worksheet.Select();
    }
}
