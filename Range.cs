using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using EX = Microsoft.Office.Interop.Excel;

namespace Excel.NET
{
    [DebuggerDisplay("{Address}")]
    public class Range : IEnumerable<Cell>
    {
        private readonly EX.Range _range;
        private readonly WorkSheet _sheet;

        public Range(EX.Range range, WorkSheet sheet)
        {
            _range = range;
            _sheet = sheet;
        }

        public Cell this[int row, int column]
            => new Cell((EX.Range)_range[row + 1, column + 1], _sheet);

        public string Address
            => _range.Address.Replace("$", "");

        public int FirstColumn
            => _range.Column;

        public int FirstRow
            => _range.Row;

        public int Width
            => _range.Columns.Count;

        public int Height
            => _range.Rows.Count;

        public int Count
            => _range.Count;

        public void Select()
            => _range.Select();

        public Range Resize(int rowSize, int columnSize)
            => new Range(_range.Resize[rowSize, columnSize], _sheet);

        public Range Offset(int row, int column)
            => new Range(_range.Offset[row, column], _sheet);

        public IEnumerator<Cell> GetEnumerator()
        {
            foreach (EX.Range cell in _range.Cells)
                yield return new Cell(cell, _sheet);
        }

        IEnumerator IEnumerable.GetEnumerator()
            => GetEnumerator();
    }
}
