using System;
using System.Collections.Generic;
using EX = Microsoft.Office.Interop.Excel;

namespace Excel.NET
{
    public readonly struct Range
    {
        private readonly EX.Range _range;
        private readonly WorkSheet _sheet;

        public Range(EX.Range range, WorkSheet sheet)
        {
            _range = range;
            _sheet = sheet;
        }

        public Cell this[int row, int column]
            => new Cell((EX.Range)_range[row, column], _sheet);

        public int FirstColumn
            => _range.Column;

        public int FirstRow
            => _range.Row;

        public int Width
            => (int)_range.Width;

        public int Height
            => (int)_range.Height;

        public Range Resize(int rowSize, int columnSize)
            => new Range((EX.Range)_range[rowSize, columnSize], _sheet);

        public Range Offset(int row, int column)
            => new Range(_range.Offset[row, column], _sheet);
    }
}
