using System;
using System.Collections.Generic;
using System.Diagnostics;
using EX = Microsoft.Office.Interop.Excel;

namespace Excel.NET
{
    [DebuggerDisplay("{Address}: {Value.String}")]
    public class Cell
    {
        internal readonly EX.Range _range;
        private readonly WorkSheet _sheet;

        public Cell(EX.Range range, WorkSheet sheet)
        {
            _range = range;
            _sheet = sheet;
            Value = new Value(this);
        }

        public Value Value { get; }

        public string Address
            => _range.Address.Replace("$", "");

        public int Column
            => _range.Column;

        public int Row
            => _range.Row;

        public double RowHeight
        {
            get => (double)_range.RowHeight;
            set => _range.RowHeight = value;
        }

        public double ColumnWidth
        {
            get => (double)_range.ColumnWidth;
            set => _range.ColumnWidth = value;
        }

        public string Formula
        {
            get => (string)_range.Formula;
            set => _range.Formula = value;
        }

        public void Select()
            => _range.Select();

        public Range Resize(int rowSize, int columnSize)
            => new Range(_range.Resize[rowSize, columnSize], _sheet);

        public Cell Offset(int row, int column)
            => new Cell(_range.Offset[row, column], _sheet);
    }
}
