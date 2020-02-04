using System;
using System.Collections.Generic;
using EX = Microsoft.Office.Interop.Excel;

namespace Excel.NET
{
    public readonly struct Cell
    {
        private readonly EX.Range _range;
        private readonly WorkSheet _sheet;

        public Cell(EX.Range range, WorkSheet sheet)
        {
            _range = range;
            _sheet = sheet;
        }

        public int Column
            => _range.Column;

        public int Row
            => _range.Row;

        public object Value
        {
            get => _range.Value;
            set => _range.Value = value;
        }

        public string ValueString
        {
            get => (string)Value;
            set => Value = value;
        }

        public double ValueDouble
        {
            get => (double)Value;
            set => Value = value;
        }

        public int ValueInt
        {
            get => (int)ValueDouble;
            set => ValueDouble = value;
        }

        public string Formula
        {
            get => (string)_range.Formula;
            set => _range.Formula = value;
        }

        public void Select()
            => _range.Select();

        public Range Resize(int rowSize, int columnSize)
            => new Range((EX.Range)_range[rowSize, columnSize], _sheet);

        public Cell Offset(int row, int column)
            => new Cell(_range.Offset[row, column], _sheet);
    }
}
