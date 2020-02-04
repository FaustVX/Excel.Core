using System;
using System.Diagnostics;

namespace Excel.NET
{
    [DebuggerDisplay("{Object}")]
    public class Value
    {
        private readonly Cell _cell;

        public Value(Cell cell)
        {
            _cell = cell;
        }

        public object Object
        {
            get => _cell._range.Value;
            set => _cell._range.Value = value;
        }

        public string String
        {
            get => Object?.ToString();
            set => Object = value;
        }

        public double Double
        {
            get => (double)Object;
            set => Object = value;
        }

        public int Int
        {
            get => (int)Double;
            set => Double = value;
        }

        public static implicit operator string(Value value)
            => value.String;
    }
}
