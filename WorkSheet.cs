using Microsoft.Office.Interop.Excel;

namespace Excel.NET
{
    public class WorkSheet
    {
        private readonly Worksheet _worksheet;
        private readonly WorkBook _book;

        public WorkSheet(Worksheet worksheet, WorkBook book)
        {
            _worksheet = worksheet;
            _book = book;
        }

        public Range Range(string start)
            => _worksheet.Range[start];

        public Range Range(string start, string end)
            => _worksheet.Range[start, end];

        public string Name
        {
            get => _worksheet.Name;
            set
            {
                _book._worksheets.Remove(Name);
                _worksheet.Name = value;
                _book._worksheets.Add(Name, this);
            }
        }
    }
}
