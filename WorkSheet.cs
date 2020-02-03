using Microsoft.Office.Interop.Excel;

namespace Excel.NET
{
    public class WorkSheet
    {
        private readonly Worksheet _worksheet;

        public WorkSheet(Worksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public Range Range(string start)
            => _worksheet.Range[start];

        public Range Range(string start, string end)
            => _worksheet.Range[start, end];
    }
}
