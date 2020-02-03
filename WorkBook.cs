using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Excel.NET
{
    public class WorkBook
    {
        private readonly Workbook _workbook;

        private readonly Dictionary<string, WorkSheet> _worksheets = new Dictionary<string, WorkSheet>();

        public WorkBook(Workbook workbook)
        {
            _workbook = workbook;
        }

        public WorkSheet ActiveSheet
            => GetSheet((Worksheet)_workbook.ActiveSheet);

        public WorkSheet Sheet(string name)
            => GetSheet((Worksheet)_workbook.Worksheets[name]);

        private WorkSheet GetSheet(Worksheet sheet)
            => _worksheets.TryGetValue(sheet.Name, out var s) ? s : _worksheets[sheet.Name] = new WorkSheet(sheet);
    }
}
