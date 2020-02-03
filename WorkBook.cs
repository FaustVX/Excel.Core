using Microsoft.Office.Interop.Excel;

namespace Excel.NET
{
    public class WorkBook
    {
        private readonly Workbook _workbook;

        public WorkBook(Workbook workbook)
        {
            _workbook = workbook;
        }
    }
}
