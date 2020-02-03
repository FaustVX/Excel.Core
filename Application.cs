using System.Collections.Generic;
using EX = Microsoft.Office.Interop.Excel;
using System.IO;


namespace Excel.NET
{
    public static class Application
    {
        private static readonly EX.Application _app = new EX.Application();

        private static readonly Dictionary<string, WorkBook> _workbooks = new Dictionary<string, WorkBook>();

        public static WorkBook OpenWorkBook(string fileName)
        {
            var path = new FileInfo(fileName);
            if (_workbooks.TryGetValue(path.Name, out var wb))
                return wb;
            return _workbooks[path.Name] = new WorkBook(_app.Workbooks.Open(path.FullName));
        }

        public static WorkBook GetWorkBook(string name)
            => _workbooks[name];
    }
}
