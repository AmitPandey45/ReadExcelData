using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcelData.Extensions
{
    public static class DataTableExtensions
    {
        public static IEnumerable<DataColumn> GetColumns(this DataTable source)
        {
            return source.Columns.AsEnumerable();
        }

        public static IEnumerable<string> GetColumnNames(this DataTable source)
        {
            return source.Columns.AsEnumerable().Select(s => s.ColumnName);
        }
    }
}
