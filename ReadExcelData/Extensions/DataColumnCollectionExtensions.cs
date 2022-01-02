using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcelData.Extensions
{
    public static class DataColumnCollectionExtensions
    {
        public static IEnumerable<DataColumn> AsEnumerable(this DataColumnCollection source)
        {
            return source.Cast<DataColumn>();
        }
    }
}
