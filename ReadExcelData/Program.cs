using IronXL;
using ReadExcelData.Extensions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ReadExcelData
{
    class Program
    {
        const string FullPath = @"D:\MY SOFTWARE AND FILES\CommitteeMemberClassification.xlsx";
        const string ScriptSeparator = "\n\n";
        const string TableName = "db_MEM.CommitteeMemberClassification";
        static Dictionary<string, BusinessRule> BusinessColumns = new Dictionary<string, BusinessRule>
        {
            { "COMMITTEEPRIMARYACTIVITYID", new BusinessRule { ColumnNumber = -1, Rule = null } },
            { "WEBSITE", new BusinessRule { ColumnNumber = -1, Rule = new int[]  { 1, 2, 3, 4, 6, 7 } } },
            { "FACILITYORGANIZATION", new BusinessRule { ColumnNumber = -1, Rule = new int[]  { 1, 2, 3, 4, 5, 6, 7, 11 } } },
            { "PARENTORGANIZATION" , new BusinessRule { ColumnNumber = -1, Rule = new int[]  { 1, 2, 3, 4 } } },
            { "CODIVISION", new BusinessRule { ColumnNumber = -1, Rule = new int[]  { 1, 2, 3, 4, 6, 7 } } }
        };

        static int BusinessDerivedColumnValue = -1;

        static void Main(string[] args)
        {
            CreateTableInsertScript(FullPath);

            Console.ReadKey();
        }

        private static void CreateTableInsertScript(string fullPath)
        {
            string tableColumnInsertScript = string.Empty;
            int columnNumber = 1;
            int totalColumns = 0;
            var tableInsertScript = new StringBuilder();
            var tableRowValueScript = new StringBuilder();

            WorkBook workbook = WorkBook.Load(fullPath);
            WorksheetsCollection workSheets = workbook.WorkSheets;
            foreach (WorkSheet sheet in workSheets)
            {
                DataTable dtWorkSheet = sheet.ToDataTable(true);

                IEnumerable<string> columns = dtWorkSheet.GetColumnNames();
                totalColumns = columns.Count();
                string commaSeparatedcolumns = string.Join(", ", dtWorkSheet.GetColumnNames());

                tableColumnInsertScript = CreateTableColumnInsertScript(commaSeparatedcolumns);

                for (int i = 0; i < dtWorkSheet.Rows.Count; i++)
                {
                    columnNumber = 1;
                    foreach (var column in columns)
                    {
                        CreateTableRowValueScript(tableRowValueScript, column, Convert.ToString(dtWorkSheet.Rows[i][column]), columnNumber, totalColumns);
                        columnNumber++;
                    }

                    tableInsertScript.Append(tableColumnInsertScript);
                    tableInsertScript.Append(tableRowValueScript);
                    tableInsertScript.Append(ScriptSeparator);
                    tableRowValueScript.Clear();
                }
            }

            Console.WriteLine(tableInsertScript.ToString());
        }

        private static string CreateTableColumnInsertScript(string commaSeparatedColumns)
        {
            return $"INSERT INTO {TableName} ({commaSeparatedColumns})";
        }

        private static void CreateTableRowValueScript(StringBuilder rowScript, string column, string columnValue, int columnNumber, int totalColumns)
        {
            if (columnNumber == 1)
            {
                rowScript.Append($"VALUES (");
            }

            string value = GetColumnValue(column, columnValue);

            if (!string.IsNullOrEmpty(value))
            {
                rowScript.Append($"'{value}'");
            }
            else
            {
                rowScript.Append("NULL");
            }

            if (columnNumber != totalColumns)
            {
                rowScript.Append(", ");
            }
            else
            {
                rowScript.Append(")");
            }
        }

        private static string GetColumnValue(string column, string columnValue)
        {
            string value = columnValue;
            if (string.IsNullOrEmpty(columnValue) || columnValue.ToLower().Equals("null"))
            {
                return null;
            }
            else if (columnValue.Contains("’"))
            {
                value = columnValue.Replace("’", "’’");
            }
            else if (columnValue.Contains("'"))
            {
                value = columnValue.Replace("'", "''");
            }

            string key = column?.Trim()?.ToUpper();
            if (BusinessColumns.ContainsKey(key))
            {
                BusinessRule businessRule = BusinessColumns[key];

                if (businessRule != null)
                {
                    if (businessRule.Rule != null)
                    {
                        value = businessRule.Rule.Contains(BusinessDerivedColumnValue) ? columnValue : null;
                    }
                    else
                    {
                        businessRule.BusinessDerivedColumnValue = columnValue;
                        BusinessDerivedColumnValue = Convert.ToInt32(businessRule.BusinessDerivedColumnValue.Trim());
                    }
                }
            }

            return value;
        }

        //public static string BuildExcelConnectionString(string fullPath)
        //{
        //    string connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=\"Excel 8.0;HDR=No;IMEX=1\";", fullPath);
        //    return connectionString;
        //}

        //public static string BuildExcel2007ConnectionString(string fullPath)
        //{
        //    string connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=No;IMEX=1\";", fullPath);
        //    return connectionString;
        //}

        //private static void ReadExcelFile()
        //{
        //    string connStr = BuildExcel2007ConnectionString(@"D:\MY SOFTWARE AND FILES\InsertScript.xlsx");
        //    string query = @"Select * From [Sheet1$] Where Row = 2";
        //    System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(connStr);

        //    conn.Open();
        //    System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand(query, conn);
        //    System.Data.OleDb.OleDbDataReader dr = cmd.ExecuteReader();
        //    DataTable dt = new DataTable();
        //    dt.Load(dr);
        //    dr.Close();
        //    conn.Close();
        //}
    }
}
