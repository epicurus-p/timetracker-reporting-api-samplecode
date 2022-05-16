using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace TimetrackerReportingClient
{
    internal class CsvParser
    {
        /// <summary>
        /// process writing the csv file
        /// </summary>
        public class CsvWriter
        {
            public string WriteToString(DataTable table, bool header, bool quoteall)
            {
                StringWriter writer = new StringWriter();
                WriteToStream(writer, table, header, quoteall);
                return writer.ToString();
            }

            public void WriteToStream(TextWriter stream, DataTable myDataTable, bool header, bool quoteall)
            {
                if (header)
                {
                    for (int i = 0; i < myDataTable.Columns.Count; i++)
                    {
                        WriteItem(stream, myDataTable.Columns[i].Caption, quoteall);
                        if (i < myDataTable.Columns.Count - 1)
                            stream.Write(',');
                        else
                            stream.Write('\n');
                    }
                }

                int rowCount = myDataTable.Rows.Count;
                int j = 1;

                foreach (DataRow row in myDataTable.Rows)
                {
                    for (int i = 0; i < myDataTable.Columns.Count; i++)
                    {
                        WriteItem(stream, row[i], quoteall);
                        if (i < myDataTable.Columns.Count - 1)
                            stream.Write(',');
                        //add new row except for last data
                        else if (rowCount != j)
                            stream.Write('\n');
                    }
                    j++;
                }
            }

            private void WriteItem(TextWriter stream, object item, bool quoteall)
            {
                if (item == null)
                    return;
                string s = item.ToString();
                if (quoteall || s.IndexOfAny("\",\x0A\x0D".ToCharArray()) > -1)
                    stream.Write("\"" + s.Replace("\"", "\"\"") + "\"");
                else
                    stream.Write(s);
                stream.Flush();
            }
        }

        /// <summary>
        /// process of getting the csv parser
        /// </summary>
        public class GetCsvParser
        {
            public DataTable Parse1(string data, bool headers)
            {
                return Parse1(new StringReader(data), headers);
            }

            public DataTable Parse1(string data)
            {
                return Parse1(new StringReader(data));
            }

            public DataTable Parse1(TextReader stream)
            {
                return Parse1(stream, false);
            }

            public DataTable Parse1(TextReader stream, bool headers)
            {
                DataTable table = new DataTable();

                char[] seperators = new char[] { ',' };

                // process header
                string line = stream.ReadLine();
                if (line != null)
                {
                    string[] cells = line.Split(seperators);
                    foreach (string cell in cells)
                    {
                        table.Columns.Add(cell.Trim(), typeof(string));
                    }
                }

                // process body
                line = stream.ReadLine();
                while (line != null)
                {
                    int index = 0;
                    string[] cells = line.Split(seperators);

                    if (cells.Length > 9)
                    {
                        List<string> list = new List<string>();
                        foreach (string listString in cells)
                        {
                            list.Add(listString);
                        }

                        int i = 0;
                        string[] newCells;
                        newCells = new string[9];

                        //insert list to string array
                        foreach (string lists in list)
                        {
                            newCells[i] = lists;
                            i++;
                        }

                        table.Rows.Add(newCells);
                        line = stream.ReadLine();
                    }
                    else
                    {
                        table.Rows.Add(cells);
                        line = stream.ReadLine();
                    }
                }
                return table;
            }
        }

        /// <summary>
        /// process of createing Data Table
        /// </summary>
        public class CreateDataTable
        {
            public DataTable CreatingDataTable()
            {
                DataTable dataTable = new DataTable();

                DataColumn dataColumn;
                dataColumn = new DataColumn();
                dataColumn.DataType = Type.GetType("System.String");
                dataColumn.ColumnName = "Week Ending";
                dataTable.Columns.Add(dataColumn);

                dataColumn = new DataColumn();
                dataColumn.DataType = Type.GetType("System.String");
                dataColumn.ColumnName = "Employee Number";
                dataTable.Columns.Add(dataColumn);

                dataColumn = new DataColumn();
                dataColumn.DataType = Type.GetType("System.String");
                dataColumn.ColumnName = "Work Date";
                dataTable.Columns.Add(dataColumn);

                dataColumn = new DataColumn();
                dataColumn.DataType = Type.GetType("System.String");
                dataColumn.ColumnName = "Hours";
                dataTable.Columns.Add(dataColumn);

                dataColumn = new DataColumn();
                dataColumn.DataType = Type.GetType("System.String");
                dataColumn.ColumnName = "SAP WBS No.";
                dataTable.Columns.Add(dataColumn);

                return dataTable;
            }
        }
    }
}
