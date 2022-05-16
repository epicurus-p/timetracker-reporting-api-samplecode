using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace TimetrackerReportingClient
{
    internal class TimesheetData
    {
            public static DataTable CreateDataTable()
            {
                DataTable dataTable = new DataTable();

                DataColumn dataColumn;
                dataColumn = new DataColumn();
                dataColumn.DataType = Type.GetType("System.String");
                dataColumn.ColumnName = "User";
                dataTable.Columns.Add(dataColumn);

                dataColumn = new DataColumn();
                dataColumn.DataType = Type.GetType("System.String");
                dataColumn.ColumnName = "Date";
                dataTable.Columns.Add(dataColumn);

                dataColumn = new DataColumn();
                dataColumn.DataType = Type.GetType("System.String");
                dataColumn.ColumnName = "Hours";
                dataTable.Columns.Add(dataColumn);

                dataColumn = new DataColumn();
                dataColumn.DataType = Type.GetType("System.String");
                dataColumn.ColumnName = "Discipline";
                dataTable.Columns.Add(dataColumn);

                dataColumn = new DataColumn();
                dataColumn.DataType = Type.GetType("System.String");
                dataColumn.ColumnName = "RnD Type";
                dataTable.Columns.Add(dataColumn);

                dataColumn = new DataColumn();
                dataColumn.DataType = Type.GetType("System.String");
                dataColumn.ColumnName = "Job Code";
                dataTable.Columns.Add(dataColumn);

                dataColumn = new DataColumn();
                dataColumn.DataType = Type.GetType("System.String");
                dataColumn.ColumnName = "Work Item Id";
                dataTable.Columns.Add(dataColumn);

                dataColumn = new DataColumn();
                dataColumn.DataType = Type.GetType("System.String");
                dataColumn.ColumnName = "Project";
                dataTable.Columns.Add(dataColumn);

                return dataTable;
            }
        
    }
}
