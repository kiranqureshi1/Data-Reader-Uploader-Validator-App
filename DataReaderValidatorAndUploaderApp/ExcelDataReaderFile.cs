using ExcelDataReader;
using Microsoft.SqlServer.Management.Smo;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataReaderValidatorAndUploader
{
    public class ExcelDataReaderFile
    {
        public IExcelDataReader reader;

        public string GetFolderName()
        {
            string path = ConfigurationManager.AppSettings["path"].ToString();
            DirectoryInfo dir_info = new DirectoryInfo(path);
            string directory = dir_info.Name;
            return directory;
        }

        public dynamic GetDataTable(string fileName)
        {
            using (var reader = ExcelReaderFactory.CreateReader(File.Open(fileName, FileMode.Open, FileAccess.Read)))
            {
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };
                var dataSet = reader.AsDataSet(conf);
                var table = new DataTable();
                return dataSet.Tables;
            }
        }



        public dynamic GetColumnNames(DataTable dataTable)
        {
            List<dynamic> columns = new List<dynamic>();
            for (int ColumnIndexNumber = 0; ColumnIndexNumber < dataTable.Columns.Count; ColumnIndexNumber++)
            {
                var column = dataTable.Columns[ColumnIndexNumber];
                columns.Add(column.ColumnName);
            }
            return columns;
        }

        public dynamic RowsDataTypes(DataTable dataTable)
        {
            List<dynamic> ColumnsDataTypes = new List<dynamic>();
            for (int columnIndexNumber = 0; columnIndexNumber < dataTable.Columns.Count; columnIndexNumber++)
            {
                var column = dataTable.Columns[columnIndexNumber];
                ColumnsDataTypes.Add(column.DataType);
            }
            return ColumnsDataTypes;
        }

        public string GetCellAddress(int column, int row)
        {
            column++;
            StringBuilder stringbuilder = new StringBuilder();
            do
            {
                column--;
                stringbuilder.Insert(0, (char)('A' + (column % 26)));
                column /= 26;

            } while (column > 0);
            stringbuilder.Append(row + 1);
            return stringbuilder.ToString();
        }
    }
}
