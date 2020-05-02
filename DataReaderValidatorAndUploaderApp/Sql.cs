using Microsoft.SqlServer.Management.Smo;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace DataReaderValidatorAndUploader
{
    public class Sql
    {
        private Database db;
        private Table tb;
        private Server srv;
        private readonly String SheetName;
        private List<dynamic> ColumnNames;
        private List<dynamic> ColumnsDataTypes;
        private readonly DataTable DataTable;
        private static readonly string server;

        static Sql()
        {
            server = System.Configuration.ConfigurationManager.AppSettings["server"];
        }

        public Sql(String sheetName, DataTable dataTable, List<dynamic> columnNames, List<dynamic> DataTypes)
        {
            SheetName = sheetName;
            DataTable = dataTable;
            ColumnNames = columnNames;
            ColumnsDataTypes = DataTypes;
        }

        public void CreateDbGet_Table_Columns_Rows_DataTypes(string database)
        {
            srv = new Server(server);
            DropDatabaseIfExists(database);
            db = new Database(srv, database);
            db.Create();
            ConvertExcelDataTypesToSql();
            CreateTable();
            CreateRows(database);
        }


        public Database UseExistingDatabaseGet_Table_Columns_Rows_DataTypes(string database)
        {
            string server = System.Configuration.ConfigurationManager.AppSettings["server"];
            srv = new Server(server);
            db = srv.Databases[database];
            ConvertExcelDataTypesToSql();
            CreateTable();
            CreateRows(database);
            return db;
        }


        public Table CreateTable()
        {
            tb = new Table(db, SheetName);
            DropTableIfExists();
            CreateColumns();
            tb.Create();
            return tb;
        }


        public void ConvertExcelDataTypesToSql()

        {
            for (dynamic columnDataTypeIndexNumber = 0; columnDataTypeIndexNumber < ColumnsDataTypes.Count; columnDataTypeIndexNumber++)
            {
                if (ColumnsDataTypes[columnDataTypeIndexNumber] == typeof(String))
                {
                    ColumnsDataTypes[columnDataTypeIndexNumber] = DataType.NVarChar(255);
                }
                else
                {
                    ColumnsDataTypes[columnDataTypeIndexNumber] = DataType.Float;
                }
            }
        }

        public void DropTableIfExists()
        {
            bool tableExists = db.Tables.Contains(SheetName);
            if (tableExists)
            {
                db.Tables[SheetName].Drop();
            }
        }

        public void DropDatabaseIfExists(string database)
        {
            Boolean databaseExists = srv.Databases.Contains(database);
            if (databaseExists)
            {
                srv.Databases[database].Drop();
            }
        }

        public void CreateColumns()
        {
            for (dynamic indexNumber = 0; indexNumber < Math.Min(ColumnNames.Count, ColumnsDataTypes.Count); indexNumber++)
            {
                Column column;
                column = new Column(tb, ColumnNames[indexNumber], ColumnsDataTypes[indexNumber])
                {
                    Collation = "Latin1_General_CI_AS",
                    Nullable = true
                };
                tb.Columns.Add(column);
            }
        }

        public void CreateRows(string database)
        {
            string sqlConnectionString = ConfigurationManager.ConnectionStrings["MyKey"].ConnectionString;
            sqlConnectionString = string.Format(sqlConnectionString, server, database);
            using (var bulkCopy = new SqlBulkCopy(sqlConnectionString))
            {
                bulkCopy.DestinationTableName = tb.Name.ToString();
                bulkCopy.WriteToServer(DataTable);
            }
        }
    }
}
