using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using DataReaderValidatorAndUploader;

namespace DataReaderValidatorAndUploaderApp
{
    class Program
    {
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

        public static void Main(string[] args)
        {
            string path = ConfigurationManager.AppSettings["excelSourcePath"].ToString();
            string search = ConfigurationManager.AppSettings["extension"].ToString();
            string database = ConfigurationManager.AppSettings["database"].ToString();
            Logger.Info("Starting validation");
            Logger.Info($"Looking for {search} files in {Regex.Unescape(path)}");

            ExcelDataReaderFile ExcelDataReaderFile = new ExcelDataReaderFile();
            string[] fileEntries = Directory.GetFiles(path, "*" + search + "*", SearchOption.AllDirectories);

            Logger.Info($"Found {fileEntries.Length} file(s) to validate");

            dynamic counter = 0;
            BasicValidation BasicValidation = new BasicValidation();
            if (!BasicValidation.NoFilesInAFolder(fileEntries, path))
            {
                foreach (string file in fileEntries)
                {
                    BasicValidation basicValidation = new BasicValidation(file, 128, 32767, 128, 128);
                    counter += 1;

                    Logger.Info($"Validating {basicValidation.GetFileName(file)}");

                    if (basicValidation.InvalidColumnNames() || basicValidation.InvalidCellData())
                    {
                        //if the above statments are true then run those methods and log an error
                        Logger.Fatal($"{counter}) {basicValidation.GetFileName(file)} in {file} has failed validation. Check the error log for details.");
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                    }
                    else
                    {
                        ExcelDataReaderFile excelDataReaderFile = new ExcelDataReaderFile();
                        foreach (DataTable dataTable in excelDataReaderFile.GetDataTable(file))
                        {
                            Sql sql = new Sql(dataTable.TableName, dataTable, excelDataReaderFile.GetColumnNames(dataTable), excelDataReaderFile.RowsDataTypes(dataTable));
                            sql.UseExistingDatabaseGet_Table_Columns_Rows_DataTypes(database);
                            Logger.Info($"{counter}) file number {counter} with file path {file} has been uploaded to the database");
                        }
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                    }
                }
            }
            else
            {
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }
        }
    }
}
