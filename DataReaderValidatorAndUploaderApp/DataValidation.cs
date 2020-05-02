using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataReaderValidatorAndUploader
{
    public class BasicValidation
    {
        private readonly string file;
        private IExcelDataReader reader;
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        private readonly int columnNameSizeLimit;
        private readonly int rowDataSizeLimit;
        private readonly int sheetNameSizeLimit;
        private readonly int fileNameSizeLimit;
        private bool errorDetected;
        private dynamic ColumnAlreadyMatchedA;
        private dynamic ColumnAlreadyMatchedB;

        public BasicValidation()
        {
        }
        public BasicValidation(string FileName, int ColumnNamseSizeLimit, int RowDataSizeLimit, int SheetNameSizeLimit, int FileNameSizeLimit)
        {
            file = FileName;
            columnNameSizeLimit = ColumnNamseSizeLimit;
            rowDataSizeLimit = RowDataSizeLimit;
            sheetNameSizeLimit = SheetNameSizeLimit;
            fileNameSizeLimit = FileNameSizeLimit;
        }

        public string GetFileName(string file)
        {
            string fileName = Path.GetFileName(Path.GetFileName(file));
            return fileName;
        }

        public IExcelDataReader ReadAndStreamFile()
        {
            reader = ExcelReaderFactory.CreateReader(File.Open(file, FileMode.Open, FileAccess.Read));
            return reader;
        }

        public Boolean NoFilesInAFolder(string[] fileEntries, string path)
        {
            if (fileEntries.Length == 0)
            {
                Logger.Error($"There are no files in the folder {path} to validate. Review the folder.");
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool SheetNameValidation()
        {
            using (ReadAndStreamFile())
            {
                reader.Read();
                {
                    string SheetName = reader.Name;
                    if (SheetName.Length > sheetNameSizeLimit)
                    {
                        Logger.Error($"[{GetFileName(file)}]{SheetName} exceeds {sheetNameSizeLimit} character sheet name limit. Supply a valid sheet name.");
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
        }


        public bool FileNameValidation()
        {
            if (GetFileName(file).Length > fileNameSizeLimit)
            {
                Logger.Error($"{GetFileName(file)} exceeds {fileNameSizeLimit} character file name limit. Supply a valid file name.");
                return true;
            }
            else
            {
                return false;
            }
        }


        public Boolean DuplicateColumnNames()
        {
            using (ReadAndStreamFile())
            {
                reader.Read();
                {
                    var ColumnsNames = Enumerable.Range(0, reader.FieldCount).Select(i => reader.GetValue(i)).ToList();
                    var excel = new ExcelDataReaderFile();
                    errorDetected = false;
                    //looping through column names
                    for (int columnIndexNumber = 0; columnIndexNumber < ColumnsNames.Count; columnIndexNumber++)
                    {
                        //looping through rows
                        for (int columnIndexNum = 0; columnIndexNum < ColumnsNames.Count; columnIndexNum++)
                        {
                            var cellAddressA = excel.GetCellAddress(columnIndexNumber, 0);
                            var cellAddressB = excel.GetCellAddress(columnIndexNum, 0);
                            //so lets say we are talking about column A and Column C, if ColumnA and Column C are not null 
                            if (ColumnsNames[columnIndexNumber] != null && ColumnsNames[columnIndexNum] != null)
                            {
                                //so if column A were never put against column C to check if they match 
                                if (ColumnsNames[columnIndexNumber] != ColumnAlreadyMatchedA && ColumnsNames[columnIndexNum] != ColumnAlreadyMatchedB)
                                {
                                    //check every column against eachother apart from checking it with itself ( say if columnIndexNumber and columnIndexNum are the same then it emans we are trying to match it with itself)
                                    // any column say ColumnsNames[columnIndexNumber] is column A already matched with B, C and D so going backwards doesnt check column B,C,D against column A
                                    if (ColumnsNames[columnIndexNumber] == ColumnsNames[columnIndexNum] && columnIndexNumber != columnIndexNum)
                                    {
                                        // and so Column a is given  "ColumnAlreadyMatchedA"name
                                        ColumnAlreadyMatchedA = ColumnsNames[columnIndexNumber];
                                        ColumnAlreadyMatchedB = ColumnsNames[columnIndexNum];
                                        Logger.Error($"[{GetFileName(file)}]{reader.Name}!{cellAddressA} with column name {ColumnsNames[columnIndexNumber]} matches {cellAddressB} with column name {ColumnsNames[columnIndexNum]}");
                                        errorDetected = true;
                                    }
                                    else
                                    {
                                        // errorDetected = false;
                                    }
                                }
                            }
                        }
                    }
                }
                reader.Dispose();
                reader.Close();
                return errorDetected;
            }
        }


        public Boolean InvalidColumnNames()
        {
            using (ReadAndStreamFile())
            {
                reader.Read();
                {
                    int counter = 0;
                    var ColumnsNames = Enumerable.Range(0, reader.FieldCount).Select(i => reader.GetValue(i)).ToList();
                    if (ColumnsNames.Count != 0 && reader.Read() == true)
                    {
                        errorDetected = false;
                        for (int columnNumber = 0; columnNumber < ColumnsNames.Count; columnNumber++)
                        {
                            var excel = new ExcelDataReaderFile();
                            var cellAddress = excel.GetCellAddress(counter, 0);
                            counter += 1;
                            if (ColumnsNames[columnNumber] != null && ColumnsNames[columnNumber].ToString().Length > columnNameSizeLimit)
                            {
                                Logger.Error($"[{GetFileName(file)}]{reader.Name}!{cellAddress} is {columnNumber.ToString().Length} characters long and exceeds {columnNameSizeLimit} character column name limit. Supply a valid column name.");
                                errorDetected = true;
                            }
                            else if (ColumnsNames[columnNumber] == null)
                            {
                                Logger.Error($"[{GetFileName(file)}]{reader.Name}!{cellAddress} is empty. Supply a valid column name.");
                                errorDetected = true;
                            }
                            else
                            {
                            }
                            continue;
                        }
                    }
                    else
                    {
                        Logger.Error($"[{GetFileName(file)}]{reader.Name} is empty and cannot be validated. Supply a non-empty file.");
                        errorDetected = true;
                    };
                }
                reader.Dispose();
                reader.Close();
                return errorDetected;
            }
        }


        public bool InvalidCellData()
        {
            using (ReadAndStreamFile())
            {
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };
                var dataSet = reader.AsDataSet(conf);
                var dataTable = dataSet.Tables[0];
                var rows = Enumerable.Range(0, reader.FieldCount).Select(i => reader.Read()).ToArray();
                errorDetected = false;
                for (var ColumnIndexNumber = 0; ColumnIndexNumber < dataTable.Columns.Count; ColumnIndexNumber++)
                {
                    for (var RowIndexNumber = 0; RowIndexNumber < dataTable.Rows.Count; RowIndexNumber++)
                    {
                        //RowIndexNumber is row number starts from index number 0
                        //ColumnIndexNumber is column number starts from index number 0
                        var data = dataTable.Rows[RowIndexNumber][ColumnIndexNumber];
                        var excel = new ExcelDataReaderFile();
                        var cellAddress = excel.GetCellAddress(ColumnIndexNumber, RowIndexNumber + 1);
                        if (data.ToString().Length != 0)
                        {
                            if (data.GetType() == reader.GetFieldType(ColumnIndexNumber) && data.ToString().Length > rowDataSizeLimit)
                            {
                                Logger.Error($"[{GetFileName(file)}]{reader.Name}!{cellAddress} is {data.ToString().Length} characters long and exceeds {rowDataSizeLimit} character cell contents limit. Supply valid cell contents.");
                                errorDetected = true;
                            }
                            else if (data.ToString().Length <= rowDataSizeLimit & data.GetType() != reader.GetFieldType(ColumnIndexNumber))
                            {
                                Logger.Error($"[{GetFileName(file)}]{reader.Name}!{cellAddress} data {data} data type {data.GetType()} does not match data type of column data {reader.GetFieldType(ColumnIndexNumber)}. Supply data with a consistent data type.");
                                errorDetected = true;
                            }
                            else if (data.ToString().Length > rowDataSizeLimit & data.GetType() != reader.GetFieldType(ColumnIndexNumber))
                            {
                                Logger.Error($"[{GetFileName(file)}]{reader.Name}!{cellAddress} is {data.ToString().Length} characters long and exceeds {rowDataSizeLimit} character cell contents limit. Supply valid cell contents. Data type {data.GetType()} does not match data type of column data {reader.GetFieldType(ColumnIndexNumber)}.  Supply data with a consistent data type.");
                                errorDetected = true;
                            }
                            else
                            {
                            }
                        }
                        //else
                        //{
                        //    Logger.Error($"[{GetFileName(file)}]{reader.Name}!{cellAddress} is empty. Supply valid cell data.");
                        //    errorDetected = false;
                        //}

                    }
                }
                reader.Dispose();
                reader.Close();
                return errorDetected;
            }

        }
    }
}
