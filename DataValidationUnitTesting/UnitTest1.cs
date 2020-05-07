using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using ExcelDataReader;
using System.Data;
using DataReaderValidatorAndUploaderApp;
using DataReaderValidatorAndUploader;
namespace UnitTestProject1
{
    [TestClass]
    public class BasicValidationTest
    {
        BasicValidation basicValidation;
        string filePath1;
        string filePath2;
        string filePath3;
        string filePath4;
        string filePath5;
        string filePath6;
        string filePath7;
        string filePath8;
        string filePath9;

        [TestInitialize]
        public void TestInit()
        {
            filePath1 = "C:\\Users\\Jojo\\source\\repos\\DataReaderValidatorAndUploaderApp\\Book3.xlsx";
            filePath2 = "C:\\Users\\Jojo\\source\\repos\\DataReaderValidatorAndUploaderApp\\Book11.xlsx";
            filePath3 = "C:\\Users\\Jojo\\source\\repos\\DataReaderValidatorAndUploaderApp\\Book9.xlsx";
            filePath4 = "C:\\Users\\Jojo\\source\\repos\\DataReaderValidatorAndUploaderApp\\Book4.xlsx";
            filePath5 = "C:\\Users\\Jojo\\source\\repos\\DataReaderValidatorAndUploaderApp\\Book7.xlsx";
            filePath6 = "C:\\Users\\Jojo\\source\\repos\\DataReaderValidatorAndUploaderApp\\Book16.xlsx";
            filePath7 = "C:\\Users\\Jojo\\source\\repos\\DataReaderValidatorAndUploaderApp\\Book15.xlsx";
            filePath8 = "C:\\Users\\Jojo\\source\\repos\\DataReaderValidatorAndUploaderApp\\Book18.xlsx";
            filePath9 = "C:\\Users\\Jojo\\source\\repos\\DataReaderValidatorAndUploaderApp\\LongestFileEvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvver";
        }

        [TestMethod]
        public void InvalidColumnNames_ColumnNameIsNotEmpty_ReturnsFalse()
        {
            //Arrange
            basicValidation = new BasicValidation(filePath1, 128, 32767, 128, 128);
            //Act
            object ColsNamesEmpty = basicValidation.InvalidColumnNames();
            //Assert
            Assert.AreEqual(false, ColsNamesEmpty);
        }


        [TestMethod]
        public void InvalidColumnNames_ColumnNameIsEmpty_ReturnsTrue()
        {
            //Arrange
            basicValidation = new BasicValidation(filePath2, 128, 32767, 128, 128);
            //Act
            object ColumnNamesEmpty = basicValidation.InvalidColumnNames();
            //Assert
            Assert.AreEqual(true, ColumnNamesEmpty);
        }

        [TestMethod]
        public void InvalidColumnNames_ColumnNameIsTooLong_ReturnsTrue()
        {
            //Arrange
            basicValidation = new BasicValidation(filePath3, 128, 32767, 128, 128);
            //Act
            object ColumnNamesTooLong = basicValidation.InvalidColumnNames();
            //Assert
            Assert.AreEqual(true, ColumnNamesTooLong);
        }

        [TestMethod]
        public void InvalidColumnNames_FileIsEmpty_ReturnsTrue()
        {
            //Arrange
            basicValidation = new BasicValidation(filePath4, 128, 32767, 128, 128);
            //Act
            object FileEmpty = basicValidation.InvalidColumnNames();
            //Assert
            Assert.AreEqual(true, FileEmpty);
        }

        [TestMethod]
        public void InvalidColumnNames_CellDataHasValidSizeAndDataType_ReturnsFalse()
        {
            //Arrange
            basicValidation = new BasicValidation(filePath1, 128, 32767, 128, 128);
            //Act
            object CellHasInValidSizeDatatype = basicValidation.InvalidCellData();
            //Assert
            Assert.AreEqual(false, CellHasInValidSizeDatatype);
        }

        [TestMethod]
        public void InvalidColumnNames_CellDataHasInValidSizeDataType_ReturnsTrue()
        {
            //Arrange
            basicValidation = new BasicValidation(filePath5, 128, 32767, 128, 128);
            //Act
            object CellHasInValidSizeAndDatatype = basicValidation.InvalidCellData();
            //Assert
            Assert.AreEqual(true, CellHasInValidSizeAndDatatype);
        }

        [TestMethod]
        public void InvalidColumnNames_CellIsEmpty_ReturnsTrue()
        {
            //Arrange
            basicValidation = new BasicValidation(filePath6, 128, 32767, 128, 128);
            //Act
            object CellIsEmpty = basicValidation.InvalidCellData();
            //Assert
            Assert.AreEqual(true, CellIsEmpty);
        }

        [TestMethod]
        public void InvalidColumnNames_CellDataHasWrongDataType_ReturnsTrue()
        {
            //Arrange
            basicValidation = new BasicValidation(filePath5, 128, 32767, 128, 128);
            //Act
            object CellHasValidSizeAndWrongDataType = basicValidation.InvalidCellData();
            //Assert
            Assert.AreEqual(true, CellHasValidSizeAndWrongDataType);
        }

        [TestMethod]
        public void InvalidColumnNames_CellDataInValidSize_ReturnsTrue()
        {
            //Arrange
            basicValidation = new BasicValidation(filePath8, 128, 32767, 128, 128);
            //Act
            object CellHasInValidSizeButRightDataType = basicValidation.InvalidCellData();
            //Assert
            Assert.AreEqual(false, CellHasInValidSizeButRightDataType);
        }

        [TestMethod]
        public void InvalidColumnNames_FileNameIsTooBig_ReturnsTrue()
        {
            //Arrange
            basicValidation = new BasicValidation(filePath9, 128, 32767, 128, 128);
            //Act
            object fileNameIsTooBig = basicValidation.FileNameValidation();
            //Assert
            Assert.AreEqual(true, fileNameIsTooBig);
        }


        //[TestMethod]
        //public void TestIfColumnNamesEmpty()
        //{
        //    //Arrange
        //    string filePath1 = "C:\\Temp\\source\\Book3.xlsx";
        //    string filePath2 = "C:\\Temp\\source\\Book11.xlsx";
        //    string filePath3 = "C:\\Temp\\source\\Book9.xlsx";
        //    string filePath4 = "C:\\Temp\\source\\Book4.xlsx";
        //    //Act
        //    BasicValidation basicValidation = new BasicValidation();
        //    object ColumnHaveNames = basicValidation.ColumnNamesValidation(filePath1, 0);
        //    object ColumnNamesEmpty = basicValidation.ColumnNamesValidation(filePath2, 1);
        //    object ColumnNamesTooLong = basicValidation.ColumnNamesValidation(filePath3, 2);
        //    object FileEmpty = basicValidation.ColumnNamesValidation(filePath4,0);
        //    //Assert
        //    Assert.AreEqual("Columns names have valid size", ColumnHaveNames);
        //   // Assert.AreEqual("column name is empty", ColumnNamesEmpty);
        //   // Assert.AreEqual("column name is too long", ColumnNamesTooLong);
        //   // Assert.AreEqual("file is empty", FileEmpty);
        //}

        //[TestMethod]
        //public void TestCellDataValidation()
        //{
        //    //Arrange
        //    string filePath1 = "C:\\Temp\\source\\Book3.xlsx";
        //    string filePath2 = "C:\\Temp\\source\\Book7.xlsx";
        //    string filePath3 = "C:\\Temp\\source\\Book16.xlsx";
        //    string filePath4 = "C:\\Temp\\source\\Book15.xlsx";
        //    //Act
        //    BasicValidation basicValidation = new BasicValidation();
        //    object CellHasValidSizeAndDatatype = basicValidation.CellDataValidation(filePath1, 0, 0);
        //    object CellHasInValidSizeAndDatatype = basicValidation.CellDataValidation(filePath2, 2, 0);
        //    object CellIsEmpty = basicValidation.CellDataValidation(filePath3, 0, 0);
        //    object CellHasValidSizeAndWrongDataType = basicValidation.CellDataValidation(filePath4, 1, 1);
        //    object CellHasInValidSizeButRightDataType = basicValidation.CellDataValidation(filePath1, 0, 2);
        //    //Assert
        //    Assert.AreEqual("data in a cell has a valid size and datatype", CellHasValidSizeAndDatatype);
        //    Assert.AreEqual("data in a cell has invalid size and dataType", CellHasInValidSizeAndDatatype);
        //    Assert.AreEqual("Cell has no data, it is empty", CellIsEmpty);
        //    Assert.AreEqual("data in a cell has a valid size but wrong datatype", CellHasValidSizeAndWrongDataType);
        //    Assert.AreEqual("data in a cell has invalid size but right dataType", CellHasInValidSizeButRightDataType);
        //}
    }
}
