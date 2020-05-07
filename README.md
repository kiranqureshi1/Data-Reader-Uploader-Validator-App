# Data-Reader-Uploader-Validator-App


Technology Used: C# ExcelDataReader SMO (Server Management Objects)

Basic Functionality: App goes to folder location and open the folder and the loops through the Excel files Opening each Excel file one by one, reading file name and then looping through sheets, reading each sheet name and columns names, rows, their dataTypes. In the next step the app checks for basic validation. first does the sheet name exceeds 128 size limit? Then is the sheet empty? It then loops through all the column names to see: Are there any null column names? does each column name exceeds 128 characters limit? Are there any duplicated column names?

Then it loops through rows Checking: does the row data exceeds – characters limit? doesnt the dataType of row data match column dataType?

Above all column and rows statements are boolean statements. If any of the statement for columns/rows or both are true, the app fails validation for that particular file and gives an error message on the console such and such file has problem with such and such row column (it displays error for all the rows and columns that has problem for the ease of the user to fix the problem) and finally displaying a message that this particular file failed validation and will jump to the next unprocessed file.

If in a loop any of the files pass validation. It will then jump to the next step for that file. This will include: A new table will be created with sheet name. Then columns names will be created with the data in the first row in excel file. Then rows will be created by sqlBulkcopy. Excel dataTypes will be converted to SQL dataTypes i.e string will convert into NVarchar, int will convert into double e.t.c) After the whole table is created for that particular file, it will then upload that file to the database. The app will check for the database. If the database with the folder name already exists. It will ask the user in console: would you like ti use existing database (displaying database names made for this app before), if the user types yes, File will be uploaded to the existing database by displaying the existing database names and asking the user: Type the name of the database form the list above to upload the file to. Whatever name user types in the list above it will create table for that file in that chosen database. ( I am going to try to add some code for, if typed database name doesn’t match any of the existing database names then it will prompt the user by asking check the spelling and type the name from existing database names in the list above) If the user types no ( saying that user doesn’t want to use existing databases) then the console asks a question: what would you like to call your database? If the user types a name of the database then the console will prompt user displaying message that such and such database already exists, please type a valid name. once the user types a valid database name, database is created with that particular name that user entered and file is then uploaded to that database ( by creating column names, table name, rows, converting datatypes to sql datatypes)

Then code jumps to the next unprocessed file in the loop to do the same thing (check for validation and if validation passes then jumps to uploading file to the database)

Unit testing is done for file validation.

All the basic code is in class library (DataReaderValidatorAndUploader) which is accessed in the console app (DataReaderValidatorAndUploaderApp).

How to make a class Library:
Select clasLibrary in the new project. 
Once made then add console app which will use code written in your class library by adding it to the references of your console app.

