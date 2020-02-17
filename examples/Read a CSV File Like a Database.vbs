SetLocale("en-ie")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set FieldLengths = CreateObject("Scripting.Dictionary")

' Connect to the database using ODBC
Set DatabaseConnection = CreateObject("ADODB.Connection")
DatabaseConnection.CommandTimeout = 0

' The path to the CSV file
PathToCSVFile = "C:\"

' 'Your Database' is the DSN name in ODBC
DatabaseConnection.Open _
       "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "DSN=Your Database;" & _ 
       "Data Source=" & PathToCSVFile & ";" & _
       "Extended Properties=""text;HDR=YES;FMT=Delimited"""


' Our Output file
OutputFileName = "010 - Query Result.txt"
Set OutputFile = FSO.CreateTextFile(OutputFileName, True)

Set DBConn = CreateObject("ADODB.Recordset")

CSVFile = "000 - Input File.txt"
DBConn.open "SELECT * FROM " & CSVFile, DatabaseConnection, 3, 3, &H0001

' Write the detail lines
Do While Not DBConn.EOF
    WScript.Echo(DBConn.Fields.Item["Name"])
    WScript.Echo(DBConn.Fields.Item["ID"])
    DBConn.MoveNext 
Loop

