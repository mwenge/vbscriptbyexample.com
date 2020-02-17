SetLocale("en-ie")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set FieldLengths = CreateObject("Scripting.Dictionary")

' Get the username and pasword for the database from the user
Username = inputbox("Enter Username:")
Password = inputbox("Enter Password:")

' Connect to the database using ODBC
Set DatabaseConnection = CreateObject("ADODB.Connection")
DatabaseConnection.CommandTimeout = 0
' 'Your Database' is the DSN name in ODBC
DatabaseConnection.Open "DSN=Your Database;" & _ 
       "Uid="&Username&";" & _ 
       "Pwd="&Password&";"

' Our input file
SQLFileName = "000 - Query.sql"
OutputFileName = "010 - Query Result.txt"
Set OutputFile = FSO.CreateTextFile(OutputFileName, True)

Set oS = FSO.OpenTextFile(SQLFileName, 1, False)
sSQL = os.ReadAll

Set DBConn = CreateObject("ADODB.Recordset")
DBConn.open sSQL, DatabaseConnection

If Not (DBConn.bof and DBConn.eof) Then

    ' Write the header line
    DBConn.MoveFirst
    For Each It In DBConn.Fields
        FieldName = Trim(It.Name)
        If LineOut = "" Then
            LineOut = FieldName
        Else
            LineOut = LineOut &chr(9)& FieldName
        End If
    next
    OutputFile.WriteLine LineOut

    ' Write the detail lines
    DBConn.MoveNext 
    Do While Not DBConn.EOF
        LineOut = ""
        For Each It In DBConn.Fields
            FieldValue = Trim(It.Value)
            If LineOut = "" Then
                LineOut = FieldValue
            Else
                LineOut = LineOut &chr(9)& FieldValue
            End If
        next
        OutputFile.WriteLine LineOut
        DBConn.MoveNext 
    Loop
End If

