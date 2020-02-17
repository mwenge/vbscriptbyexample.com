Set FSO = CreateObject("Scripting.FileSystemObject")

' Open the file called "000 - First File.txt"
Set InputFile = FSO.OpenTextFile("000 - First File.txt", 1, False)
' Read each line inthe file.
While Not InputFile.AtEndOfStream
    LineIn = InputFile.ReadLine
' Split the tab-separated line and store the fields in an array.
    LineArray = Split(LineIn, chr(9))
' Get the second tab-separated field in the line and store in 'ID'.
    ID = LineArray(2)
WEnd

