Set FSO = CreateObject("Scripting.FileSystemObject")

' Open the file called "000 - First File.txt"
Set InputFile = FSO.OpenTextFile("000 - First File.txt", 1, False)
' Read the Header Line
LineIn = InputFile.ReadLine
' Read each line in the file.
While Not InputFile.AtEndOfStream
' Split the comma-separated line and store the fields in an array.
    LineArray = Split(LineIn, ',')
' Get the second tab-separated field in the line and store in 'ID'.
    ID = LineArray(2)
WEnd

