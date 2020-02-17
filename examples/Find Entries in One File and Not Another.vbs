Set FSO = CreateObject("Scripting.FileSystemObject")
Set Dictionary = CreateObject("Scripting.Dictionary")
  

' Read in all the IDs in the file.
Set InputFile = FSO.OpenTextFile("000 - First File.txt", 1, False)
While Not InputFile.AtEndOfStream

    LineIn = InputFile.ReadLine
' Split the tab-separated line and store the second field in ID.
    LineArray = Split(LineIn, chr(9))
    ID = LineArray(2)
    
' If ID doesn't exist in the dictionary already, then add it.
    If Not Dictionary.Exists(ID) then
        Dictionary.add ID, ""
    End If
    
WEnd

Set InputFile = FSO.OpenTextFile("010 - Second File.txt", 1, False)
Set OutputFile = FSO.CreateTextFile("030 - Ones in Second File But Not First File.txt", True)

' Read through the File
While Not InputFile.AtEndOfStream

    LineIn = InputFile.ReadLine
    LineArray = Split(LineIn, chr(9))
    ID = LineArray(2)

' If it doesn't exist in the dictionary then write it out.
    If Not Dictionary.Exists(ID) then
        OutputFile.writeline LineIn
    End If

WEnd

