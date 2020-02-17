Set FSO = CreateObject("Scripting.FileSystemObject")
Set Dictionary = CreateObject("Scripting.Dictionary")
  

' Read in all the IDs in the file.
Set InputFile = FSO.OpenTextFile("000 - First File.txt", 1, False)
While Not InputFile.AtEndOfStream

    LineIn = InputFile.ReadLine
    
' If line doesn't exist in the dictionary already, then add it.
    If Not Dictionary.Exists(LineIn) then
        Dictionary.add LineIn, ""
    End If
    
WEnd

Set InputFile = FSO.OpenTextFile("010 - Second File.txt", 1, False)
Set OutputFile = FSO.CreateTextFile("030 - Ones in Second File But Not First File.txt", True)

' Read through the File
While Not InputFile.AtEndOfStream

    LineIn = InputFile.ReadLine

' If it doesn't exist in the dictionary then write it out.
    If Not Dictionary.Exists(LineIn) then
        OutputFile.writeline LineIn
    End If

WEnd

