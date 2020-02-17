Set FSO = CreateObject("Scripting.FileSystemObject")

Set OutputFile = FSO.CreateTextFile("010 - Output File.xls", True)

'Write a simple Excel 2013 xml format header, with three columns
WriteExcelHeader()
'Write a row
WriteExcelRow "1002131", "2016-01-01", "20.00"
'Write the footer
WriteExcelFooter()

Function WriteExcelHeader()

    OutputFile.writeline "<?xml version=""1.0""?>"
    OutputFile.writeline "<?mso-application progid=""Excel.Sheet""?>"
    OutputFile.writeline "<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"""
    OutputFile.writeline " xmlns:o=""urn:schemas-microsoft-com:office:office"""
    OutputFile.writeline " xmlns:x=""urn:schemas-microsoft-com:office:excel"""
    OutputFile.writeline " xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"""
    OutputFile.writeline " xmlns:html=""http://www.w3.org/TR/REC-html40"">"
    OutputFile.writeline " <Worksheet ss:Name=""Sheet1"">"
' Change the '3' to the number of fields you want
    OutputFile.writeline "  <Table ss:ExpandedColumnCount=""3"" x:FullColumns=""1"""
    OutputFile.writeline "  x:FullRows=""1"">"
' Add extra Column lines here for every new field you add 
    OutputFile.writeline "  <Column ss:Width=""90.75""/>"
    OutputFile.writeline "  <Column ss:Width=""56.25""/>"
    OutputFile.writeline "  <Column ss:Width=""54.75""/>"
    OutputFile.writeline "   <Row>"
' Add extra lines here for every new field you add 
    OutputFile.writeline "     <Cell><Data ss:Type=""String"">Account Number</Data></Cell>"
    OutputFile.writeline "     <Cell><Data ss:Type=""String"">Date</Data></Cell>"
    OutputFile.writeline "     <Cell><Data ss:Type=""String"">Posted Amt</Data></Cell>"
    OutputFile.writeline "   </Row>"

End Function

Function WriteExcelRow(AccNum, PostedDte, Amt)

    OutputFile.writeline "   <Row>"
' Add extra lines here for every new field you add 
    OutputFile.writeline "     <Cell><Data ss:Type=""String"">" & AccNum & "</Data></Cell>"
    OutputFile.writeline "     <Cell><Data ss:Type=""String"">" & PostedDte & "</Data></Cell>"
    OutputFile.writeline "     <Cell><Data ss:Type=""String"">" & Amt & "</Data></Cell>"
    OutputFile.writeline "   </Row>"

End Function

Function WriteExcelFooter()

    OutputFile.writeline "  </Table>"
    OutputFile.writeline " </Worksheet>"
    OutputFile.writeline "</Workbook>"

End Function
