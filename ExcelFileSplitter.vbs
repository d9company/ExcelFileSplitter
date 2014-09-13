
x=msgbox("Please select the Import Excel File." ,64, "Excel File Splitter")
Set wShell=CreateObject("WScript.Shell")
Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
sFileSelected = oExec.StdOut.ReadLine
'wscript.echo sFileSelected

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False
Set objWorkbook = objExcel.Workbooks.Open(sFileSelected)
       Set xlmodule = objworkbook.VBProject.VBComponents.Add(1) 
       strCode = _
        "Sub SplitFile()" & vbCr & _
		"    Dim rowsperfile As String" & vbCr & _
        "    rowsperfile = InputBox(""How many rows do you want to add per one file?"", ""Excel File Splitter"")" & vbCr & _
        "    Dim lLoop As Long, lCopy As Long" & vbCr & _
        "    Dim LastRow As Long" & vbCr & _
        "    Dim wbNew As Workbook" & vbCr & _
        "    With ThisWorkbook.Sheets(1)" & vbCr & _
        "    LastRow = .Range(""A"" & Rows.Count).End(xlUp).Row" & vbCr & _
        "    For lLoop = 1 To LastRow Step rowsperfile" & vbCr & _
        "    lCopy = lCopy + 1" & vbCr & _
        "    Set wbNew = Workbooks.Add" & vbCr & _
        "    .Range(.Cells(lLoop, 1), .Cells(lLoop + rowsperfile, .Columns.Count)).EntireRow.Copy Destination:=wbNew.Sheets(1).Range(""A1"")" & vbCr & _
        "    wbNew.Close SaveChanges:=True, Filename:=ThisWorkbook.Path & ""\Chunk"" & lCopy & ""Rows"" & lLoop & ""-"" & lLoop + rowsperfile" & vbCr & _
        "    Next lLoop" & vbCr & _
        "    End With" & vbCr & _
        "End Sub"
       xlmodule.CodeModule.AddFromString strCode

objExcel.Application.Run "breakuptestfile.xlsx!SplitFile"  
objWorkbook.Save
objExcel.Quit
x=msgbox("Files are successfully splitted. Please check the file location!" ,64, "Excel File Splitter")