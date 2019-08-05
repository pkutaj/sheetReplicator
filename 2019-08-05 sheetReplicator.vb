' # PURPOSE: LOOP through the folder full of folders and copy an active sheet.
' 3 DISCLAIMER: This is a quick and dirty VBA, no proper abstraction mechanisms applied ! 
'

Sub sheetReplicator()
    '0 DECLARE MASTER DATA
    Dim mainWB As Workbook: Set mainWB = Workbooks("sheetReplicator.xlsm")
    Dim masterDataSheet As Worksheet: Set masterDataSheet = mainWB.Sheets("COMPILATION_MASTER_DATA")
    Dim definitionSheet As Worksheet: Set definitionSheet = mainWB.Sheets("COMPILATION_DEFINITION")
    Dim definitionSheetHeader As Range: Set definitionSheetHeader = definitionSheet.Rows("1")
    
    '1. GET file-names from a current folder
    '1.1.1 DECLARE dynamic array for filenames
    Dim allFiles() As String
    ReDim allFiles(0 To 0)

    '1.2 Make sure you are in a current folder and assign the first file
    Dim fileName As String
    ChDir ActiveWorkbook.Path
    Debug.Print CurDir()
    fileName = Dir("")
    Debug.Print fileName

    '1.3 Loop through the files in the current folder to fill an array
    Do While fileName <> ""
        allFiles(UBound(allFiles)) = fileName
        fileName = Dir()
        If fileName = "" Then Exit Do
        ReDim Preserve allFiles(0 To UBound(allFiles) + 1)
    Loop

    '2. Loop through the array and move the data in
    For Each file In allFiles
        If file <> "sheetReplicator.xlsm" Then   '<-- enter the sheet name here
            Dim singleWorkBook As Excel.Workbook
            Set singleWorkBook = Workbooks.Open(file)
            singleWorkBook.UpdateLinks = xlUpdateLinksAlways
            Dim compilationDefinitionSheet As Worksheet: Set compilationDefinitionSheet = singleWorkBook.Sheets.Add(Before:=Sheets(1))
            Dim destinationHeader As Range: Set destinationHeader = compilationDefinitionSheet.Rows("1")
                     
            '2.1 copy the master data sheet
            masterDataSheet.Copy Before:=singleWorkBook.Sheets(2)
            
            '2.2 copy the headers range
            compilationDefinitionSheet.Name = "COMPILATION_DEFINITION"
            definitionSheetHeader.Copy
            destinationHeader.PasteSpecial Paste:=xlPasteFormats
            destinationHeader.PasteSpecial Paste:=xlPasteValues
            
            '2.3 save and close without asking
            singleWorkBook.Close True
            
        End If
    Next file
End Sub
