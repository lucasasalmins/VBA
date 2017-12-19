Sub combineAllSheets()

Application.ScreenUpdating = False

    Dim metersPath As String, metersName As String
    Dim sourceMeterDataBook As Workbook
    Dim combinationBook As Workbook
    Dim sourceMeterDataBookName As String
    
    'Loop through the folder, open each and copy the data into the main sheet
    Set combinationBook = ThisWorkbook
    metersPath = "C:\Users\lucas.salmins\Dropbox\PP Comms\PDHU\Originals\"
    sourceMeterDataBookName = Dir(metersPath & "*.xls")
    
    Do While sourceMeterDataBookName <> ""
        'diable macros on opening workbook
        Application.EnableEvents = False
        Set sourceMeterDataBook = Workbooks.Open(metersPath & sourceMeterDataBookName)
        'unlock the worksheet
        ActiveSheet.Unprotect ("pdhu")

        sourceMeterDataBook.Sheets("Energy Centre").Range("A2:J5").Copy
        combinationBook.Sheets("MasterComb").Range("A1000000").End(xlUp).Offset(1, 0).PasteSpecial Paste:=xlPasteValues
        
        sourceMeterDataBook.Sheets("Abbots Manor").Range("A2:J25").Copy
        combinationBook.Sheets("MasterComb").Range("A1000000").End(xlUp).Offset(1, 0).PasteSpecial Paste:=xlPasteValues
        
        sourceMeterDataBook.Sheets("Churchill Gardens").Range("A2:J54").Copy
        combinationBook.Sheets("MasterComb").Range("A1000000").End(xlUp).Offset(1, 0).PasteSpecial Paste:=xlPasteValues
        
        sourceMeterDataBook.Sheets("Lillington & Longmoore").Range("A2:J22").Copy
        combinationBook.Sheets("MasterComb").Range("A1000000").End(xlUp).Offset(1, 0).PasteSpecial Paste:=xlPasteValues
            
        sourceMeterDataBook.Sheets("Miscellaneous Commercials").Range("A2:J5").Copy
        combinationBook.Sheets("MasterComb").Range("A1000000").End(xlUp).Offset(1, 0).PasteSpecial Paste:=xlPasteValues
     
        'paste the filename to the next blank row
        combinationBook.Sheets("MasterComb").Range("A1000000").End(xlUp).Offset(-104, 10) = sourceMeterDataBook.Name
        Columns("K:K").Select
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.FormulaR1C1 = "=R[-1]C"
          
    Application.CutCopyMode = False

    sourceMeterDataBook.Close SaveChanges:=False
    
    Application.EnableEvents = True
    
    sourceMeterDataBookName = Dir
    
    Loop
    Columns("K:K").Select
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.FormulaR1C1 = "=R[-1]C"
    
    combinationBook.Save
    
    Application.ScreenUpdating = True

End Sub

