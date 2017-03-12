Private Sub cancelButton_Click()
Unload Me
End Sub
Private Sub chooseCustomer_Change()
End Sub
Private Sub clearButton_Click()
Call UserForm_Initialize
End Sub
Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
End Sub
Private Sub CommandButton1_Click()
End Sub
Private Sub invoiceNumber_Change()
End Sub

Private Sub populateAndGenerate_Click()

Application.ScreenUpdating = False

'populate the 'Sales Records' sheet and populate the invoice with the unique identifiers
Dim emptyRow As Long

'Make Sheet1 active
Sheet1.Activate

'Determine emptyRow
emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1

'Transfer information
Cells(emptyRow, 1).Value = invoiceDate.Value
Cells(emptyRow, 2).Value = invoiceNumber.Value
Cells(emptyRow, 3).Value = chooseCustomer.Value
Cells(emptyRow, 4).Value = item1.Value
Cells(emptyRow, 5).Value = item1Cost.Value
Cells(emptyRow, 6).Value = item2.Value
Cells(emptyRow, 7).Value = item2Cost.Value
Cells(emptyRow, 8).Value = item3.Value
Cells(emptyRow, 9).Value = item3Cost.Value
Cells(emptyRow, 10).Value = item4.Value
Cells(emptyRow, 11).Value = item4Cost.Value
Cells(emptyRow, 13).Value = vatRate.Value

'Populate the invoice with the unique identifiers
Dim invoice As Worksheet
Set invoice = Worksheets("Invoice")
With invoice
         .Range("K11").Value = Me.invoiceNumber.Value
    .Range("K3").Value = Me.chooseCustomer.Text
End With

'save invoice as PDF
Dim invoiceNo As String
Dim customerName As String
invoiceNo = Me.invoiceNumber.Value
customerName = Me.chooseCustomer.Text
invoice.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:="C:\Users\lucas.salmins\Dropbox\BABA SHEET\Invoices Export\" & invoiceNo & " " & customerName & ".pdf", _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

'Close InvoiceBox1

Application.ScreenUpdating = True

Unload Me
End Sub
Private Sub populateButton_Click()
Application.ScreenUpdating = False
Dim emptyRow As Long

'Make Sheet1 active
Sheet1.Activate

'Determine emptyRow
emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1

'Transfer information
Cells(emptyRow, 1).Value = invoiceDate.Value
Cells(emptyRow, 2).Value = invoiceNumber.Value
Cells(emptyRow, 3).Value = chooseCustomer.Value
Cells(emptyRow, 4).Value = item1.Value
Cells(emptyRow, 5).Value = item1Cost.Value
Cells(emptyRow, 6).Value = item2.Value
Cells(emptyRow, 7).Value = item2Cost.Value
Cells(emptyRow, 8).Value = item3.Value
Cells(emptyRow, 9).Value = item3Cost.Value
Cells(emptyRow, 10).Value = item4.Value
Cells(emptyRow, 11).Value = item4Cost.Value
Cells(emptyRow, 13).Value = vatRate.Value

'Close InvoiceBox1
Application.ScreenUpdating = True
Unload Me
End Sub

Private Sub UserForm_Initialize()

'Set invoiceNumber as next in sequence
Dim NextNum As Long
         NextNum = Application.WorksheetFunction.Max(Sheet1.UsedRange.Columns(2))
    Me.invoiceNumber.Value = NextNum + 1
    Me.invoiceNumber.Enabled = True

'Fill VAT Rate
With vatRate
    .AddItem "0%"
    .AddItem "20%"
End With

'Empty chooseCustomer
chooseCustomer.Clear

'Fill Customer Box with selection from Database
'chooseCustomer.RowSource = "Database!B:B"
Dim rngCustomer As Range
Dim ws As Worksheet
Set ws = Worksheets("Database")
For Each rngCustomer In ws.Range("Customer")
Me.chooseCustomer.AddItem rngCustomer.Value
Next rngCustomer

'Set Focus on invoiceDate
invoiceDate.SetFocus
End Sub
Private Sub UserForm_InitializeTest()

'Set invoiceNumber as next in sequence
Dim NextNum As Long
         NextNum = Application.WorksheetFunction.Max(Sheet1.UsedRange.Columns(2))
    Me.invoiceNumber.Value = NextNum + 1
    Me.invoiceNumber.Enabled = True

'Fill Customer Box with selection from Database

'chooseCustomer.RowSource = "Database!B:B"
Dim rngCustomer As Range
Dim ws As Worksheet
Set ws = Worksheets("Database")
For Each rngCustomer In ws.Range("Customer")
Me.chooseCustomer.AddItem rngCustomer.Value
Next rngCustomer

'prepopulate with test
chooseCustomer.Value = "English Session Orchestra"
item1.Value = "Session 1"
item2.Value = "Porterage"
item3.Value = "Profit Share"
item4.Value = "Transcription"
item1Cost.Value = "1000"
item2Cost.Value = "200"
item3Cost.Value = "300"
item4Cost.Value = "500"
vatRate.Value = "20%"

'Set Focus on invoiceDate
invoiceDate.SetFocus
End Sub
