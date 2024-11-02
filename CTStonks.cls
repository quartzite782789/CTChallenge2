VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stonks()

Dim u As Integer    'Worksheet count variable
Dim v As Integer    'Iterative Counter for worksheets
Dim w As Worksheet  'Variable to shorten reference to current Worksheet

Dim l As Long   'Initial Data Length
Dim i As Long   'Iterative Counter
Dim j As Long   'Iterative Unique Entry Counter into Unique Entry Data Length
Dim tic As String   'Iterative Ticker Variable
Dim ctic As Double  'Iterative Ticker Counter to determine Unique Values

Dim idat As Variant 'Iterative Date Variable
Dim odat() As Variant   'Array of Quarter Opening Dates (In case different years don't have the same start/end dates)
Dim cdat() As Variant   'Array of Quarter Closing Dates (In case different years don't have the same start/end dates)

Dim opric As Single 'Iterative Opening Price Variable
Dim cpric As Single 'Iterative Closing Price Variable
Dim qc As Single    'Iterative Quartely Change Variable

Dim ginc As Single  'Greatest Quartely Percent Increase
Dim gdec As Single  'Greatest Quartely Percent Decrease
Dim gvol As Double    'Greatest Total Quartely Volume


u = Worksheets.Count        'Find # of Worksheets in Workbook

For v = 1 To u  'Loop to cycle through each worksheet and apply the subroutine to each

    j = 0       'Reset Unique Entry Counter at beginning of subroute run each worksheet
    Set w = Worksheets(v)       'Set w to refer to current worksheet
    With w.Columns("I:P")        'Clear the data entry columns for this routine (Values and formatting)
        .ClearContents
        .FormatConditions.Delete
    End With
    'Setting titles
    w.Range("I1:P1") = Array("Ticker", "Quarterly Change", "Percent Change", "Total Stock Volume", "", "", "Ticker", "Value")
    w.Range("I1:P1").Font.Bold = True
    w.Range("N2:N4") = WorksheetFunction.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))
    w.Range("N2:N4").Font.Bold = True
    
    l = w.Rows.Count      'Assign total worksheet length, for cell clearing
    l = WorksheetFunction.Count(w.Range("C:C"))      'Find raw data length by exploiting COUNT on field that contains all numerical data
    w.Range("B2:B" & l).NumberFormat = "General"      'Set the date field data to general excel timecode format, to prevent comparison issues later

    For i = 1 To l  'Loop to Populate a list of Unique Ticker values and the quaterly opening and closing dates for each
        tic = w.Cells(i + 1, 1).Value     'Set Ticker value for iteration
        idat = w.Cells(i + 1, 2).Value        'Set Date value for iteration
        ctic = WorksheetFunction.CountIf(w.Range("I:I"), tic)        'Check the Unique Ticker so far to see if current ticker is new or duplicate
        If ctic = 0 Then        'If the current ticker vale is new, store it in the print cell range
            j = j + 1       'Increase Unique Counter
            w.Cells(j + 1, 9).Value = tic     'Store the unique ticker in the unique list
            If j = 1 Then       'Setting new dimensions for the first elements of the opening/closing date arrays
                ReDim odat(j)
                ReDim cdat(j)
            Else        'Setting new dimensions of the opening/closing date arrays to accommodate the new Ticker
                ReDim Preserve odat(j - 1)
                ReDim Preserve cdat(j - 1)
            End If
            odat(j - 1) = idat      'Setting initial opening/closing date values for new ticker
            cdat(j - 1) = idat
        ElseIf idat < odat(j - 1) Then      'Compare previously stored earliest opening date against current iterative date
            odat(j - 1) = idat      'If current date earlier, replace previously stored opening date
        ElseIf idat > cdat(j - 1) Then      'Compare previously stored latetest closing date against current iterative date
            cdat(j - 1) = idat      'If current date later, replace previously stored closing date
        End If
    Next i

    For i = 1 To j  'Loop to run through newly made unique ticker list and populate the rest of the data
        tic = w.Cells(i + 1, 9).Value     'Set Ticker value for current iteration in the unique list
        opric = WorksheetFunction.CountIfs(w.Range("A:A"), tic, w.Range("B:B"), odat(i - 1))      'Check for any duplicate data for opening/closing dates for a given Ticker
        cpric = WorksheetFunction.CountIfs(w.Range("A:A"), tic, w.Range("B:B"), cdat(i - 1))
        
        If opric > 1 Or cpric > 1 Then  'Completely [[Optional]] If statement to make sure users are populating data in the sheet at least SOMEWHAT correctly
            MsgBox ("Duplicate date report data detected! Please check data and try again!")        'Inform user of error in the data, erase the latest data and formatting in the print area, and exit the sub routine
            l = w.Rows.Count
            With w.Columns("I:P")
                .ClearContents
                .FormatConditions.Delete
            End With
            Exit Sub
        End If  'End completely [[Optional]] If statement for the sake of my own paranoia
        
        opric = WorksheetFunction.SumIfs(w.Range("C:C"), w.Range("A:A"), tic, w.Range("B:B"), odat(i - 1))     'Set the opening price for given Ticker
        cpric = WorksheetFunction.SumIfs(w.Range("F:F"), w.Range("A:A"), tic, w.Range("B:B"), cdat(i - 1))     'Set the closing price for given Ticker
        qc = cpric - opric      'Calculate and set Quarterly change value for given Ticker
        With w.Cells(i + 1, 10)     'Print Quarterly Change with up to 2 decimal places
            .Value = qc
            .NumberFormat = "$0.00"
            With .FormatConditions      'Add conditional formatting to color Quarterly Change Green when an increase and Red when a decrease
                .Add(xlCellValue, xlGreater, 0).Interior.ColorIndex = 4
                .Add(xlCellValue, xlLess, 0).Interior.ColorIndex = 3
            End With
        End With
        With w.Cells(i + 1, 11)       'Calculate and print Percent Change as a 2 decimal percent
            .Value = qc / opric
            .NumberFormat = "0.00%"
            With .FormatConditions      'Add conditional formatting to color Quarterly Change Green when an increase and Red when a decrease
                .Add(xlCellValue, xlGreater, 0).Interior.ColorIndex = 4
                .Add(xlCellValue, xlLess, 0).Interior.ColorIndex = 3
            End With
        End With
        w.Cells(i + 1, 12).Value = WorksheetFunction.SumIf(w.Range("A:A"), tic, w.Range("G:G"))     'Sum up the total volume for a given Ticker
    Next i

    ginc = WorksheetFunction.Max(w.Range("K2:K" & (j + 1)))       'Find Greatest % Increase
    gdec = WorksheetFunction.Min(w.Range("K2:K" & (j + 1)))       'Find Greatest % Decrease
    gvol = WorksheetFunction.Max(w.Range("L2:L" & (j + 1)))       'Find Greatest Total Volume
    w.Range("O2").Value = WorksheetFunction.XLookup(ginc, w.Range("K:K"), w.Range("I:I"), "Whoops")     'Print the Ticker for the Greatest % Increase
    With w.Range("P2")        'Print Greatest % Increase as a 2 decimal percent
        .Value = ginc
        .NumberFormat = "0.00%"
    End With
    w.Range("O3").Value = WorksheetFunction.XLookup(gdec, w.Range("K:K"), w.Range("I:I"), "Whoops")     'Print the Ticker for the Greatest % Decrease
    With w.Range("P3")        'Print Greatest % Decrease as a 2 decimal percent
        .Value = gdec
        .NumberFormat = "0.00%"
    End With
    w.Range("O4").Value = WorksheetFunction.XLookup(gvol, w.Range("L:L"), w.Range("I:I"), "Whoops")     'Print the Ticker for the Greatest Total Volume
    w.Range("P4").Value = gvol        'Print Greatest Total Volume

    w.Range("B2:B" & l).NumberFormat = "m/d/yyyy"     'Reset formatting for the date column of the origninal data to short date format
    w.Columns("I:L").AutoFit
    w.Columns("N:P").AutoFit

Next v
End Sub
