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
'> Variables <
Dim u As Integer    'Worksheet count variable
Dim v As Integer    'Iterative Counter for worksheets

Dim w As Worksheet  'Variable to shorten reference to current Worksheet
Dim l As Long   'Initial Data Length
Dim gpi As Single   'Greatest % Increase
Dim gpd As Single   'Greatest % Decrease
Dim gtv As Double     'Greatest Total Volume
Dim i As Long   'Iterative Counter

Dim itic As String  'Iterative Ticker Variable
Dim idat As Long 'Iterative Date Variable
Dim iopr As Currency    'Iterative Opening Price Variable
Dim icpr As Currency    'Iterative Closing Price Variable
Dim ivol As Double    'Iterative Volume Variable

Dim j As Integer   'Iterative index variable for intermediate storage arrays
Dim k As Integer   'Iterative index variable for output arrays
Dim o As Integer    'Index variable pertaining to the earliest opening price
Dim odat As Long 'Iterative opening date variable
Dim c As Integer    'Index variable pertaining to the latest closing price
Dim cdat As Long 'Iterative closing date variable
Dim tica() As Variant   'Intermediate Storage Array for Tickers
Dim opra() As Currency   'Intermediate Storage Array for Opening Prices
Dim cpra() As Currency   'Intermediate Storage Array for Closing Prices
Dim vol As Double     'Iterative sum for total volume

Dim rtica() As Variant  'Output array for tickers
Dim rqca() As Currency   'Output array for quarterly changes
Dim rpca() As Currency   'Output array for percent changes
Dim rvola() As Double  'Output array for total volume

Dim gpitic As String    'Ticker for greatest percent increase
Dim gpdtic As String    'Ticker for greatest percent decrease
Dim gtvtic As String    'Ticker for greatest total volume


'> Find # of Worksheets in Workbook <
u = Worksheets.Count

'> Cycle through each worksheet and apply the subroutine to each <
For v = 1 To u
'>>>> Set up the worksheet <
'   >[Initialize Variables]
    Set w = Worksheets(v)       'Set w to refer to current worksheet
    l = w.Cells(w.Rows.Count, 1).End(xlUp).Row      'Find last row with data
    gpi = 0
    gpd = 0
    gtv = 0
    
'   >[Prep date data and data print area]
    w.Range("B2:B" & l).NumberFormat = "General"        'Set the date field data to general excel timecode format, to prevent comparison issues later
    With w.Columns("I:Q")       'Clear the data print area
        .ClearContents
        .ClearFormats
        .FormatConditions.Delete
    End With
    
'   >[Format the data headers]
    w.Range("I1:Q1") = Array("Ticker", "Quarterly Change", "Percent Change", "Total Stock Volume", "", "", "", "Ticker", "Value")
    w.Range("I1:Q1").Font.Bold = True
    w.Range("O2:O4") = WorksheetFunction.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))
    w.Range("O2:O4").Font.Bold = True

'>>>> Data retrieval loop <
    For i = 2 To l
'       >[Set values for current row]
        itic = w.Cells(i, 1).Value      'Ticker for current row
        idat = w.Cells(i, 2).Value      'Date for current row
        iopr = w.Cells(i, 3).Value      'Opening Price for current row
        icpr = w.Cells(i, 6).Value      'Closing Price for current row
        ivol = w.Cells(i, 7).Value      'Volume for current row
        
'>>>>>>>> If structure to store and analyze data <
'       >[Set up variables and arrays upon first data entry]
        If i = 2 Then
            j = 0       'Zero element variables and initialize comparison dates as the date in the first row
            k = 0
            o = 0
            odat = idat
            c = 0
            cdat = idat
            
            ReDim tica(1)       'Reinitialize intermediate storage arrays and set first element value
            ReDim opra(1)
            ReDim cpra(1)
            tica(j) = itic
            opra(j) = iopr
            cpra(j) = icpr
            vol = ivol      'Start summing volumes with the volume in first row
            
'       >[Store data if under the same ticker]
        ElseIf itic = tica(j) And i < l Then
            j = j + 1       'Increase element counter
            ReDim Preserve tica(j)      'Resize intermediate storage arrays and set next element value
            ReDim Preserve opra(j)
            ReDim Preserve cpra(j)
            tica(j) = itic
            opra(j) = iopr
            cpra(j) = icpr
            vol = vol + ivol        'Continue summing volumes
            
            If odat > idat Then     'If structure to compare opening/closing dates with current row date, updating the comparison values and storing the element index
                odat = idat
                o = j
            ElseIf c < idat Then
                cdat = idat
                c = j
            End If
            
'       >[Finalize previous Ticker data in output arrays and set up for next ticker]
        Else
            If k = 0 Then       'Reinitialize output arrays on first element
                ReDim rtica(1)
                ReDim rqca(1)
                ReDim rpca(1)
                ReDim rvola(1)
            Else        'Resize output arrays to append new elements
                ReDim Preserve rtica(k)
                ReDim Preserve rqca(k)
                ReDim Preserve rpca(k)
                ReDim Preserve rvola(k)
            End If
            
            rtica(k) = tica(j)      'Set output array values
            rqca(k) = cpra(c) - opra(o)
            rpca(k) = rqca(k) / opra(o)
            rvola(k) = vol
            
            If gpi = 0 Then     'Initialize greatest increase/decrease variables for iterative comparison
                gpi = rpca(k)
                gpitic = rtica(k)
                gpd = rpca(k)
                gpdtic = rtica(k)
            ElseIf gpi < rpca(k) Then       'If larger greatest percent increase found, store
                gpi = rpca(k)
                gpitic = rtica(k)
            ElseIf gpd > rpca(k) Then       'If larger greatest percent decrease found, store
                gpd = rpca(k)
                gpdtic = rtica(k)
            End If
            
            If gtv = 0 Then     'Initialize greatest total volume variables for iterative comparison
                gtv = rvola(k)
                gtvtic = rtica(k)
            ElseIf gtv < rvola(k) Then      'If larger total volume found, store
                gtv = rvola(k)
                gtvtic = rtica(k)
            End If
            
            k = k + 1       'Advance output array index and reset intermediate indexes and variables
            j = 0
            o = 0
            odat = idat
            c = 0
            cdat = odat
            
            ReDim tica(1)       'Reset intermediate arrays and set first element
            ReDim opra(1)
            ReDim cpra(1)
            tica(j) = itic
            opra(j) = iopr
            cpra(j) = icpr
            vol = ivol      'Reset volume sum variable to start with the current row
        End If
    Next i
    
'>>>> Print output data and format
    w.Range("I2:I" & k + 1).Value = WorksheetFunction.Transpose(rtica)
    With w.Range("J2:J" & k + 1)
        .Value = WorksheetFunction.Transpose(rqca)
        .NumberFormat = "$0.00"
        With .FormatConditions      'Add conditional formatting to color Quarterly Change Green when an increase and Red when a decrease
            .Add(xlCellValue, xlGreater, 0).Interior.ColorIndex = 4
            .Add(xlCellValue, xlLess, 0).Interior.ColorIndex = 3
        End With
    End With
    With w.Range("K2:K" & k + 1)
        .Value = WorksheetFunction.Transpose(rpca)
        .NumberFormat = "0.00%"
        With .FormatConditions      'Add conditional formatting to color Quarterly Change Green when an increase and Red when a decrease
            .Add(xlCellValue, xlGreater, 0).Interior.ColorIndex = 4
            .Add(xlCellValue, xlLess, 0).Interior.ColorIndex = 3
        End With
    End With
    With w.Range("L2:L" & k + 1)
        .Value = WorksheetFunction.Transpose(rvola)
        .NumberFormat = "General"
    End With
    w.Range("P2").Value = gpitic
    With w.Range("Q2")
        .Value = gpi
        .NumberFormat = "0.00%"
    End With
    w.Range("P3").Value = gpdtic
    With w.Range("Q3")
        .Value = gpd
        .NumberFormat = "0.00%"
    End With
    w.Range("P4").Value = gtvtic
    With w.Range("Q4")
        .Value = gtv
        .NumberFormat = "General"
    End With
    
    w.Columns("I:L").AutoFit        'Autofit output columns and size the intermediate columns
    w.Columns("M:N").ColumnWidth = 6
    w.Columns("O:Q").AutoFit
    w.Range("B2:B" & l).NumberFormat = "m/d/yyyy"     'Reset formatting for the date column of the origninal data to short date format
Next v
End Sub
