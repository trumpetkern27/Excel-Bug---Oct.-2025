Option Explicit

'WinAPI calls
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, _
    ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
        (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
    
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
        
Const GWL_STYLE As Long = -16

'populate Sheet1 with graphs
Public Sub populateGraphs()
    Dim i As Long, j As Long, k As Long
    Dim ch As ChartObject
    
    Dim yVals As Variant
    Dim xVals As Variant
    For Each ch In Sheet1.ChartObjects
        ch.Delete
    Next ch
    For i = 1 To 7
        For j = 1 To 7
            Set ch = Sheet1.ChartObjects.Add((i - 1) * 240, (j - 1) * 105, 240, 105)
            yVals = Array()
            xVals = Array()
            For k = 1 To 30: append yVals, Rnd(): append xVals, k: Next k
            With ch.Chart
                .ChartType = xlLine
                .SeriesCollection.NewSeries
                With .SeriesCollection(1)
                    .XValues = xVals
                    .Values = yVals
                    'trendlines are to ensure there's enough junk on the screen
                    .Trendlines.Add Type:=xlLinear
                    .Trendlines.Add Type:=xlExponential
                End With

            End With
        Next j
    Next i


End Sub

'helper append -- why doesn't VBA have an append function?!
Public Sub append(ByRef arr As Variant, ByVal val As Variant)
    Dim lb As Long: lb = LBound(arr)
    Dim i As Long: i = UBound(arr) + 1
    ReDim Preserve arr(lb To i)
    arr(i) = val
End Sub

'output tooltip log with hwnd, caption, and style -- in practice, I set ctrl+shift+Q as a shortcut for this sub so I could test when the tooltip was visible
Public Sub logTips()
    Dim hWndTooltip As LongPtr
    Dim buffer As String * 256
    Dim length As Long
    Dim style As Long
    Dim foundCount As Long
    Dim results As String
    
    hWndTooltip = 0
    foundCount = 0
    results = "Tooltips:" & vbCrLf
    
    Do
        hWndTooltip = FindWindowEx(0, hWndTooltip, "tooltips_class32", vbNullString)
        If hWndTooltip <> 0 Then
            foundCount = foundCount + 1
            length = GetWindowText(hWndTooltip, buffer, 255)
            style = GetWindowLong(hWndTooltip, GWL_STYLE)
            results = results & "Tooltip #" & foundCount & ": hWnd = 0x" & Hex(hWndTooltip) & vbTab & _
                    "caption: " & Left(buffer, length) & vbTab & _
                    "style flags: " & Hex(style) & vbCrLf
        End If
    Loop While hWndTooltip <> 0
    Debug.Print results
End Sub