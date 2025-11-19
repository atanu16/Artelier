'==============================
' Auto-close messagebox in 3 sec
'==============================
#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxTimeout Lib "user32" Alias "MessageBoxTimeoutA" _
        (ByVal hwnd As LongPtr, ByVal text As String, ByVal caption As String, _
         ByVal type As Long, ByVal wlange As Long, ByVal timeout As Long) As Long
#Else
    Private Declare Function MessageBoxTimeout Lib "user32" Alias "MessageBoxTimeoutA" _
        (ByVal hwnd As Long, ByVal text As String, ByVal caption As String, _
         ByVal type As Long, ByVal wlange As Long, ByVal timeout As Long) As Long
#End If

Public Sub MsgBox3Sec(Msg As String, Optional Title As String = "Info")
    MessageBoxTimeout 0, Msg, Title, 0, 0, 3000
End Sub

' Helper: Find header column by name
Private Function GetHeaderCol(ws As Worksheet, headerName As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim(UCase(ws.Cells(1, c).Value)) = Trim(UCase(headerName)) Then
            GetHeaderCol = c
            Exit Function
        End If
    Next
    GetHeaderCol = 0
End Function

' Helper: Extract only text (remove numbers)
Private Function ExtractTextOnly(s As String) As String
    Dim i As Long, ch As String, outS As String
    outS = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not (ch >= "0" And ch <= "9") Then outS = outS & ch
    Next i
    ExtractTextOnly = Trim(outS)
End Function

'==================================
' MAIN FINAL MACRO
'==================================
Public Sub AppendData_Final()

    On Error GoTo ErrHandler

    Dim masterWB As Workbook, newWB As Workbook
    Dim masterWS As Worksheet, newWS As Worksheet
    Dim masterPath As String, newPath As String

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    Dim PortCol As Long, LoadCol As Long, PAPSCol As Long
    Dim PapsTextCol As Long, RecDateCol As Long, TypeCol As Long

    Dim newPortCol As Long, newLoadCol As Long, newPAPSCol As Long

    Dim lastMasterRow As Long, lastNewRow As Long
    Dim r As Long, uniqueID As String
    Dim receivedDate As String

    '--------------------------------------------
    ' Read paths from A2 (master) and A3 (new file)
    '--------------------------------------------
    masterPath = ThisWorkbook.Sheets(1).Range("A2").Value
    newPath = ThisWorkbook.Sheets(1).Range("A3").Value

    Set masterWB = Workbooks.Open(masterPath)
    Set masterWS = masterWB.Sheets(1)

    Set newWB = Workbooks.Open(newPath)
    Set newWS = newWB.Sheets(1)

    receivedDate = FileDateTime(newPath)

    '--------------------------------------------
    ' Get MASTER columns
    '--------------------------------------------
    PortCol = GetHeaderCol(masterWS, "Port of Crossing")
    LoadCol = GetHeaderCol(masterWS, "Load Order Number")
    PAPSCol = GetHeaderCol(masterWS, "PAPS Number")

    PapsTextCol = GetHeaderCol(masterWS, "PAPS Text")
    RecDateCol = GetHeaderCol(masterWS, "ReceivedDate")
    TypeCol = GetHeaderCol(masterWS, "Type")

    '--------------------------------------------
    ' Get NEW FILE columns
    '--------------------------------------------
    newPortCol = GetHeaderCol(newWS, "Port of Crossing")
    newLoadCol = GetHeaderCol(newWS, "Load Order Number")
    newPAPSCol = GetHeaderCol(newWS, "PAPS Number")

    '--------------------------------------------
    ' Build dictionary of existing Unique IDs
    '--------------------------------------------
    lastMasterRow = masterWS.Cells(masterWS.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastMasterRow
        uniqueID = Trim(masterWS.Cells(r, PortCol).Value) & "_" & _
                   Trim(masterWS.Cells(r, LoadCol).Value) & "_" & _
                   Trim(masterWS.Cells(r, PAPSCol).Value)

        dict(uniqueID) = 1
    Next r

    '--------------------------------------------
    ' Process NEW rows
    '--------------------------------------------
    lastNewRow = newWS.Cells(newWS.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastNewRow

        uniqueID = Trim(newWS.Cells(r, newPortCol).Value) & "_" & _
                   Trim(newWS.Cells(r, newLoadCol).Value) & "_" & _
                   Trim(newWS.Cells(r, newPAPSCol).Value)

        If Not dict.Exists(uniqueID) Then

            dict.Add uniqueID, 1
            lastMasterRow = lastMasterRow + 1

            ' Copy entire row from new to master
            newWS.Rows(r).Copy masterWS.Rows(lastMasterRow)

            '--------------------------------------
            ' Fill PAPS TEXT (text only)
            '--------------------------------------
            masterWS.Cells(lastMasterRow, PapsTextCol).Value = _
                ExtractTextOnly(newWS.Cells(r, newPAPSCol).Value)

            '--------------------------------------
            ' Fill ReceivedDate
            '--------------------------------------
            masterWS.Cells(lastMasterRow, RecDateCol).Value = receivedDate

            '--------------------------------------
            ' Fill TYPE based on LoadOrderNumber
            '--------------------------------------
            Dim firstDigit As String
            firstDigit = Left(Trim(newWS.Cells(r, newLoadCol).Value), 1)

            If firstDigit = "2" Or firstDigit = "4" Then
                masterWS.Cells(lastMasterRow, TypeCol).Value = "OIL"
            Else
                masterWS.Cells(lastMasterRow, TypeCol).Value = "MEAL"
            End If

        End If

    Next r

    masterWB.Save
    newWB.Close False

    MsgBox3Sec "Appending completed successfully!", "Done"

    Exit Sub

ErrHandler:
    MsgBox3Sec "Error: " & Err.Description, "Error"
End Sub