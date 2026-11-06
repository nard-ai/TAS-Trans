Attribute VB_Name = "PayrollModule"
Public cnPayroll As ADODB.Connection


Public Function Navigation(ByRef rsData As Recordset, _
                ByVal mycmd As CommandButton, ByRef cmdtop As CommandButton, _
                ByRef cmdPrev As CommandButton, ByRef cmdNext As CommandButton, _
                ByRef cmdlast As CommandButton)

    On Error Resume Next
    Select Case mycmd.Index
        Case 0 ' move to first record
            rsData.MoveFirst
            cmdtop.Enabled = False
            cmdPrev.Enabled = False
            cmdNext.Enabled = True
            cmdlast.Enabled = True
        Case 1 ' move to previous record
            rsData.MovePrevious
            cmdNext.Enabled = True
            cmdlast.Enabled = True
            If rsData.BOF Then
                rsData.MoveFirst
                cmdPrev.Enabled = False
                cmdtop.Enabled = False
                MsgBox "Beginning of file encountered.", vbInformation, PayrollAlert
            End If
        Case 2 ' move to next record
            rsData.MoveNext
            cmdPrev.Enabled = True
            cmdtop.Enabled = True
            If rsData.EOF Then
                rsData.MoveLast
                cmdNext.Enabled = False
                cmdlast.Enabled = False
                MsgBox "End of file encountered.", vbInformation, PayrollAlert
            End If
        Case 3 ' Move to last record
            rsData.MoveLast
            cmdtop.Enabled = True
            cmdPrev.Enabled = True
            cmdNext.Enabled = False
            cmdlast.Enabled = False
    End Select
End Function

Public Sub CustomizedHeader(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal mRange As String, ByVal Head1 As String, Optional Head2 As String = "", Optional HorAlign As Variant = xlCenter, _
    Optional VerAlign As Variant = xlCenter, Optional WrapText As Boolean = True, _
    Optional ShrinkToFit As Boolean = False, Optional MergedCells As Boolean = False, _
    Optional ColorIndex As Variant = xlNone, Optional FontName As String = "Arial", _
    Optional FontSize As Integer = 8, Optional FontBold As Boolean = False, Optional ColumnWidth As Integer = 7.5)
    
    xlSheet.Range(mRange).Select
    If Head2 <> "" Then
        xlApp1.ActiveCell.FormulaR1C1 = Head1 & Chr(10) & Head2
    Else
        xlApp1.ActiveCell.FormulaR1C1 = Head1
    End If
    With xlApp1
        .Selection.Interior.ColorIndex = ColorIndex '15 or xlnone
        .Selection.HorizontalAlignment = HorAlign 'xlCenter
        .Selection.VerticalAlignment = VerAlign 'xlCenter
        .Selection.WrapText = WrapText 'True
        .Selection.Orientation = 0
        .Selection.AddIndent = False
        .Selection.ShrinkToFit = ShrinkToFit 'False
        .Selection.MergeCells = MergedCells 'False
    End With
    With xlSheet.Range(mRange)
        .Font.Name = FontName '"Arial"    'Font.Name
        .Font.Size = FontSize '14         'Font.size
        .Font.Bold = FontBold 'True       'Font.Bold
    End With
    With xlSheet.Columns.Range(mRange)
        .ColumnWidth = ColumnWidth '7.5
    End With
End Sub

Public Sub PlaceBox(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal mRange As String)

'xlSheet.Range(mRange).Select
xlApp1.Range(mRange).Select
With xlApp1
    '.Selection.Columns.AutoFit
    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With xlApp1.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End With
'With xlApp1.Range(mRange)
'    .Font.Name = "Arial"   'Font.Name
'    .Font.Size = 8         'Font.size
'End With

End Sub

Public Sub SaveLoginStatus(ByVal mStatus As Integer, mUser As String, mPass As String)

StrQuery = "UPDATE UsersList SET LoginStatus = " & mStatus & _
                    " Where Username = '" & mUser & _
                    "' And Password = '" & mPass & "'"
Call Action_Query(StrQuery, cnnstr)

End Sub

Public Sub SaveUserTransactions(ByVal mDetails As String, ByVal mRemarks As String, _
    Optional mClient As String)

mDetails = Replace(Replace(mDetails, "(", ""), ")", "")
If LogoutTime = "" And LoginTime <> "" Then
    StrQuery = "INSERT INTO UserTransactions (UserID, UserName, LoginTime, TransDate, TransType, TransDetails, TransStatus, TransRemarks, AgencyName, ClientName) " & _
    "VALUES (" & operatorID & ", '" & LoginName & "', '" & LoginTime & "', '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "', '" & AccountType & "', '" & Replace(mDetails, "'", "''") & _
    "', 'N', '" & mRemarks & "', '" & CompanyShort & "', '" & mClient & "')"
ElseIf LoginTime <> "" Then
    StrQuery = "INSERT INTO UserTransactions (UserID, UserName, LoginTime, LogoutTime, TransDate, TransType, TransDetails, TransStatus, TransRemarks, AgencyName, ClientName) " & _
    "VALUES (" & operatorID & ", '" & LoginName & "', '" & LoginTime & "', '" & LogoutTime & "', '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "', '" & AccountType & "', '" & Replace(mDetails, "'", "''") & _
    "', 'N', '" & mRemarks & "', '" & CompanyShort & "', '" & mClient & "')"
End If
Call Action_Query(StrQuery, CnnstrHR)

End Sub

'Public Function InvalidDate(ByVal mDate As String) As Boolean
'Dim i As Integer
'Dim midDate As String
'Dim mmonth As String
'Dim mday As String
'
'InvalidDate = False
'mmonth = Mid(mDate, 1, 2)
'mday = Mid(mDate, 4, 2)
'If Val(mmonth) > 12 Or Val(mmonth) < 1 Then
'    InvalidDate = True
'ElseIf Val(mday) > 31 Or Val(mday) < 1 Then
'    InvalidDate = True
'Else
'    For i = 1 To 10
'        midDate = Mid(mDate, i, 1)
'        If i = 3 Or i = 6 Then
'            If midDate <> "/" Then
'                InvalidDate = True
'                Exit For
'            End If
'        ElseIf Not IsNumeric(midDate) Then
'            InvalidDate = True
'            Exit For
'        End If
'    Next
'End If
'End Function
'
'Public Function InvalidDateRange(ByVal mDate As String, ByVal mDate2 As String) As Boolean
'Dim i As Integer
'Dim date1 As Long
'Dim date2 As Long
'
'InvalidDateRange = False
'date1 = Val(Mid(Replace(mDate, "/", ""), 5, 4) & Mid(Replace(mDate, "/", ""), 1, 4))
'date2 = Val(Mid(Replace(mDate2, "/", ""), 5, 4) & Mid(Replace(mDate2, "/", ""), 1, 4))
'' Compare date2, it should be equal or greater than mdate
'If date2 < date1 Then
'    'MsgBox "End date range should be equal or greater than start date range!", vbInformation, TIS_ALERT
'    InvalidDateRange = True
'End If
'End Function
'
'Function GetDescription(ByRef rsTable As ADODB.Recordset, Fldname As Variant, KeyField As String, KeySw As Integer) As String
'Dim strFind As String
'Dim strSearch As String
'Dim varBookMark As Variant
'Dim tmpSearch As String
'
'With rsTable
'    If .RecordCount = 0 Then
'        GetDescription = "NOT FOUND"
'        Exit Function
'    End If
'    If .EOF Then
'        .MoveLast
'    ElseIf .BOF Then
'        .MoveFirst
'    End If
'    varBookMark = .Bookmark
'    .MoveFirst
'    If KeySw = 3 Then ' search for numeric type
'        strFind = Fldname & " = " & KeyField
'    Else
'        strFind = Fldname & " = '" & KeyField & "'"
'    End If
'    .Find strFind, 0, adSearchForward
'    ' Trap for BOF or EOF.
'    If .BOF Or .EOF Then    ' Search failed.
'        .Bookmark = varBookMark
'        GetDescription = "NOT FOUND"
'    Else
'        If KeySw = 1 Then  ' return customer name
'            GetDescription = Trim(!projectname)
'        ElseIf KeySw = 2 Then  ' return the county name
'            GetDescription = Trim(!countyname)
'        Else                    ' return the customer code
'            GetDescription = Trim(!CustCode)
'        End If
'    End If
'End With
'End Function
'
'Public Function mRpl(ByVal mytxt As String) As String
'    mRpl = Replace(mytxt, "'", "''")
'End Function
'


