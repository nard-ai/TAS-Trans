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
    Optional mID As Integer = 0, Optional mEmpNo As String = "", Optional mClient As String)

Dim mEmpInfo As String

'   To add more fields for recording of changes. April 25, 2015.
'    [EmployeeID] [int] NULL ,
'    [EmpNo] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [ClientCode] [int] NULL ,
'    [PaymentBasis] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
'    [MonthlyRate] [float] NULL ,
'    [DailyRate] [float] NULL ,
'    [StartContract] [datetime] NULL ,
'    [EndContract] [datetime] NULL ,

'mClient = frmHRFinalEmployees.cmbCompany ' comment by bernard july 18, 2016
mDetails = Replace(Replace(mDetails, "(", ""), ")", "")
If mRemarks = "Civil Status change." Or mRemarks = "Employment Status change." _
    Or mRemarks = "Date Hired change." Or mRemarks = "Start Contract change." _
    Or mRemarks = "End Contract change." Or mRemarks = "Payment Basis change." _
    Or mRemarks = "Daily Rate change." Or mRemarks = "Monthly Rate change." _
    Or mRemarks = "Cash Bond Rate change." Or mRemarks = "Accnt# change." _
    Or mRemarks = "Client change." Then

    mEmpInfo = "', " & IIf(mID = 0, frmHRFinalEmployees.TxtID, mID) & ", '" & _
            IIf(mEmpNo = "", frmHRFinalEmployees.TxtControlNo, mEmpNo) & "', " & _
            frmHRFinalEmployees.txtClientCode & ", '" & _
            frmHRFinalEmployees.cmbPayBasis & "', " & _
            frmHRFinalEmployees.txtMonthlyRate & ", " & _
            frmHRFinalEmployees.txtRate & ", '" & _
            frmHRFinalEmployees.mskStartContract & "', '" & _
            frmHRFinalEmployees.mskEndContract & "')"
       
    If LogoutTime = "" And LoginTime <> "" Then
        StrQuery = "INSERT INTO UserTransactions (UserID, UserName, LoginTime, TransDate, TransType, TransDetails, TransStatus, TransRemarks, AgencyName, ClientName, " & _
        "EmployeeID, EmpNo, ClientCode, PaymentBasis, MonthlyRate, DailyRate, StartContract, EndContract) " & _
        "VALUES (" & operatorID & ", '" & LoginName & "', '" & LoginTime & "', '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "', '" & AccountType & "', '" & Replace(mDetails, "'", "''") & _
        "', 'N', '" & mRemarks & "', '" & CompanyShort & "', '" & mClient & mEmpInfo
    ElseIf LoginTime <> "" Then
        StrQuery = "INSERT INTO UserTransactions (UserID, UserName, LoginTime, LogoutTime, TransDate, TransType, TransDetails, TransStatus, TransRemarks, AgencyName, ClientName, " & _
        "EmployeeID, EmpNo, ClientCode, PaymentBasis, MonthlyRate, DailyRate, StartContract, EndContract) " & _
        "VALUES (" & operatorID & ", '" & LoginName & "', '" & LoginTime & "', '" & LogoutTime & "', '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "', '" & AccountType & "', '" & Replace(mDetails, "'", "''") & _
        "', 'N', '" & mRemarks & "', '" & CompanyShort & "', '" & mClient & mEmpInfo
    End If
Else
    If LogoutTime = "" And LoginTime <> "" Then
        StrQuery = "INSERT INTO UserTransactions (UserID, UserName, LoginTime, TransDate, TransType, TransDetails, TransStatus, TransRemarks, AgencyName, ClientName, EmployeeID, EmpNo) " & _
        "VALUES (" & operatorID & ", '" & LoginName & "', '" & LoginTime & "', '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "', '" & AccountType & "', '" & Replace(mDetails, "'", "''") & _
        "', 'N', '" & mRemarks & "', '" & CompanyShort & "', '" & mClient & "', " & mID & ", '" & mEmpNo & "')"
    ElseIf LoginTime <> "" Then
        StrQuery = "INSERT INTO UserTransactions (UserID, UserName, LoginTime, LogoutTime, TransDate, TransType, TransDetails, TransStatus, TransRemarks, AgencyName, ClientName, EmployeeID, EmpNo) " & _
        "VALUES (" & operatorID & ", '" & LoginName & "', '" & LoginTime & "', '" & LogoutTime & "', '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "', '" & AccountType & "', '" & Replace(mDetails, "'", "''") & _
        "', 'N', '" & mRemarks & "', '" & CompanyShort & "', '" & mClient & "', " & mID & ", '" & mEmpNo & "')"
    End If
End If
StrQuery = Replace(StrQuery, "__/__/____", "")
Call Action_Query(StrQuery, CnnstrHR)

End Sub

Private Function ChkVal2(ByRef mTBox As TextBox) As String
    Dim mAmt As String
    If Trim(mTBox.Text) = "" Then
        mAmt = Trim(Replace(mTBox, "'", "''"))
        mAmt = "NULL"
    Else
    'Modified by Bernard, replace single quote with double quote for Height field.
        mAmt = "'" & Trim(Replace(mTBox.Text, "'", "''")) & "'"
    End If
    ChkVal = mAmt
End Function

