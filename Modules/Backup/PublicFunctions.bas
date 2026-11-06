Attribute VB_Name = "PublicFunctions"
Public MMEmp As Boolean
Public FeMP As Boolean
Public mHRSEARCH As Boolean
Public HeadCountCtr As Integer

''For DataEntry Screen
Public AddFlag As Boolean
Public EditFlag As Boolean
Public addflagship As Boolean
Public editflagship As Boolean

'For HR PersonalInfo
Public EmpNo, m_empno, dept As String
Public city As Boolean
Public birth, hired As Boolean

'For HR Leaves
Public leaveserved As Boolean

'For Report
Public strReport As String
Public strSQLReport As String

Enum GEnmCheckingValues1
    Exists = 0
    notExists = 1
End Enum

Sub Main()
Dim ServerName As String
Dim DbName As String
Dim ConnString As String

    PayrollAlert = "HR and Payroll System Alert"
'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=HRAndPayroll;Data Source=.

'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=HRAndPayroll;Data Source=.
    If IsFileExists(App.Path & "\Application.ini") = False Then
        MsgBox "Application.ini file is not found in program directory. Please create", vbInformation, PayrollAlert
        End
    End If
    
    ServerName = GETINI("Database", "ServerName", "", App.Path & "\Application.ini")
    DbName = GETINI("Database", "DBName", "", App.Path & "\Application.ini")
    ConnString = GETINI("Database", "ConnectionString", "", App.Path & "\Application.ini")
    
    LoadPictureFlag = GETINI("LoadPicture", "LoadPicturePath", "", App.Path & "\Application.ini")
    SLWithPay = GETINI("LeaveWithPay", "SickLeave", "", App.Path & "\Application.ini")
    VLWithPay = GETINI("LeaveWithPay", "VacationLeave", "", App.Path & "\Application.ini")
    HalfWithPay = GETINI("LeaveWithPay", "HalfDay", "", App.Path & "\Application.ini")

    cnnstr = ConnString & ";Initial Catalog=" & DbName & ";Data Source=" & ServerName
    
    CnnstrHR = cnnstr
    
    'ReportConnection = GETINI("ReportConnection", "ReportCon", "", App.Path & "\Application.ini")
    
    'frmLogin.Show
    frmScreen.Show
End Sub

Public Sub ConnectDB()
   Dim strDrv As String
   
   'Connect to SQL Server Database.
    On Error GoTo Errh
    
    Set cnn = New ADODB.Connection
    cnn.CursorLocation = adUseClient
    cnn.CommandTimeout = 60
    cnn.ConnectionString = cnnstr
    cnn.Open
    
    Exit Sub

Errh:
        MsgBox "Connection is not established. Pls contact your System Administrator!", vbCritical, PayrollAlert
        Set cnn = Nothing
        End
    MsgBox Err.Number & " " & Err.Description, vbCritical, CompanyName
    Err.Clear
End Sub

Public Function OpenDatabaseAdo(ByRef pdbconinput As ADODB.Connection, _
    ByVal pstrconnectionstring As String, _
    Optional ByVal pstrusername As String = "", _
    Optional ByVal pstrpassword As String = "", _
    Optional ByVal pintoptions As Integer = -1) _
    As Integer
    On Error GoTo errdesc:
    Set pdbconinput = New ADODB.Connection
        pdbconinput.Open pstrconnectionstring, pstrusername, pstrpassword, pintoptions
        OpenDatabaseAdo = 0
         Exit Function
errdesc:
        OpenDatabaseAdo = 1
End Function

Public Function IsTableEXIST(ByRef pdbconinput As ADODB.Connection, PStrtablename As String) As GEnmCheckingValues1
    Dim intretval As String
    Dim blnretvalue As Boolean
    Dim rectemp1 As ADODB.Recordset
On Error GoTo Err_Handler:
    strSQL = " select * From " & PStrtablename & ""
    Set rectemp1 = New ADODB.Recordset
    rectemp1.Open strSQL, pdbconinput
    IsTableEXIST = Exists
    rectemp1.Close
    Set rectemp1 = Nothing
    Exit Function
Err_Handler:
    IsTableEXIST = notExists

End Function

Public Function NewID(Optional ByVal connStr As String = "") As String
    If connStr <> "" Then
        cnnstr = connStr
    End If
    Dim RsNewID As New ADODB.Recordset
    'Set RsNewID = ReturnRst("Select max(format(mid(EmpNo,7,5),'#####')) from HREMPLOYEES", cnnStr, mread)
    Set RsNewID = ReturnRst("Select MAX(format(mid(EmpNo,7,5),'#####')) from HREMPLOYEES where   mid(EmpNo,1,2)= MID(YEAR(Date()),3,2)", cnnstr, mstatic)
    If RsNewID.EOF = False Then
        If IsNull(RsNewID.Fields(0)) = False Then
            NewID = Val(RsNewID.Fields(0))
        Else
            NewID = 0
        End If
    End If
End Function

Public Function GetNewDTRID() As Integer
    Dim RsNewID As New ADODB.Recordset

    ' Get New DTR ID.
    StrQuery = "Select MAX(EmployeeID) AS NewID FROM HREmployees where Agencycode = " & CompanyCode
    Set RsNewID = ReturnRst(StrQuery, cnnstr, mstatic)
    GetNewDTRID = RsNewID!NewID + 1
    Set RsNewID = Nothing

End Function

Public Function GetNewID(Optional UpdateEmpNoSw As Boolean = False) As String
    Dim RsNewID As New ADODB.Recordset

    ' If successfully save, update employee number counter.
'    If UpdateEmpNoSw = True Then
'        Call Action_Query("UPDATE EmpControlNo SET agencycount = " & HeadCountCtr & ", headcount = " & HeadCountCtr & " WHERE agencycode= " & CompanyCode & " ", cnnstr)
'        'GetNewID = TxtControlNo.Text
'        Exit Function
'    End If

ReturnHere:
    
    If reContract = False Then
        StrQuery = "Select agencycode, controlyear, agencycount, headcount from EmpControlNo where agencycode = " & CompanyCode
        Set RsNewID = ReturnRst(StrQuery, cnnstr, mstatic)
    Else
        StrQuery = "Select agencycode, controlyear, agencycount, headcount from EmpControlNo where agencycode = " & CompanyCode
        Set RsNewID = ReturnRst(StrQuery, cnnstr, mstatic)
        'Set RsNewID = ReturnRst("Select agencycode,controlyear,agencycount,headcount from EmpControlNo where agencycode = " & CompanyCode, CnnstrHR, mread)
        'For Update 2010
        'Set RsNewID = ReturnRst("Select agencycode,controlyear,agencycount,headcount from EmpControlNo where agencycode=" & destinationCompanyCode, CnnstrHR, mread)
    End If
    
    'Insert new agency.
    If RsNewID.EOF = True Then
        StrQuery = "Insert into EmpControlNo(agencycode, controlyear, agencycount, headcount) values (" & CompanyCode & ", 0, 0, 0)"
        Call Action_Query(StrQuery, cnnstr)
        GoTo ReturnHere:
    ElseIf UpdateEmpNoSw = True Then
        StrQuery = "UPDATE EmpControlNo SET ControlYear = '" & Format(Now, "yy") & "', AgencyCount = " & RsNewID!AgencyCount + 1 & ", HeadCount = " & RsNewID!Headcount + 1
        Call Action_Query(StrQuery, cnnstr)
        Set RsNewID = Nothing
        Exit Function
    End If
     
    If RsNewID.EOF = False Then
        If IsNull(RsNewID.Fields(0)) = False Then
            'GetNewID = RsNewID!AgencyCode & "-" & Trim$(RsNewID!ControlYear) & "-" & Format(Val(RsNewID!AgencyCount) + 1, "0000") & "-" & Format(Val(RsNewID!HeadCount) + 1, "00000")
            GetNewID = Format(Now, "yy") & "-" & Format(Val(RsNewID!Headcount) + 1, "0000")
        Else
            GetNewID = Format(Now, "yy") & "-00000"
        End If
    Else
        Exit Function
    End If
    
   Set RsNewID = Nothing
   
End Function


Public Function ValidateButton(ByVal FormName As Form, Optional ByVal a As Boolean, Optional ByVal b As Boolean, Optional ByVal c As Boolean, Optional ByVal D As Boolean, Optional ByVal e As Boolean, Optional ByVal f As Boolean)
    FormName.cmdadd.Enabled = a
    FormName.cmdedit.Enabled = b
    FormName.cmdsave.Enabled = c
    FormName.cmdcancel.Enabled = D
    FormName.cmdrefresh.Enabled = e
    FormName.cmddelete.Enabled = f
End Function

Function getMaxWAGEid(ByVal TableName As String) As Integer
    Dim rs As New ADODB.Recordset
    Set rs = ReturnRst("Select MAX(ID) from " & TableName, cnnstr, mstatic)
    If rs.EOF = False Then
    getMaxWAGEid = rs(0)
    End If
End Function

 Public Function GETuSERNAME(ByVal USERID As Integer) As String
    Dim rsUNAME As New ADODB.Recordset
    Set rsUNAME = ReturnRst("Select USERNAME from USERSLIST WHERE OPERATORID = " & USERID, cnnstr, mstatic)
    If rsUNAME.EOF = False Then
        GETuSERNAME = rsUNAME(0)
    Else
        GETuSERNAME = "SYSTEM"
    End If
 End Function
 
Public Function getAgencyName(ByVal AgencyShort As String) As String
        Dim rstAgency As New ADODB.Recordset
        Set rstAgency = ReturnRst("select agencyname, agencyshort, agencycode from companytable where agencyshort='" & AgencyShort & "'", cnnstr, mstatic)
        If Not rstAgency.EOF Then
             getAgencyName = rstAgency!AgencyName
        Else
            getAgencyName = "Agency not found"
       End If
End Function

Public Function getAgencyNamebyID(ByVal companyID As Integer) As String
        Dim rstAgency As New ADODB.Recordset
        
        Set rstAgency = ReturnRst("select agencyname, agencyshort, agencycode from CompanyTable where agencycode = " & companyID & "", cnnstr, mstatic)
        If Not rstAgency.EOF Then
             getAgencyNamebyID = rstAgency!AgencyName
        Else
            getAgencyNamebyID = "Agency not found"
       End If
End Function

Public Function getClientNamebyID(ByVal clientID As Integer, cnnstr As String) As String
    Dim rstclient As New ADODB.Recordset
        
'    If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'        StrQuery = "select clientName,clientcode from clienttable where CLIENTCODE2 = 'FMS' AND clientcode = " & clientID
'    Else
'        StrQuery = "select clientName,clientcode from clienttable where clientcode = " & clientID
'    End If
    StrQuery = GetClientRecord(clientID, False, , " ClientName, ClientCode ")
    Set rstclient = ReturnRst(StrQuery, cnnstr, mstatic)
        
    'Set rstclient = ReturnRst("select clientName,clientcode from clienttable where clientcode=" & clientID & "", cnnstr, mstatic)
    If Not rstclient.EOF Then
       getClientNamebyID = rstclient!ClientName
    Else
       getClientNamebyID = "client code " & clientID & " not found"
    End If
    
    Set rstclient = Nothing
    
End Function

Public Function getClientNamebyEmployeeID(ByVal EmployeeID As Integer) As String
        
    Dim empClientCode As Integer
    Dim rsemp As New ADODB.Recordset
    
    Set rsemp = ReturnRst("Select ClientCode from HREMPLOYEES where EmployeeID = " & EmployeeID & "", CnnstrHR, mstatic)
    If rsemp.EOF = False Then
        empClientCode = rsemp.Fields!ClientCode
    Else
        empClientCode = 0
    End If
    
    Dim rstclient As New ADODB.Recordset
    
'    If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'        StrQuery = "Select clientName, clientcode FROM clienttable where CLIENTCODE2 = 'FMS' AND clientcode = " & empClientCode
'    Else
'        StrQuery = "Select clientName, clientcode FROM clienttable where clientcode = " & empClientCode
'    End If
    StrQuery = GetClientRecord(empClientCode, False, , " ClientName, ClientCode")
    Set rstclient = ReturnRst(StrQuery, cnnstr, mstatic)
    
    'Set rstclient = ReturnRst("Select clientName, clientcode FROM clienttable where clientcode = " & empClientCode & "", cnnstr, mstatic)
    If Not rstclient.EOF Then
        getClientNamebyEmployeeID = RTrim(rstclient!ClientName)
    Else
        getClientNamebyEmployeeID = " EmployeeID : " & EmployeeID & ";clientCode:" & empClientCode & " not found"
    End If

    Set rstclient = Nothing
    Set rsemp = Nothing
        
End Function

Public Function CheckHRCodes(ByVal mstr As String) As Boolean
Dim mRSCheck As ADODB.Recordset

Set mRSCheck = ReturnRst(mstr, cnnstr, mstatic)
If mRSCheck.EOF = False Then
    CheckHRCodes = True
Else
    CheckHRCodes = False
End If
Set mRSCheck = Nothing

End Function

Public Function getHRDBCnnStr() As String
    Dim hrDB As String
    Dim CutOffSetting As String     ' updated by Bernard 07/19/2008
    
'    hrDB = GETINI("Company", "HRDatabase", "", App.Path & "\Application.ini")
'
'    CutOffSetting = GETINI("PayrollCutOff", "CutOff", "", App.Path & "\Application.ini")
'    CutOffPeriod = CutOffSetting
'
'    If CutOffPeriod <> "END" And CutOffPeriod <> "START" Then
'        CutOffPeriod = "END"
'    End If
'
'    If hrDB = "" Or IsFileExists(hrDB) = False Then
'        MsgBox "Please check application.ini file, HR database path not correct or not found!", vbCritical
'        End
'    Else
'        cnnstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & hrDB & "'"
'    End If
    
End Function
    
Public Function getTargetCnnstr(ByVal companyID As Integer) As String
    Dim db As String
    Dim cShort As String
    cShort = getCompanyShort(companyID)
    db = GetSetting("DataSearch", cShort, "Path2Version2")

    If db = "" Or IsFileExists(db) = False Then
        MsgBox "Unable to connect to " & cShort & ", please log in to this agency!", vbCritical
        End
    Else
        getTargetCnnstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & db & "'"
    End If
End Function

Public Function getCompanyShort(ByVal companyID As Integer) As String
        Dim rstAgency As New ADODB.Recordset
        Set rstAgency = ReturnRst("select agencyname,agencyshort,agencycode from companytable where agencycode=" & companyID & "", cnnstr, mstatic)
        If Not rstAgency.EOF Then
             getCompanyShort = rstAgency!AgencyShort
        Else
            getCompanyShort = "Agency not found"
       End If
End Function


Public Function getAgencyAddress(ByVal CompanyCode As Integer) As String
        Dim rstAgency As New ADODB.Recordset
        Set rstAgency = ReturnRst("select agencyaddress from companytable where agencycode=" & CompanyCode & "", cnnstr, mstatic)
        If Not rstAgency.EOF Then
             getAgencyAddress = IIf(IsNull(rstAgency!AgencyAddress), "No Address specified.", rstAgency!AgencyAddress)
        Else
            getAgencyAddress = ""
       End If
End Function

Public Function getLogoPath(ByVal CompanyCode As Integer) As String
     Dim rstAgency As New ADODB.Recordset
     Set rstAgency = ReturnRst("select logoPath from CompanyTable where agencycode = " & CompanyCode & "", cnnstr, mstatic)
     If Not rstAgency.EOF Then
        getLogoPath = rstAgency!logoPath
     Else
        getLogoPath = ""
    End If
End Function

Public Function getAgencyTelNo(ByVal CompanyCode As Integer) As String
        Dim rstAgency As New ADODB.Recordset
        Set rstAgency = ReturnRst("select agencytelno from companytable where agencycode=" & CompanyCode & "", cnnstr, mstatic)
        If Not rstAgency.EOF Then
             getAgencyTelNo = rstAgency!AgencyTelNo
        Else
            getAgencyTelNo = ""
       End If
End Function

' Updated by Bernard Sept. 6, 2008.
Public Sub ProtectDatabaseRecords(ByVal mComCode As Integer, ByVal mCCode As Integer, ByVal SetValue As Boolean, _
    Optional AllRec As Boolean = False)

If AllRec = False Then
    StrQuery = "UPDATE EmployeeDeductions SET NoEditFlag = " & IIf(SetValue = True, 1, 0) & _
                " Where AgencyCode = " & mComCode & " and ClientCode = " & mCCode
Else
    StrQuery = "UPDATE EmployeeDeductions SET NoEditFlag = " & IIf(SetValue = True, 1, 0) & _
                " Where AgencyCode = " & mComCode
End If
Call Action_Query(StrQuery, CnnstrHR)

End Sub

Public Sub FillClientList(ByRef mLsEmp As ListView, _
                        mRs As Recordset)
Dim litem1 As ListItem

    mLsEmp.View = lvwReport
    mLsEmp.ListItems.Clear
    While Not (mRs.EOF Or mRs.BOF)
        Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!ClientName), "Unknown Client", Replace(mRs!ClientName, "'", "''")))
        litem1.SubItems(1) = mRs(0)
        mRs.MoveNext
    Wend
    
End Sub

Public Sub PrintReportHeader(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal mRange As String, ByVal Head1 As String, Optional Head2 As String = "", Optional HorAlign As Variant = xlCenter, _
    Optional VerAlign As Variant = xlCenter, Optional WrapText As Boolean = True, _
    Optional ShrinkToFit As Boolean = False, Optional MergedCells As Boolean = False, _
    Optional ColorIndex As Variant = xlNone, Optional FontName As String = "Arial", _
    Optional FontSize As Integer = 7, Optional FontBold As Boolean = False, Optional ColumnWidth As Integer = 7.5)
    
    'xlSheet.Range(mRange).Select
    xlApp1.Range(mRange).Select
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
    With xlApp1.Range(mRange)
        .Font.Name = FontName '"Arial"    'Font.Name
        .Font.Size = FontSize '14         'Font.size
        .Font.Bold = FontBold 'True       'Font.Bold
    End With
    With xlApp1.Range(mRange)
        .ColumnWidth = ColumnWidth '7.5
    End With
End Sub

Public Function getMonthNo(ByRef mMonth As String) As String

Select Case mMonth
    Case "January"
        getMonthNo = "01"
        QtrNo = 1
        DTRNo = 1
    Case "February"
        getMonthNo = "02"
        QtrNo = 1
        DTRNo = 2
    Case "March"
        getMonthNo = "03"
        QtrNo = 1
        DTRNo = 3
    Case "April"
        getMonthNo = "04"
        QtrNo = 2
        DTRNo = 1
    Case "May"
        getMonthNo = "05"
        QtrNo = 2
        DTRNo = 2
    Case "June"
        getMonthNo = "06"
        QtrNo = 2
        DTRNo = 3
    Case "July"
        getMonthNo = "07"
        QtrNo = 3
        DTRNo = 1
    Case "August"
        getMonthNo = "08"
        QtrNo = 3
        DTRNo = 2
    Case "September"
        getMonthNo = "09"
        QtrNo = 3
        DTRNo = 3
    Case "October"
        getMonthNo = "10"
        QtrNo = 4
        DTRNo = 1
    Case "November"
        getMonthNo = "11"
        QtrNo = 4
        DTRNo = 2
    Case "December"
        getMonthNo = "12"
        QtrNo = 4
        DTRNo = 3
End Select

End Function

Public Function getMonthNum(ByVal mMonth As String, ByVal mYear As Integer, _
    Optional mSw As Boolean = True, Optional mSw1 As Boolean = False, _
    Optional mSw2 As Boolean = False) As Integer

Dim mYear2 As Integer, Lastday As Integer

Select Case mMonth
    Case "January"
        getMonthNum = 12
        mYear2 = mYear - 1
        Lastday = 31
    Case "February"
        getMonthNum = 1
        mYear2 = mYear
        Lastday = 28
    Case "March"
        getMonthNum = 2
        mYear2 = mYear
        Lastday = 31
    Case "April"
        getMonthNum = 3
        mYear2 = mYear
        Lastday = 30
    Case "May"
        getMonthNum = 4
        mYear2 = mYear
        Lastday = 31
    Case "June"
        getMonthNum = 5
        mYear2 = mYear
        Lastday = 30
    Case "July"
        getMonthNum = 6
        mYear2 = mYear
        Lastday = 31
    Case "August"
        getMonthNum = 7
        mYear2 = mYear
        Lastday = 31
    Case "September"
        getMonthNum = 8
        mYear2 = mYear
        Lastday = 30
    Case "October"
        getMonthNum = 9
        mYear2 = mYear
        Lastday = 31
    Case "November"
        getMonthNum = 10
        mYear2 = mYear
        Lastday = 30
    Case "December"
        getMonthNum = 11
        mYear2 = mYear
        Lastday = 31
End Select

If mSw1 = True Then     ' return the next month
    If getMonthNum = 12 Then
        getMonthNum = 1
    End If
    If mSw = True Then ' return the next year
        If getMonthNum = 12 Then
            getMonthNum = mYear2 + 1
        Else
            getMonthNum = mYear2
        End If
    End If
Else
    If mSw = False Then ' return the year
        getMonthNum = mYear2
    End If
End If

If mSw2 = True Then
    getMonthNum = Lastday
End If

End Function

Public Sub PrintHeadings(ByRef xlSheet As Excel.Worksheet, ByRef xlApp1 As Excel.Application, _
    ByVal t As Integer, ByVal mPos As Integer, Optional mPlus As Integer = 0, Optional Headsw As Boolean = True)

Dim mHeader1 As String
Dim mHeader2 As String
Dim mRange As String
Dim mSize As Integer

    'mRange = Chr(t) & 6
    mRange = Chr(t + mPlus) & mPos & ":" & Chr(t + mPlus) & mPos + 1
    Select Case t
        Case 65:
            mHeader1 = "No."
            mHeader2 = ""
            mSize = 4.14
        Case 66:
            mHeader1 = "Client Name"
            mHeader2 = ""
            mSize = 50
        Case 67:
            mHeader1 = "Total"
            mHeader2 = "Employees"
            mSize = 7.7
        Case 68:
            If Headsw = True Then
                mHeader1 = "PS"
                mHeader2 = "Share"
            Else
                mHeader1 = "Actual"
                mHeader2 = "EE"
            End If
            mSize = 7.7
        Case 69:
            If Headsw = True Then
                mHeader1 = "ES"
                mHeader2 = "Share"
            Else
                mHeader1 = "Actual"
                mHeader2 = "ER"
            End If
            mSize = 7.7
        Case 70:
            If Headsw = True Then
                mHeader1 = "PS+ES"
                mHeader2 = "Share"
            Else
                mHeader1 = "Actual"
                mHeader2 = "EE+ER Share"
            End If
            mSize = 8
        Case 71:
            If Headsw = True Then
                mHeader1 = "Adjusted"
                mHeader2 = "PS"
            Else
                mHeader1 = "Remitted"
                mHeader2 = "EE"
            End If
            mSize = 9.5
        Case 72:
            If Headsw = True Then
                mHeader1 = "Adjusted"
                mHeader2 = "ER"
            Else
                mHeader1 = "Remitted"
                mHeader2 = "ER"
            End If
            mSize = 9.5
        Case 73:
            If Headsw = True Then
                mHeader1 = "Total"
                mHeader2 = "PS+ES"
            Else
                mHeader1 = "Total"
                mHeader2 = "EE+ER"
            End If
            mSize = 9.5
        Case 74:
            If Headsw = True Then
                mHeader1 = "Difference"
                mHeader2 = "PS"
            Else
                mHeader1 = "Difference"
                mHeader2 = "EE"
            End If
            mSize = 10.5
        Case 75:
            If Headsw = True Then
                mHeader1 = "Difference"
                mHeader2 = "ES"
            Else
                mHeader1 = "Difference"
                mHeader2 = "ER"
            End If
            mSize = 10.5
        Case 76:
            mHeader1 = "Total"
            mHeader2 = "Difference"
            mSize = 9.5
        Case 77:
            mHeader1 = "SBR#"
            mHeader2 = "OR#"
            mSize = 7.5
        Case 78:
            mHeader1 = "Date Paid"
            mHeader2 = ""
            mSize = 7.5
        Case 79:
            mHeader1 = "Date Posted"
            mHeader2 = ""
            mSize = 7.5
        Case 80:
            mHeader1 = "Remarks"
            mHeader2 = ""
            mSize = 10.5
    End Select
    Call PrintReportHeader(xlSheet, xlApp1, mRange, mHeader1, mHeader2, xlCenter, xlCenter, True, False, True, 15, , 8, False, mSize)
End Sub

Public Sub DisplayMessage(ByVal mMsg As String, Optional CloseMsg As Boolean = False)

    If CloseMsg = False Then
        frmProgress.msgDisplay.Caption = mMsg
        If frmProgress.Visible = False Then
            Call frmProgress.Show
        End If
    Else
        Unload frmProgress
    End If

End Sub

Public Function GetClientRecord(ByVal CCode As Integer, _
    Optional ActiveOnly As Boolean = False, Optional OrderBy As String = "", _
    Optional FieldList As String = " * ") As String
Dim mWhere As String

    If CCode = 0 Then
        mWhere = " WHERE ClientCode <> 0"
    Else
        mWhere = " WHERE ClientCode = " & CCode
    End If
    If Left(AccountType, 3) = "FMS" Then ' Drivers and Conductors only
        mWhere = mWhere & " AND CLIENTCODE2 = 'FMS'"
    ElseIf UCase(Left(AccountType, 3)) = "STA" Then ' Admin and Staff only
        mWhere = mWhere & " AND CLIENTCODE2 <> 'FMS' OR ClientCode2 IS NULL "
    End If
    If ActiveOnly = True Then
        mWhere = mWhere & " AND ClientStatus = 1"
    End If
    GetClientRecord = "SELECT " & FieldList & " FROM ClientTable " & mWhere & OrderBy

End Function

Public Function ForPayroll(AcctType As String) As Boolean
If AcctType = "Accounting" Or AcctType = "Administrator" Or AcctType = "FMS Payroll" Or AcctType = "FMS Admin" Then
    ForPayroll = True
Else
    ForPayroll = False
End If
End Function
