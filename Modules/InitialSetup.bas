Attribute VB_Name = "PublicFunctions"
Public MMEmp As Boolean
Public FeMP As Boolean
Public mHRSEARCH As Boolean

''For DataEntry Screen
Public AddFlag As Boolean
Public EditFlag As Boolean
Public addflagship As Boolean
Public editflagship As Boolean

'For HR PersonalInfo
Public empno, m_empno, dept As String
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
    PayrollAlert = "Payroll System Alert"
    frmLogin.Show
End Sub

Public Sub ConnectDB()
   Dim strDrv As String
   'connect to database
    On Error GoTo Errh
    
    Set cnn = New ADODB.Connection
    cnn.CursorLocation = adUseClient
    cnn.CommandTimeout = 60
    cnn.ConnectionString = cnnStr
    cnn.Open
    
     Set cnnHR = New ADODB.Connection
    cnnHR.CursorLocation = adUseClient
    cnnHR.CommandTimeout = 60
    cnnHR.ConnectionString = CnnstrHR
    cnnHR.Open
    Exit Sub
Errh:
        MsgBox "Connection is not established. Pls contact your System Administrator!", vbCritical, CompanyName
        Set cnn = Nothing
        Set cnnHR = Nothing
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
        cnnStr = connStr
    End If
    Dim RsNewID As New ADODB.Recordset
    'Set RsNewID = ReturnRst("Select max(format(mid(controlno,7,5),'#####')) from HREMPLOYEES", cnnStr, mread)
    Set RsNewID = ReturnRst("Select MAX(format(mid(controlno,7,5),'#####')) from HREMPLOYEES where   mid(controlno,1,2)= MID(YEAR(Date()),3,2)", cnnStr, mread)
    If RsNewID.EOF = False Then
        If IsNull(RsNewID.Fields(0)) = False Then
            NewID = Val(RsNewID.Fields(0))
        Else
            NewID = 0
        End If
    End If
End Function

Public Function GetNewID() As String
    Dim RsNewID As New ADODB.Recordset
    'Set RsNewID = ReturnRst("Select max(format(mid(controlno,7,5),'#####')) from HREMPLOYEES", cnnStr, mread)
    'Set RsNewID = ReturnRst("Select MAX(format(mid(controlno,7,5),'#####')) from HREMPLOYEES where   mid(controlno,1,2)= MID(YEAR(Date()),3,2)", CnnstrHR, mread)
    Dim GetNewIDCnnstr As String
ReturnHere:
    If reContract = False Then
     Set RsNewID = ReturnRst("Select agencycode,controlyear,agencycount,headcount from control_no where agencycode=" & CompanyCode, CnnstrHR, mread)
     Else
      Set RsNewID = ReturnRst("Select agencycode,controlyear,agencycount,headcount from control_no where agencycode=" & destinationCompanyCode, CnnstrHR, mread)
   
     End If
     
     If RsNewID.EOF = True Then
      Call Action_Query("Insert into control_no(agencycode, controlyear,agencycount,headcount)values (" & IIf(reContract = True, destinationCompanyCode, CompanyCode) & ",0,0,0)", CnnstrHR)
      GoTo ReturnHere:
     End If
     
     
     If RsNewID.EOF = False Then
        If IsNull(RsNewID.Fields(0)) = False Then
            GetNewID = RsNewID!agencycode & "-" & RsNewID!controlyear & "-" & Format(Val(RsNewID!agencycount) + 1, "0000") & "-" & Format(Val(RsNewID!headcount) + 1, "00000")
        Else
            GetNewID = 0
        End If
      Call Action_Query("Update control_no set agencycount=" & Val(RsNewID!agencycount) + 1 & ",headcount=" & Val(RsNewID!headcount) + 1 & " where agencycode= " & IIf(reContract = True, destinationCompanyCode, CompanyCode) & " ", CnnstrHR)
      Call Action_Query("Update control_no set headcount=" & Val(RsNewID!headcount) + 1 & "", CnnstrHR)
    
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
    FormName.cmdDelete.Enabled = f
End Function

Function getMaxWAGEid(ByVal TableName As String) As Integer
    Dim rs As New ADODB.Recordset
    Set rs = ReturnRst("Select MAX(ID) from " & TableName, cnnStr, mstatic)
    If rs.EOF = False Then
    getMaxWAGEid = rs(0)
    End If
End Function
 Public Function GETuSERNAME(ByVal USERID As Integer) As String
    Dim rsUNAME As New ADODB.Recordset
    Set rsUNAME = ReturnRst("Select USERNAME from USERSLIST WHERE OPERATORID=" & USERID, cnnStr, mstatic)
    If rsUNAME.EOF = False Then
        GETuSERNAME = rsUNAME(0)
    Else
        GETuSERNAME = "SYSTEM"
    End If
 End Function
 
 
Public Function getAgencyName(ByVal AgencyShort As String) As String
        Dim rstAgency As New ADODB.Recordset
        Set rstAgency = ReturnRst("select agencyname,agencyshort,agencycode from companytable where agencyshort='" & AgencyShort & "'", getHRDBCnnstr, mstatic)
        If Not rstAgency.EOF Then
             getAgencyName = rstAgency!AgencyName
        Else
            getAgencyName = "Agency not found"
       End If
End Function

Public Function getAgencyNamebyID(ByVal companyID As Integer) As String
        Dim rstAgency As New ADODB.Recordset
        Set rstAgency = ReturnRst("select agencyname,agencyshort,agencycode from companytable where agencycode=" & companyID & "", getHRDBCnnstr, mstatic)
        If Not rstAgency.EOF Then
             getAgencyNamebyID = rstAgency!AgencyName
        Else
            getAgencyNamebyID = "Agency not found"
       End If
End Function

Public Function getClientNamebyID(ByVal clientID As Integer, cnnStr As String) As String
        Dim rstclient As New ADODB.Recordset
        Set rstclient = ReturnRst("select clientName,clientcode from clienttable where clientcode=" & clientID & "", cnnStr, mstatic)
        If Not rstclient.EOF Then
             getClientNamebyID = rstclient!ClientName
        Else
            getClientNamebyID = "client code " & clientID & " not found"
       End If
End Function

    
Public Function getHRDBCnnstr() As String
    Dim hrDB As String
    hrDB = GETINI("Company", "HRDatabase", "", App.Path & "\Application.ini")

    If hrDB = "" Or IsFileExists(hrDB) = False Then
        MsgBox "Please check application.ini file, HR database path not correct or not found!", vbCritical
        End
    Else
        getHRDBCnnstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & hrDB & "'"
    End If
End Function

    
Public Function getTargetCnnstr(ByVal companyID As Integer) As String
    Dim db As String
    Dim cShort As String
    cShort = getCompanyShort(companyID)
    db = GetSetting("DataSearch", cShort, "Path2Version2")

    If db = "" Or IsFileExists(db) = False Then
        MsgBox "Unable to connecto to " & cShort & ",please log in to this agency!", vbCritical
        End
    Else
        getTargetCnnstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & db & "'"
    End If
End Function

Public Function getCompanyShort(ByVal companyID As Integer) As String
        Dim rstAgency As New ADODB.Recordset
        Set rstAgency = ReturnRst("select agencyname,agencyshort,agencycode from companytable where agencycode=" & companyID & "", getHRDBCnnstr, mstatic)
        If Not rstAgency.EOF Then
             getCompanyShort = rstAgency!AgencyShort
        Else
            getCompanyShort = "Agency not found"
       End If
End Function


Public Function getAgencyAddress(ByVal CompanyCode As Integer) As String
        Dim rstAgency As New ADODB.Recordset
        Set rstAgency = ReturnRst("select agencyaddress from companytable where agencycode=" & CompanyCode & "", getHRDBCnnstr, mstatic)
        If Not rstAgency.EOF Then
             getAgencyAddress = rstAgency!agencyaddress
        Else
            getAgencyAddress = ""
       End If
End Function


Public Function getAgencyTelNo(ByVal CompanyCode As Integer) As String
        Dim rstAgency As New ADODB.Recordset
        Set rstAgency = ReturnRst("select agencytelno from companytable where agencycode=" & CompanyCode & "", getHRDBCnnstr, mstatic)
        If Not rstAgency.EOF Then
             getAgencyTelNo = rstAgency!AgencyTelNo
        Else
            getAgencyTelNo = ""
       End If
End Function

    
