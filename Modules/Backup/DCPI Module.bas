Attribute VB_Name = "Module1"
Public MMEmp As Boolean
Public FeMP As Boolean
Public mHRSEARCH As Boolean


'For User Login
Public cnDCPI As ADODB.Connection
Public ISVisibleHRF As Boolean
Public DC As String 'Data Control
Public pr As String 'Production
Public QC As String 'Quality Control
Public SY As String 'Systems
Public PG As String 'Programs
Public TE As String 'Technical
Public HR As String 'Human Resources
Public AC As String 'Accounting
Public IV As String 'Inventory
Public UL As String 'User Level
Public UN As String 'User Name
Public DEP As Integer 'Department
Public emp As String  'Employee No in Login

'DCTrackShipment
Public Projects As String
Public County As String
Public DateRangeFrom As String
Public DateRangeTo As String
Public cnDCPI_User As ADODB.Connection
'Public cnDCPI As ADODB.Connection

'For Filtering
Public ProjectShort As String
Public Turnaroud As String
Public Turnaround_Days As String
Public projectcode As String
Public countycode As String
Public projectlong As String
Public countylong As String
Public fromdate As String
Public todate As String
Public DefCourier As String
Public DecDocumentType As String
Public DefBatchID As String
Public DefBatchSubID As String

'For DataEntry Screen
Public AddFlag As Boolean
Public EditFlag As Boolean
Public addflagship As Boolean
Public editflagship As Boolean
Public flgedit As Boolean
Public flgnew As Boolean
Public addflagline As Boolean
Public editflagline As Boolean
Public addflagbatch As Boolean
Public editflagbatch As Boolean
Public dblclick As Boolean
Public lev1, lev2, lev3 As Boolean  'used for pointing the record chose in dctrackshipment
Public m_notcnum As String
Public m_linenum As String
Public m_Subnum As String
Public NoticeFlag As Boolean
Public LineFlag As Boolean
Public SubFlag As Boolean
Public lineadd As Boolean
Public subadd As Boolean
Public ifedit, ifadd As Boolean
'For Production
Public dcedit As Boolean
Public prodedit As Boolean
'For Data Control Views
Public ViewClick As Integer
Public FromProduction As Boolean
Public edilev1 As Boolean
Public edilev2 As Boolean
Public edilev3 As Boolean
Public addleave As Boolean
Public editleave As Boolean

'For ClientFeedback Screen
Public m_feedbacknum As Integer
Public m_feeedbacknum As Integer
Public addflagfbdet As Boolean
Public editflagfbdet As Boolean
Public addflagfb As Boolean
Public editflagfb As Boolean

'For Transmission Screen
Public m_transactnum As Integer
Public filenmalpha As String
Public filefrom As String
Public fileto As String
Public filenm As String
Public numfrom As Integer
Public numto As Integer
'Public alpha As Boolean
'Public alltx As Boolean
'Public seltx As Boolean
Public selalldt As Boolean
Public optAll As Boolean
Public seldt As Boolean
Public filt As Boolean
Public addqc As Boolean
Public editqc As Boolean
Public start_doc As Long
Public end_doc As Long
Public start_page As Integer
Public end_page As Integer
Public start_recdt As Date
Public end_recdt As Date
Public proj_trans As Date
Public transmited As Boolean
Public list As ListItem
Public zipfile As Boolean
Public sublineno As Integer
'add by Jun
Public FileAlpha As String
Public selstr As String
Public txtFinalFilterPR As String
Public pn, cn As String
Public prj, cnty As Integer

'For HR PersonalInfo
Public atndctym As Date
Public empno, m_empno, dept As String
Public appnd, serch, certcty, bday As Boolean
Public provadd, city As Boolean
Public birth, hired As Boolean
Public married, sngle As Boolean
Public mempnum, mfullnm As String
Public mdegree As String
Public dependent As Boolean
Public rec_no As Integer

'For HR ApplicantcInfo
Public blisted As Boolean
Public appfailed, apppassed, exam As Boolean
Public appintfailed, appintpassed, interview As Boolean
Public appmedfailed, appmedpassed, medical As Boolean
Public degree, bydate As Boolean
Public applicantpassed As Boolean
'For HR Movement
Public Department As Boolean
Public EmpHis As Boolean
'For HR Offense
Public discoffense As Boolean
Public profile As Boolean
Public GivenMonth As Boolean
'For Production Projected transmission
Public PRNotice As Boolean
Public PRProjTrans As Boolean
'For HR Training
Public training As Boolean
Public IndTraining As Boolean
Public GivenTraining As Boolean
Public ServingTraining As Boolean
Public Finishedtraining As Boolean
'For HR Leaves
Public leaves As Boolean
Public maternity As Boolean
Public paternity As Boolean
Public leaveserved As Boolean
'For HR PRoject Assignment
Public projemployee As Boolean

'For Job Request
Public VarDeptChoice As String
Public mdept As String
Public mdeptcd As Integer
Public VarDeptChoiceNum As Integer
'For Inventory
Public item_num As Integer
Public details, assgn As Boolean
Public ass_num As Integer

'For Report
Public strReport As String
Public strSQLReport As String

Enum GEnmCheckingValues1
    Exists = 0
    notExists = 1
End Enum

Sub Main()
    Call InitialAppSetup
    PayrollAlert = "Payroll System Alert"
    
  '  If FrmSetDatabase.Visible = True Then
   '     Unload frmloading
   ' Else
    '    frmloading.Show
    'End If
End Sub

Public Sub InitialAppSetup()
    Dim strcondb    As String
    Dim intretval   As Integer
    'strDrv = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Search"
    'Cnnstr = strDrv
    'StrDbpath = "c:\PayrollDev\SEARCH\Database\SearchAsiaData.mdb"
    CompanyCode = GETINI("Company", "CompanyName", "", App.Path & "\Application.ini")
    If CompanyCode = 1 Then 'searchasia
        StrDbpath = GetSetting("DataSearch", "SE", "Path2")
    ElseIf CompanyCode = 2 Then
        StrDbpath = GetSetting("DataSearch", "JCI", "Path2")
    End If

    Cnnstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrDbpath
      'Cnnstr = "Driver={Microsoft Access Driver (*.mdb)};" & _
                        "Dbq=" & StrDbpath & ";" & _
                        "Uid=;Pwd=alpha;"

    intretval = OpenDatabaseAdo(cnn, Cnnstr)
    If intretval <> 0 Then
        MsgBox "Unable To Establish the Connection. ", vbCritical
        FrmSetDatabase.cmdOK.Enabled = False
        FrmSetDatabase.Show
        Exit Sub
    Else
        frmScreen.Show
    End If
End Sub
Public Sub ConnectDCPI()

Dim strDrv As String
'connect to database
    On Error GoTo Errh
    'strDrv = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=DCPI"
    strDrv = Cnnstr
    Set cnDCPI_User = New ADODB.Connection
    cnDCPI_User.CursorLocation = adUseClient
    cnDCPI_User.CommandTimeout = 60
    cnDCPI_User.ConnectionString = Cnnstr
    If cnDCPI_User.State = 1 Then cnDCPI_User.Close
    cnDCPI_User.Open
   
    Set cnDCPI = New ADODB.Connection
    cnDCPI.CursorLocation = adUseClient
    cnDCPI.CommandTimeout = 60
    cnDCPI.ConnectionString = Cnnstr
    cnDCPI.Open
    
    Exit Sub
Errh:
        MsgBox "Connection is not established. Pls contact your System Administrator!", vbCritical, CompanyName
        Set cnDCPI = Nothing
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
    'Dim strSQL As String
    Dim rectemp1 As ADODB.Recordset
On Error GoTo Err_Handler:
    strSQL = " select * " & _
             " From " & PStrtablename & ""
   ' Set rectemp1 = pdbconinput.Execute(strSQL)
    Set rectemp1 = New ADODB.Recordset
    rectemp1.Open strSQL, pdbconinput
         IsTableEXIST = Exists
        rectemp1.Close
    Set rectemp1 = Nothing
    Exit Function
Err_Handler:
    IsTableEXIST = notExists
 '  Call GetAPPError("IsExistsOrNotInTable")

End Function

Public Function NEWID() As String
Dim RsNewID As New ADODB.Recordset
 Set RsNewID = ReturnRst("Select max(format(mid(controlno,7,5),'00000')) from HREMPLOYEES", Cnnstr, mread)
           
If RsNewID.EOF = False Then
       If IsNull(RsNewID.Fields(0)) = False Then
               NEWID = Val(RsNewID.Fields(0))
               Else
               NEWID = 0
               End If
End If
End Function


Public Function ValidateButton(ByVal FormName As Form, Optional ByVal a As Boolean, Optional ByVal b As Boolean, Optional ByVal c As Boolean, Optional ByVal D As Boolean, Optional ByVal e As Boolean, Optional ByVal f As Boolean)

FormName.cmdadd.Enabled = a
FormName.cmdedit.Enabled = b
FormName.cmdsave.Enabled = c
FormName.cmdCancel.Enabled = D
FormName.cmdrefresh.Enabled = e
FormName.cmddelete.Enabled = f
End Function

Function getMaxWAGEid(ByVal TableName As String) As Integer
    Dim rs As New ADODB.Recordset
    Set rs = ReturnRst("Select MAX(ID) from " & TableName, Cnnstr, mstatic)
    If rs.EOF = False Then
'      Dim rs2 As New ADODB.Recordset
'      Set rs2 = ReturnRst("Select MAXWAGE from SSStable WHERE ID=" & rs(0).Value & "", Cnnstr, mstatic)
    getMaxWAGEid = rs(0)
    End If
End Function
 Public Function GETuSERNAME(ByVal USERID As Integer) As String
    Dim rsUNAME As New ADODB.Recordset
    Set rsUNAME = ReturnRst("Select USERNAME from USERSLIST WHERE OPERATORID=" & USERID, Cnnstr, mstatic)
    If rsUNAME.EOF = False Then
        GETuSERNAME = rsUNAME(0)
    Else
        GETuSERNAME = "SYSTEM"
    End If
 End Function
