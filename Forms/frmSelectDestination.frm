VERSION 5.00
Begin VB.Form frmHRSelectDestination 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recontract to :"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.OptionButton OptCSPI 
      Caption         =   "Corporate Staffers Placement Inc."
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.OptionButton OptJCI 
      Caption         =   "Jobs and Placement Inc."
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.OptionButton OptSAPI 
      Caption         =   "Search Asia Placement Inc."
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
   End
End
Attribute VB_Name = "frmHRSelectDestination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctlOpt(10) As Control
Dim ctlFrm(10) As Object
Dim i As Integer
Dim intCompanyCount As Integer


Private Sub Form_Load()
   
    Dim rsemp As New ADODB.Recordset
  
    strSQL = "Select * from companytable where agencystatus = 1"
    Set rsemp = ReturnRst(strSQL, CnnstrHR, mstatic)
  
    intCompanyCount = rsemp.RecordCount
  
    Me.Height = 1800 + (intCompanyCount * 400)
    Me.cmdcancel.Top = 900 + (intCompanyCount * 400)
    Me.cmdOK.Top = 900 + (intCompanyCount * 400)
  

    For i = 1 To rsemp.RecordCount
        'Set ctlOpt(i) = Form1.Controls.Add("VB.OptionButton", "Opt" & i, ctlFrm(1))
        Set ctlOpt(i) = Me.Controls.Add("VB.OptionButton", "Opt" & rsemp.Fields!AgencyCode)
        With ctlOpt(i)
            If i = 1 Then
            .Top = 100
        Else
            .Top = ctlOpt(i - 1).Top + 500
        End If
        
        .Left = 1000
        .Width = 6000
        .Caption = getAgencyNamebyID(rsemp.Fields!AgencyCode) '.Name
        .Visible = True
       
        End With
  
        If rsemp.Fields!AgencyCode = CompanyCode Then
           ctlOpt(i).Enabled = False
           Else
             ctlOpt(i).Enabled = True
        End If
        If intCompanyCount = 1 Then
        ctlOpt(i).Enabled = False
        End If
         If rsemp.EOF = False Then rsemp.MoveNext
    
    Next

 If intCompanyCount > 1 Then
 Else
     MsgBox ("You cannot recontract this employee, only 1 agency is active.")
     Me.cmdOK.Enabled = False
 End If
 

'    If CompanyCode = 1 Then
'        OptSAPI.Enabled = False
'    ElseIf CompanyCode = 2 Then
'        OptJCI.Enabled = False
'    ElseIf CompanyCode = 3 Then
'        OptCSPI.Enabled = False
'    End If
End Sub

Private Sub cmdCancel_Click()
cancelRecontract = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    For i = 1 To intCompanyCount
    If ctlOpt(i).VALUE = True And ctlOpt(i).Enabled = True Then
    destinationCompanyCode = Replace(ctlOpt(i).Name, "Opt", "")
    End If
    
    
    
    Next
    
'    Dim CName As String
'    If Me.OptSAPI.VALUE = True Then
'        destinationCompanyCode = 1
'        StrDbpath = GetSetting("DataSearch", "SAPI", "Path2Version2")
'        CName = "SAPI Manpower & Mgt. Services"
'    ElseIf Me.OptJCI.VALUE = True Then
'        destinationCompanyCode = 2
'        CName = "Jobs and Careers Inc."
'        StrDbpath = GetSetting("DataSearch", "JCI", "Path2Version2")
'    ElseIf Me.OptCSPI.VALUE = True Then
'        destinationCompanyCode = 3
'        CName = "Corporate Staffers Placement Inc."
'        StrDbpath = GetSetting("DataSearch", "CSPI", "Path2Version2")
'    End If
    cancelRecontract = False

    frmHRFinalEmployees.cmbCompany.Clear
    If destinationCompanyCode = "" Or IsNull(destinationCompanyCode) Then
        MsgBox "No Destination Company specified in Company Table! Please update first.", vbInformation
        GoTo ExitRtn2
    Else
        frmHRFinalEmployees.LoadCompany (destinationCompanyCode)
        frmHRFinalEmployees.TxtControlNo.Text = "<auto>"
    End If
    Unload Me
    
    Exit Sub
ExitRtn2:

'
'    CnnstrCopyDestination = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrDbpath
'
'     If StrDbpath = "" Then MsgBox "Unable to connect to " & CName & ". Run the system for this company to establish the connection.": Exit Sub
'     If MsgBox("Are you sure you want to move [" & frmHRFinalEmployees.txtsurname.Text & "," & frmHRFinalEmployees.txtfirstname.Text & "] to " & CName & "?", vbYesNo + vbQuestion) = vbNo Then
'         Exit Sub
'     End If
'
'
'     intretval = OpenDatabaseAdo(cnn, Cnnstr)
'     If intretval <> 0 Then
'         MsgBox "Unable To Establish the Connection. ", vbCritical
'     Exit Sub
'     End If
'
'     'Note we should use the firstname and lastname since control number is not reliable because it is editable by user.
'     strSQL = "Select * from hremployees where firstname ='" & LTrim(RTrim(frmHRFinalEmployees.txtfirstname.Text)) & "'  and surname= '" & LTrim(RTrim(frmHRFinalEmployees.txtsurname.Text)) & "'"
'     Dim rsemp As New ADODB.Recordset
'     Set rsemp = ReturnRst(strSQL, CnnstrCopyDestination, mstatic)
'     If rsemp.EOF = False Then
'         If MsgBox("CANNOT TRANFER THIS EMPLOYEE because the name [" & frmHRFinalEmployees.txtsurname.Text & "," & frmHRFinalEmployees.txtfirstname.Text & "]" & " is ALREADY EXIST in " & CName & " database." & vbCrLf & "" & vbCrLf & "Do you want to update this record in " & CName & "", vbYesNo + vbQuestion) = vbYes Then
'           'updates data
'           Call frmHRFinalEmployees.SaveHistory(True)
'         End If
'     Else
'         'inserts data
'            Call frmHRFinalEmployees.SaveHistory(False)
'            Call frmHRFinalEmployees.SaveOtherInfo(CnnstrCopyDestination)
'            MsgBox "[" & frmHRFinalEmployees.txtsurname.Text & "," & frmHRFinalEmployees.txtfirstname.Text & "] successfully moved to " & CName & " " & vbCrLf & " " & vbCrLf & "Pls. update the control no., client name and contract date of this employee.", vbInformation
'    End If
'        'refresh HR form
'        frmHRFinalEmployees.btnEmpRefresh_Click
'        Unload Me
'      Exit Sub
'bb:
'      MsgBox "Unable to tranfer this employee... Pls check the employee's information!", vbInformation
'
End Sub


