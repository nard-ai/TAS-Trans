VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmHRCopy 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HR Search"
   ClientHeight    =   5805
   ClientLeft      =   4155
   ClientTop       =   3510
   ClientWidth     =   8940
   Icon            =   "frmHRCopy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   8940
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   6120
      TabIndex        =   20
      Top             =   1215
      Width           =   2760
      Begin VB.CheckBox Check1 
         Height          =   195
         Index           =   1
         Left            =   2475
         TabIndex        =   37
         Top             =   0
         Width           =   195
      End
      Begin VB.ComboBox cmbClient 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         ItemData        =   "frmHRCopy.frx":058A
         Left            =   135
         List            =   "frmHRCopy.frx":058C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   4050
         Width           =   2595
      End
      Begin VB.Frame FraSettings 
         Height          =   1500
         Index           =   1
         Left            =   90
         TabIndex        =   21
         Top             =   135
         Visible         =   0   'False
         Width           =   2595
         Begin VB.CheckBox chkStatus 
            Caption         =   "Resigned"
            Height          =   255
            Index           =   17
            Left            =   1440
            TabIndex        =   30
            Top             =   135
            Width           =   1005
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Early EOC"
            Height          =   255
            Index           =   16
            Left            =   1440
            TabIndex        =   29
            Top             =   405
            Width           =   1020
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "EOC"
            Height          =   255
            Index           =   15
            Left            =   1440
            TabIndex        =   28
            Top             =   675
            Width           =   960
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "AWOL"
            Height          =   255
            Index           =   14
            Left            =   1440
            TabIndex        =   27
            Top             =   945
            Width           =   960
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Contractual"
            Height          =   255
            Index           =   13
            Left            =   45
            TabIndex        =   26
            Top             =   405
            Value           =   1  'Checked
            Width           =   1230
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Project Base"
            Height          =   255
            Index           =   12
            Left            =   45
            TabIndex        =   25
            Top             =   135
            Width           =   1230
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Probationary"
            Height          =   255
            Index           =   11
            Left            =   45
            TabIndex        =   24
            Top             =   675
            Width           =   1185
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Regular"
            Height          =   255
            Index           =   10
            Left            =   45
            TabIndex        =   23
            Top             =   945
            Width           =   1230
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Terminated"
            Height          =   255
            Index           =   9
            Left            =   45
            TabIndex        =   22
            Top             =   1215
            Width           =   1230
         End
      End
      Begin MSComctlLib.ListView lvwdisplay 
         Height          =   3915
         Index           =   1
         Left            =   135
         TabIndex        =   31
         Top             =   135
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   6906
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NAME"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "SEX"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ADDRESS"
            Object.Width           =   4586
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "count"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2925
         TabIndex        =   32
         Top             =   4050
         Width           =   1200
      End
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   6615
      TabIndex        =   19
      Top             =   810
      Width           =   2265
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
      Height          =   375
      Left            =   4230
      TabIndex        =   18
      Top             =   2835
      Width           =   690
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   375
      Left            =   4230
      TabIndex        =   17
      Top             =   2340
      Width           =   690
   End
   Begin VB.CheckBox Check1 
      Height          =   240
      Index           =   0
      Left            =   2520
      TabIndex        =   5
      Top             =   1170
      Width           =   195
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   540
      TabIndex        =   3
      Top             =   855
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   45
      TabIndex        =   0
      Top             =   1215
      Width           =   2760
      Begin VB.ComboBox cmbClient 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         ItemData        =   "frmHRCopy.frx":058E
         Left            =   45
         List            =   "frmHRCopy.frx":0590
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   4050
         Width           =   2595
      End
      Begin VB.Frame FraSettings 
         Height          =   1500
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   135
         Visible         =   0   'False
         Width           =   2595
         Begin VB.CheckBox chkStatus 
            Caption         =   "Terminated"
            Height          =   255
            Index           =   4
            Left            =   45
            TabIndex        =   15
            Top             =   1215
            Width           =   1230
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Regular"
            Height          =   255
            Index           =   3
            Left            =   45
            TabIndex        =   14
            Top             =   945
            Width           =   1230
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Probationary"
            Height          =   255
            Index           =   2
            Left            =   45
            TabIndex        =   13
            Top             =   675
            Width           =   1185
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Project Base"
            Height          =   255
            Index           =   0
            Left            =   45
            TabIndex        =   12
            Top             =   135
            Width           =   1230
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Contractual"
            Height          =   255
            Index           =   1
            Left            =   45
            TabIndex        =   11
            Top             =   405
            Value           =   1  'Checked
            Width           =   1230
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "AWOL"
            Height          =   255
            Index           =   8
            Left            =   1440
            TabIndex        =   10
            Top             =   945
            Width           =   960
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "EOC"
            Height          =   255
            Index           =   7
            Left            =   1440
            TabIndex        =   9
            Top             =   675
            Width           =   960
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Early EOC"
            Height          =   255
            Index           =   6
            Left            =   1440
            TabIndex        =   8
            Top             =   405
            Width           =   1020
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Resigned"
            Height          =   255
            Index           =   5
            Left            =   1440
            TabIndex        =   7
            Top             =   135
            Width           =   1005
         End
      End
      Begin MSComctlLib.ListView lvwdisplay 
         Height          =   3915
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   6906
         View            =   3
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NAME"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "SEX"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ADDRESS"
            Object.Width           =   4586
         EndProperty
      End
      Begin VB.Label lblRecount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "count"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2925
         TabIndex        =   2
         Top             =   4050
         Width           =   1200
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Search Asia Inc."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   6120
      TabIndex        =   34
      Top             =   405
      Width           =   2715
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   33
      Top             =   900
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jobs and Careers Inc."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   45
      TabIndex        =   16
      Top             =   360
      Width           =   2715
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   855
      Width           =   750
   End
End
Attribute VB_Name = "frmHRCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsClient As New ADODB.Recordset
Dim cnnJCI As ADODB.Connection
Dim cnnSAPI As ADODB.Connection
Dim searchempno As String
Dim res
Dim SelectString As String
Dim rs As ADODB.Recordset



Private Sub Check1_Click(Index As Integer)

   If FraSettings(Index).Visible = True Then
        FraSettings(Index).Visible = False
        lvwdisplay(Index).Height = Me.Height - 2420
        lvwdisplay(Index).Top = 225
    Else
        FraSettings(Index).Visible = True
        lvwdisplay(Index).Height = Me.Height - 3760
        lvwdisplay(Index).Top = 1600
    End If

End Sub

Public Sub chkStatus_Click(Index As Integer)


Dim str As String
str = ""

If chkStatus(0).Value = 1 Then
str = "'B'"
End If

If chkStatus(1).Value = 1 Then
If str <> "" Then str = str & ",'C'" Else str = "'C'"
End If

If chkStatus(2).Value = 1 Then
If str <> "" Then str = str & ",'P'" Else str = "'P'"
End If

If chkStatus(3).Value = 1 Then
If str <> "" Then str = str & ",'R'" Else str = "'R'"
End If

If chkStatus(4).Value = 1 Then
If str <> "" Then str = str & ",'T'" Else str = "'T'"
End If


If chkStatus(5).Value = 1 Then
If str <> "" Then str = str & ",'X'" Else str = "'X'"
End If

If chkStatus(6).Value = 1 Then
If str <> "" Then str = str & ",'Y'" Else str = "'Y'"
End If

If chkStatus(7).Value = 1 Then
    If str <> "" Then str = str & ",'Z'" Else str = "'Z'"
End If

If chkStatus(8).Value = 1 Then
    If str <> "" Then str = str & ",'W'" Else str = "'W'"
End If


If str = "" Then MsgBox "Pls check at least 1 check box.", vbInformation:  chkStatus(1).Value = 1: Exit Sub
If whereCon <> "" Then
    whereCon = whereCon & " AND EMPLOYMENTSTATUS IN " & "(" & str & ")"
Else
    whereCon = " EMPLOYMENTSTATUS IN " & "(" & str & ")"
End If

 If cmbClient(Index).Text <> "--ALL COMPANIES--" Then
        whereCon = whereCon & "and clientcode=" & Me.cmbClient(Index).ItemData(Me.cmbClient(Index).ListIndex)
 End If
       
If whereCon <> "" Then whereCon = " where " & whereCon
 
 If Index = 0 Then
 Call Loadlist(whereCon, Index, cnnJCI)
 Else
 Call Loadlist(whereCon, Index, cnnSAPI)
End If

End Sub








Private Sub cmbClient_Click(Index As Integer)
        If Me.cmbClient(Index).Text = "--ALL COMPANIES--" Then
        Filter_var = ""
        Else
             Filter_var = " WHERE clientcode=" & Me.cmbClient(Index).ItemData(Me.cmbClient(Index).ListIndex)
       End If
   Call frmHRCopy.chkStatus_Click(Index)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub


Private Sub JCI()
    Dim intretval   As Integer
        StrDbpath = GetSetting("DataSearch", "JCI", "Path2")
    CnnstrJCI = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrDbpath
    intretval = OpenDatabaseAdo(cnnJCI, CnnstrJCI)
    If intretval <> 0 Then
        MsgBox "Unable To Establish the Connection to JCI. ", vbCritical
    Else
        Set rstclient = cnnJCI.Execute("Select * from clienttable")
        Do While Not rstclient.EOF
        DoEvents
            cmbClient(0).AddItem rstclient!ClientName
            cmbClient(0).ItemData(cmbClient(0).NewIndex) = rstclient!ClientCode
            rstclient.MoveNext
        Loop
        rstclient.Close
        cmbClient(0).AddItem "--ALL COMPANIES--", 0
'        cmbClient(1).ListIndex = 0

    End If
End Sub

Private Sub SAPI()
    Dim intretval   As Integer
    StrDbpath = GetSetting("DataSearch", "SE", "Path2")
    CnnstrSAPI = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrDbpath
    intretval = OpenDatabaseAdo(cnnSAPI, CnnstrSAPI)
    If intretval <> 0 Then
        MsgBox "Unable To Establish the Connection. ", vbCritical
Else
    Set rstclient = cnnSAPI.Execute("Select * from clienttable")
        Do While Not rstclient.EOF
        DoEvents
            cmbClient(1).AddItem rstclient!ClientName
            cmbClient(1).ItemData(cmbClient(1).NewIndex) = rstclient!ClientCode
            rstclient.MoveNext
        Loop
        rstclient.Close
        cmbClient(1).AddItem "--ALL COMPANIES--", 0
'        cmbClient(0).ListIndex = 0
        

     End If
End Sub


Private Sub Form_Load()
    
    Me.Top = (mdiMain.Height - Me.Height) / 4
    Me.Left = (mdiMain.Width - Me.Width) / 2
    
    SAPI
    JCI
        Call Loadlist("", 0, cnnSAPI)
        Call Loadlist("", 1, cnnJCI)
End Sub
Public Sub Loadlist(ByVal whereCon As String, ByVal NDX As Integer, ByVal strConn As ADODB.Connection)
    lvwdisplay(NDX).ListItems.Clear
    Dim rsEmp1 As ADODB.Recordset
    Set rsEmp1 = New ADODB.Recordset
    With lvwdisplay(NDX)
    rsEmp1.Open "select surname + ', ' & + firstname + ' ' & IIf(isnull(middlename) or middlename='','',Left(middlename,1)+'.') as NAME,EMPLOYEENUM,CONTROLNO AS [CONTROL NO] from HREMPLOYEES " & whereCon & " order by surname, firstname", strConn, adOpenDynamic, adLockOptimistic
    PopulateList Me.lvwdisplay(NDX), rsEmp1
    lvwdisplay(NDX).ColumnHeaders(1).Text = "  " & .ListItems.Count & " RECORD/S "
    lvwdisplay(NDX).ColumnHeaders(1).Width = 2300
    lvwdisplay(NDX).ColumnHeaders(2).Width = 0
    lvwdisplay(NDX).ColumnHeaders(3).Width = 0
    lblRecount = .ListItems.Count & " record/s "
    End With
End Sub

Private Sub lvwdisplay_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
txtSearch(0) = lvwdisplay(0).ListItems.Item(lvwdisplay(0).SelectedItem.Index).Text
End Sub

'Private Sub lvwdisplay_DblClick()
'    strrow = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).SubItems(1)
'    txtSearch = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).Text
'    If lvwdisplay.ListItems.Count = 0 Then
'        MsgBox "List doesn't have any record from the table selected !", vbOKOnly + vbInformation, "Select..."
'    Exit Sub
'    Else
'    End If
'    Call search(strrow)
'End Sub
'
'Private Sub lvwdisplay_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    txtSearch(0) = lvwdisplay(0).ListItems.Item(lvwdisplay(0).SelectedItem.Index).Text
'End Sub

'Private Sub lvwdisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'If FeMP = True Then
'Call lvwdisplay_DblClick
'End If
'End Sub
'
'Private Sub lvwdisplay_KeyUp(KeyCode As Integer, Shift As Integer)
'If FeMP = True Then
'Call lvwdisplay_DblClick
'End If
'End Sub

Private Sub txtSearch_Change(Index As Integer)
 Autocomplete lvwdisplay(0), (Trim$(txtSearch(0))), txtSearch(0)
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
    Call txtHLigths(txtSearch(0))
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Execute = True
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
    Execute = False
    End If
End Sub

'Private Sub txtSearch_KeyPress(KeyAscii As Integer)
'strrow = ""
' KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
'    If KeyAscii = 13 Then
'    If lvwdisplay.ListItems.Count > 0 Then
'     strrow = lvwdisplay.ListItems.Item(lvwdisplay.SelectedItem.Index).SubItems(1)
'    Call search(strrow)
'    Else
'    MsgBox "No records in list.", vbInformation
'    End If
'     End If
'
'End Sub

Private Sub search(ByVal strSearch As Integer)
On Error Resume Next
mmempTemp = MMEmp
'If MMEmp = False And FeMP = False Then MMEmp = True
If MMEmp = False And FeMP = False Then FeMP = True
Dim res, bkmrk
If FeMP = True Then
With frmHRFinalEmployees
   Set .rsout = New ADODB.Recordset
    .rsout.Open "Select * from HREmployees order by  surname", cnDCPI, adOpenKeyset, adLockOptimistic

    bkmrk = .rsout.Bookmark
    .rsout.MoveFirst
    .rsout.Find "Employeenum=" & Trim(strSearch) & ""
    If Not .rsout.EOF Then
        .Showdata .rsout
    Else
        .rsout.Bookmark = bkmrk
    End If
End With
If MMEmp = False Then MMEmp = mmempTemp
End If
    
If MMEmp = True Then

Dim rsout As New ADODB.Recordset
Set frmHRMainEmployee.rs = New ADODB.Recordset
    If frmHRMainEmployee.rs.State = 1 Then rs.Close
    

With frmHRMainEmployee
    .rs.Open "Select * from HREmployees order by surname, firstname", cnDCPI, adOpenKeyset, adLockOptimistic
        
              
    .rs.Find "Employeenum=" & Trim(strSearch) & ""
    If Not .rs.EOF Then
        .Showdata .rs
    Else
        .rs.Bookmark = bkmrk
    End If
End With
End If
    
    
  '  Unload Me

End Sub
Private Sub Showdata(ByRef rec As ADODB.Recordset)
On Error Resume Next
Dim rs As ADODB.Recordset
Dim rsout As ADODB.Recordset
Dim str As String
     frmHRMainEmployee.TxtEmployeeNum.Text = IIf(IsNull(rec!employeenum), "", rec!employeenum)
    frmHRMainEmployee.txtEmpName.Text = rec!Surname & ", " & rec!FirstName
     frmHRMainEmployee.txtClient.Text = IIf(IsNull(rec!ClientCode), "", rec!ClientCode)
    frmHRMainEmployee.mskDtHired.Text = IIf(IsNull(rec!datehired), "__/__/____", rec!datehired)


   ' Info rec!employeenum
 
     Set rs = Nothing
     Set rsM = Nothing
     Set rsA = Nothing
     Set rst = Nothing
     Set rsO = Nothing
     Set rsL = Nothing
     Set rsD = Nothing
     Set rsB = Nothing
     Set rsP = Nothing
     
End Sub
