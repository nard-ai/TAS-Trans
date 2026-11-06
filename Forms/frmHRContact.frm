VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmHRContact 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HR Contact"
   ClientHeight    =   2220
   ClientLeft      =   4590
   ClientTop       =   7530
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   9285
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1575
      Left            =   3360
      TabIndex        =   5
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employeenum"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Surname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Firstname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Middlename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Contact1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Contact2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Contact3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Contact4"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdselected 
         Caption         =   "&Export to Excel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "&Preview"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtemployeenum 
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin MSForms.CheckBox chkall 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   735
         BackColor       =   -2147483639
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1296;661"
         Value           =   "0"
         Caption         =   "All"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "Employee Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdall 
      Caption         =   "&All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmHRContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conallexcel As New Excel.Application
Dim conexcel As New Excel.Application

Private Sub cmdall_Click()
'Dim conallworkbook As New Excel.Workbook
'Dim conallworksheet As New Excel.Worksheet
'Dim rsconall As New ADODB.Recordset
'Dim strconall As String
'Dim stremp As String
'Dim x As Integer
'Dim i As Integer
'
'Set conallexcel = New Excel.Application
'Set conallworkbook = conallexcel.Workbooks.add()
'Set conallworksheet = conallexcel.Worksheets(1)
'
'ProgressBar1.Visible = True
'
'conallexcel.Cells(1, 1) = "Contact No."
'conallexcel.Cells(2, 1) = "Employee No."
'conallexcel.Cells(2, 2) = "Surname"
'conallexcel.Cells(2, 3) = "First Name"
'conallexcel.Cells(2, 4) = "Middle Name"
'conallexcel.Cells(2, 5) = "Contact 1"
'conallexcel.Cells(2, 6) = "Contact 2"
'conallexcel.Cells(2, 7) = "Contact 3"
'conallexcel.Cells(2, 8) = "Contact 4"
'conallexcel.Range("A1:IV1").EntireColumn.AutoFit
'conallexcel.Cells.Range("A1:E1").Merge
'conallexcel.Range("A1:IV1").Font.Bold = True
'conallexcel.Range("A2:IV2").Font.Bold = True
'x = 2
'stremp = txtemployeenum.Text
'strconall = "Select employeenum, surname, firstname, middlename, telnum1, telnum2, telnum3, telnum4 from HREmployAll"
'Set rsconall = cnDCPI.Execute(strconall)
'If rsconall.EOF = False Then
'    Do While Not rsconall.EOF
'    x = x + 1
'        For i = 1 To 8
'        ProgressBar1.Value = (((x - 2) * 10) / 100)
'            conallexcel.Cells(x, i).Value = rsconall.Fields(i - 1)
'        Next i
'    rsconall.MoveNext
'    Loop
'End If
'
'conallexcel.Visible = True
'ProgressBar1.Visible = False
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdpreview_Click()
Dim rscontact As New ADODB.Recordset
Dim strcontact As String
Dim strcon As String
Dim strtemp As String
Dim strres As String
Dim lstcontact As ListItem
Dim lstcontact1 As ListItem

strres = "Resigned"
If chkall.Value = True Then
    ListView1.ListItems.Clear
    strtemp = "Select employeenum, surname, firstname, middlename, telnum1, telnum2, telnum3, telnum4, employmentstatus from HREmployAll where employmentstatus <> '" & strres & "'"
    Set rscontact = cnDCPI.Execute(strtemp)
    Do While Not rscontact.EOF
        Set lstcontact1 = ListView1.ListItems.add(, , rscontact!employeenum)
        lstcontact1.SubItems(1) = IIf(IsNull(rscontact!surname), "", (rscontact!surname))
        lstcontact1.SubItems(2) = IIf(IsNull(rscontact!FirstName), "", (rscontact!FirstName))
        lstcontact1.SubItems(3) = IIf(IsNull(rscontact!middlename), "", (rscontact!middlename))
        lstcontact1.SubItems(4) = IIf(IsNull(rscontact!telnum1), "", (rscontact!telnum1))
        lstcontact1.SubItems(5) = IIf(IsNull(rscontact!telnum2), "", (rscontact!telnum2))
        lstcontact1.SubItems(6) = IIf(IsNull(rscontact!telnum3), "", (rscontact!telnum3))
        lstcontact1.SubItems(7) = IIf(IsNull(rscontact!telnum4), "", (rscontact!telnum4))
        rscontact.MoveNext
    Loop
    rscontact.Close
    Set rscontact = Nothing
    cmdselected.Enabled = True
Else
    strcon = txtemployeenum.Text
    If txtemployeenum.Text = "" Then
        MsgBox "Employee Number Cannot Be Blank. Please check.", vbCritical, "Search Asia"
        txtemployeenum.SetFocus
    ElseIf Not IsNumeric(txtemployeenum.Text) Then
        MsgBox "Invalid format of employee number. Please enter numbers only.", vbCritical, "Search Asia"
        txtemployeenum.SetFocus
    Else
        ListView1.ListItems.Clear
        strcontact = "Select employeenum, surname, firstname, middlename, telnum1, telnum2, telnum3, telnum4 from HREmployAll where employeenum = '" & strcon & "'"
        Set rscontact = cnDCPI.Execute(strcontact)
        If rscontact.EOF = False Then
            Do While Not rscontact.EOF
                Set lstcontact = ListView1.ListItems.add(, , rscontact!employeenum)
                lstcontact.SubItems(1) = IIf(IsNull(rscontact!surname), "", (rscontact!surname))
                lstcontact.SubItems(2) = IIf(IsNull(rscontact!FirstName), "", (rscontact!FirstName))
                lstcontact.SubItems(3) = IIf(IsNull(rscontact!middlename), "", (rscontact!middlename))
                lstcontact.SubItems(4) = IIf(IsNull(rscontact!telnum1), "", (rscontact!telnum1))
                lstcontact.SubItems(5) = IIf(IsNull(rscontact!telnum2), "", (rscontact!telnum2))
                lstcontact.SubItems(6) = IIf(IsNull(rscontact!telnum3), "", (rscontact!telnum3))
                lstcontact.SubItems(7) = IIf(IsNull(rscontact!telnum4), "", (rscontact!telnum4))
            rscontact.MoveNext
            Loop
            rscontact.Close
            Set rscontact = Nothing
        End If
    End If
    cmdselected.Enabled = True
End If
End Sub

Private Sub cmdselected_Click()
Dim conworkbook As New Excel.Workbook
Dim conworksheet As New Excel.Worksheet
Dim rscon As New ADODB.Recordset
Dim strcont As String
Dim strcont1 As String
Dim stremployeenum As String
Dim X As Integer
Dim i As Integer
Dim z As Integer

On Error GoTo Errh

    cmdpreview_Click
  
    Set conexcel = New Excel.Application
    Set conworkbook = conexcel.Workbooks.add()
    Set conworksheet = conexcel.Worksheets(1)
    
    ProgressBar1.Visible = True
    
    conexcel.Cells(1, 1) = "Contact No."
    conexcel.Cells(2, 1) = "Employee No. "
    conexcel.Cells(2, 2) = "Surname "
    conexcel.Cells(2, 3) = "First Name "
    conexcel.Cells(2, 4) = "Middle Name "
    conexcel.Cells(2, 5) = "Contact 1 "
    conexcel.Cells(2, 6) = "Contact 2 "
    conexcel.Cells(2, 7) = "Contact 3 "
    conexcel.Cells(2, 8) = "Contact 4 "
    conexcel.Cells(2, 9) = "Employment Status "
    conexcel.Range("A1:IV1").EntireColumn.AutoFit
    conexcel.Cells.Range("A1:E1").Merge
    conexcel.Range("A1:IV1").Font.Bold = True
    conexcel.Range("A2:IV2").Font.Bold = True
    
    X = 2
    z = 0
    strres = "Resigned"
    If chkall.Value = True Then
        strcont1 = "Select employeenum, surname, firstname, middlename, telnum1, telnum2, telnum3, telnum4, employmentstatus from HREmployAll where employmentstatus <> '" & strres & "'"
        Set rscon = cnDCPI.Execute(strcont1)
        Do While Not rscon.EOF
            X = X + 1
            z = z + 1
            ProgressBar1.Value = ((z / ListView1.ListItems.Count) * 100)
                For i = 1 To 9
                    conexcel.Cells(X, i).Value = rscon.Fields(i - 1)
                Next i
            rscon.MoveNext
        Loop
        X = X + 2
        conexcel.Cells(X, 1).Value = "Total Records:"
        conexcel.Cells(X, 1).Font.Bold = True
        conexcel.Cells(X, 2).Value = rscon.RecordCount & " " & "Records"
        conexcel.Cells(X, 2).Font.Bold = True
        rscon.Close
        Set rscon = Nothing
    Else
        stremployeenum = txtemployeenum.Text
        strcont = "Select employeenum, surname, firstname, middlename, telnum1, telnum2, telnum3, telnum4 from HREmployAll where employeenum = '" & stremployeenum & "'"
        Set rscon = cnDCPI.Execute(strcont)
        If rscon.EOF = False Then
            Do While Not rscon.EOF
            X = X + 1
            z = z + 1
            ProgressBar1.Value = ((z / ListView1.ListItems.Count) * 100)
                For i = 1 To 8
                    conexcel.Cells(X, i).Value = rscon.Fields(i - 1)
                Next i
            rscon.MoveNext
            Loop
            X = X + 2
            conexcel.Cells(X, 1).Value = "Total Records:"
            conexcel.Cells(X, 2).Value = ListView1.ListItems.Count & " " & "Records"
            rscon.Close
            Set rscon = Nothing
        End If
    End If

conexcel.Visible = True
ProgressBar1.Visible = False

Exit Sub
Errh:
    MsgBox "Cannot Create Excel Report", vbCritical, "Search Asia"
    Set conexcel = Nothing
End Sub

Private Sub Form_Load()
'cmdpreview.Enabled = False
cmdselected.Enabled = False
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ListView1.Sorted = True
ListView1.SortKey = ColumnHeader.Index - 1
ListView1.Refresh
End Sub

Private Sub txtemployeenum_GotFocus()
txtemployeenum.BackColor = &HFFFFC0
End Sub

Private Sub txtemployeenum_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub

Private Sub txtemployeenum_LostFocus()
txtemployeenum.BackColor = vbWhite
'cmdPreview.Enabled = True
'cmdPreview.SetFocus
End Sub
