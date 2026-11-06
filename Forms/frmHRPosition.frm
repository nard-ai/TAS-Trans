VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmHRPosition 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Position"
   ClientHeight    =   5760
   ClientLeft      =   6000
   ClientTop       =   4710
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   7530
   Begin MSComctlLib.ListView ListView3 
      Height          =   2055
      Left            =   0
      TabIndex        =   11
      Top             =   3240
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3625
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
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
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
         Text            =   "Position"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Employmentstatus"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   5400
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   5760
      TabIndex        =   7
      Top             =   360
      Width           =   1695
      Begin VB.CommandButton cmdpreview 
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
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdexport 
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
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
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
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Caption         =   "Employment Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   2880
      TabIndex        =   6
      Top             =   120
      Width           =   2775
      Begin VB.ListBox lststatus 
         Enabled         =   0   'False
         Height          =   1035
         Left            =   120
         TabIndex        =   9
         Top             =   3120
         Width           =   2535
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2655
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "EmploymentStatus"
            Object.Width           =   4410
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.ListBox lstposition 
         Enabled         =   0   'False
         Height          =   2400
         Left            =   120
         TabIndex        =   8
         Top             =   3120
         Width           =   2535
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Position"
            Object.Width           =   4057
         EndProperty
      End
   End
End
Attribute VB_Name = "frmHRPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim posexcel As New Excel.Application

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdexport_Click()
Dim posworkbook As New Excel.Workbook
Dim posworksheet As New Excel.Worksheet
Dim rspos As New ADODB.Recordset
Dim i As Integer
Dim X As Integer
Dim Y As Integer
Dim z As Integer
Dim m As Integer
Dim strposition As String
Dim strstatus As String
Dim strpos As String

On Error GoTo Errh

If ListView3.ListItems.Count = 0 Then
    MsgBox "Please click preview", vbCritical, "Search Asia"
Else
    Set posexcel = New Excel.Application
    Set posworkbook = posexcel.Workbooks.add()
    Set posworksheet = posexcel.Worksheets(1)
    
    ProgressBar1.Visible = True
    
    posexcel.Cells(1, 1) = "Employee's Position and Employment Status"
    posexcel.Cells(2, 1) = "Employee No."
    posexcel.Cells(2, 2) = "Surname"
    posexcel.Cells(2, 3) = "First Name"
    posexcel.Cells(2, 4) = "Middle Name"
    posexcel.Cells(2, 5) = "Position"
    posexcel.Cells(2, 6) = "Employment Status"
    posexcel.Range("A1:IV1").EntireColumn.AutoFit
    posexcel.Cells.Range("A1:E1").Merge
    posexcel.Range("A1:IV1").Font.Bold = True
    posexcel.Range("A2:IV2").Font.Bold = True
    
    X = 2
    m = 0
    For z = 0 To lststatus.ListCount
        strstatus = lststatus.list(z)
        For Y = 0 To lstposition.ListCount
            'strpos = ListView1.ListItems(y).Checked
            strpos = lstposition.list(Y)
            strposition = "Select employeenum, surname, firstname, middlename, position, employmentstatus from HREmployAll where position = '" & strpos & "' And employmentstatus = '" & strstatus & "'"
            Set rspos = cnDCPI.Execute(strposition)
            Do While Not rspos.EOF
                X = X + 1
                m = m + 1
                ProgressBar1.Value = ((m / ListView3.ListItems.Count) * 100)
                For i = 1 To 6
                    posexcel.Cells(X, i).Value = rspos.Fields(i - 1)
                Next i
                rspos.MoveNext
            Loop
        Next Y
        'x = x + 1
        'posexcel.Cells(x, 1).Value = "Total Records:"
        'posexcel.Cells(x, 2).Value = rspos.RecordCount & " " & "Records"
    Next z
    X = X + 2
    posexcel.Cells(X, 1).Value = "Total Records:"
    posexcel.Cells(X, 1).Font.Bold = True
    posexcel.Cells(X, 2).Value = ListView3.ListItems.Count & " " & "Records"
    posexcel.Cells(X, 2).Font.Bold = True
    posexcel.Visible = True
    ProgressBar1.Visible = False
End If
'x = 2
''y = 1
'For z = 1 To ListView1.ListItems.Count
'    For y = 1 To ListView1.ListItems.Count
'        strpos = ListView2.ListItems(y).Checked
'        strposition = "Select employeenum, surname, firstname, middlename, position, employmentstatus from hremployall where position = '" & strpos & "' " And " employmentstatus = '" & ListView2.ListItems(z).Text & "'"
'        Set rspos = cnDCPI.Execute(strposition)
'        Do While Not rspos.EOF
'            x = x + 1
'            For i = 1 To 6
'                posexcel.Cells(x, i).Value = rspos.Fields(i - 1)
'            Next i
'            rspos.MoveNext
'        Loop
'    Next y
'Next z
Exit Sub
Errh:
    MsgBox "Cannot Create Excel Report", vbCritical, "Search Asia"
    Set posexcel = Nothing
    
End Sub

Private Sub cmdpreview_Click()
Dim rsTemp As New ADODB.Recordset
Dim strtemp As String
Dim strpos1 As String
Dim strstat1 As String
Dim lsttemp As ListItem
Dim z As Integer
Dim m As Integer
If lstposition.ListCount = 0 Then
    MsgBox "Please choose position.", vbCritical, "Search Asia"
ElseIf lststatus.ListCount = 0 Then
    MsgBox "Please choose status.", vbCritical, "Search Asia"
Else
    ListView3.ListItems.Clear
    For z = 0 To lstposition.ListCount - 1
    strpos1 = lstposition.list(z)
        For m = 0 To lststatus.ListCount - 1
        strstat1 = lststatus.list(m)
            strtemp = "Select employeenum, surname, firstname, middlename, position, employmentstatus from HREmployAll where position = '" & strpos1 & "' and employmentstatus = '" & strstat1 & "'"
            Set rsTemp = cnDCPI.Execute(strtemp)
            Do While Not rsTemp.EOF
            Set lsttemp = ListView3.ListItems.add(, , rsTemp!employeenum)
                lsttemp.SubItems(1) = IIf(IsNull(rsTemp!surname), "", rsTemp!surname)
                lsttemp.SubItems(2) = IIf(IsNull(rsTemp!FirstName), "", rsTemp!FirstName)
                lsttemp.SubItems(3) = IIf(IsNull(rsTemp!middlename), "", rsTemp!middlename)
                lsttemp.SubItems(4) = IIf(IsNull(rsTemp!Position), "", rsTemp!Position)
                lsttemp.SubItems(5) = IIf(IsNull(rsTemp!employmentstatus), "", rsTemp!employmentstatus)
            rsTemp.MoveNext
            Loop
            rsTemp.Close
            Set rsTemp = Nothing
        Next m
    Next z
    cmdexport.Enabled = True
End If
End Sub

Private Sub Form_Load()
   Me.Top = 0 '(mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
    
Dim rsposition As New ADODB.Recordset
Dim rsstatus As New ADODB.Recordset
Dim strposition As String
Dim strstatus As String
Dim lstposition As ListItem
Dim lststatus As ListItem

strposition = "Select distinct position from HREmployAll"
Set rsposition = cnDCPI.Execute(strposition)

Do While Not rsposition.EOF
    If (IsNull(rsposition.Fields(0)) = False) And (rsposition.Fields(0) <> "") Then
        Set lstpostion = ListView1.ListItems.add(, , rsposition.Fields(0))
    End If
    rsposition.MoveNext
Loop

strstatus = "Select distinct employmentstatus from HREmployAll"
Set rsstatus = cnDCPI.Execute(strstatus)

Do While Not rsstatus.EOF
    If (IsNull(rsstatus.Fields(0)) = False) And (rsstatus.Fields(0) <> "") Then
        Set lststatus = ListView2.ListItems.add(, , rsstatus.Fields(0))
    End If
    rsstatus.MoveNext
Loop
cmdexport.Enabled = False
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i As Long
Dim lst As ListItem

lstposition.Clear
For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(i).Checked = True Then
        lstposition.AddItem ListView1.ListItems(i).Text
        'ListView3.ListItems.add , , ListView1.ListItems(i).Text
    End If
Next i

'lstGroceries.AddItem cboProducts.Text
'List1.AddItem List2.Text
'List2.RemoveItem List2.ListIndex

'Dim lst As ListItem
'Dim str As String
'Dim i As Long
'
'If Item.Checked = True Then
'    ListView3.ListItems.add , , ListView1.ListItems(i).Text
'Else
'    For i = 1 To ListView3.ListItems.Count
'        If InStr(ListView3.ListItems(i).Text, Item.Text) <> 0 Then
'            ListView3.ListItems.Remove (i)
'        End If
'    Next i
'End If
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim X As Long
Dim lst1 As ListItem
Dim i As Long

lststatus.Clear
For X = 1 To ListView2.ListItems.Count
    If ListView2.ListItems(X).Checked = True Then
        lststatus.AddItem ListView2.ListItems(X).Text
        'ListView3.ListItems.add , , ListView2.ListItems(x).Text
    End If
Next X


'ListView3.ListItems.Clear
'For x = 1 To ListView2.ListItems.Count
'    If ListView2.ListItems(x).Checked = False Then
'        ListView3.ListItems.add , , ListView2.ListItems(x).Text
'    Else
'        For i = 1 To ListView3.ListItems.Count
'            ListView3.ListItems.Remove (i)
'        Next i
'    End If
'Next x
End Sub
