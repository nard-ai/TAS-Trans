VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHRViewAttendance 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Attendance"
   ClientHeight    =   5895
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnRefresh 
      Caption         =   "&Refresh"
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
      Left            =   8280
      Picture         =   "frmHRViewAttendance.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Refresh"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton btnViewSearch 
      Caption         =   "Search"
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
      Index           =   1
      Left            =   1680
      Picture         =   "frmHRViewAttendance.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Search"
      Top             =   120
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton btnViewExit 
      Caption         =   "Close"
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
      Index           =   3
      Left            =   10560
      Picture         =   "frmHRViewAttendance.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Exit"
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cmbDept 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   5400
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPCalendar 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      ToolTipText     =   "Calendar"
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   48103425
      CurrentDate     =   37533
   End
   Begin VB.CommandButton btnViewEdit 
      Caption         =   "Edit"
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
      Index           =   1
      Left            =   720
      Picture         =   "frmHRViewAttendance.frx":0396
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Edit"
      Top             =   120
      Width           =   855
   End
   Begin MSComctlLib.ListView lvwAttendance 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   12303
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employee No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Shift"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Choronolog ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Unit ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Attendance Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Day"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Attendance Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Log Mode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Extracted Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Dept Code"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmHRViewAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsHRAtndc As ADODB.Recordset
Dim rsdept As ADODB.Recordset
Dim HRList As ListItem
Dim dept As String
Dim atndcdt As Date

Private Sub btnRefresh_Click()
 Dim str As String

    Set rsHRAtndc = New ADODB.Recordset
    If dept = 9 Then
        str = "SELECT * From HREmployeesAttendance where attendancedate=#" & atndcdt & "#"
    Else
        str = "SELECT * From HREmployeesAttendance where attendancedate=#" & atndcdt & "# and deptcode =" & Format(dept, "00000")
    End If
    Set rsHRAtndc = cnDCPI.Execute(str)
    lvwAttendance.ListItems.Clear
    Do While Not rsHRAtndc.EOF
        ListAttendance
    rsHRAtndc.MoveNext
    Loop
    rsHRAtndc.Close
End Sub

Private Sub btnViewEdit_Click(Index As Integer)
    frmHRAttendance.Show vbModal
End Sub

Private Sub btnViewExit_Click(Index As Integer)
    Unload Me
End Sub

Private Sub btnViewSearch_Click(Index As Integer)
frmHRSrchEmp.Show vbModal
serch = True
LoadSearch
End Sub

Private Sub cmbDept_Click()
    dept = Me.cmbDept.ItemData(Me.cmbDept.ListIndex)
    
End Sub

Private Sub DTPCalendar_Change()
    atndcdt = DTPCalendar.Value
    
   
End Sub

Private Sub Form_Load()
    Set rsHRAtndc = New ADODB.Recordset
    Set rsHRAtndc = cnDCPI.Execute("Select * From HREmployeesAttendance")
            
    Set rsdept = cnDCPI.Execute("Select * From DCDepartment")
    Do While Not rsdept.EOF
        'cmbDept.AddItem rsdept!DeptCode & "-" & rsdept!Department
        cmbDept.AddItem rsdept!Department
        cmbDept.ItemData(cmbDept.NewIndex) = rsdept!DeptCode
        rsdept.MoveNext
    Loop
    rsdept.Close
    
    
    lvwAttendance.ListItems.Clear
    Do While Not rsHRAtndc.EOF
        ListAttendance
    rsHRAtndc.MoveNext
    Loop
    
    'to get latest date that extracted
    If rsHRAtndc.EOF Then
        rsHRAtndc.MoveFirst
        atndcdt = rsHRAtndc!attendancedate
    End If
    
    rsHRAtndc.Close
    
    
    DTPCalendar.Value = atndcdt
    cmbDept.Text = "*All"
        
        
End Sub

Private Sub lvwAttendance_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim str As String
    
    str = "Select * From HREmployeesAttendance where employeenum=" & Item.Text
    
    m_empno = Item.Text
    atndctym = list.SubItems(6)
    
End Sub

Private Sub Timer1_Timer()
    Me.Caption = "View Attendance    " & Now
End Sub

Private Sub ListAttendance()
    Set list = lvwAttendance.ListItems.add(, , rsHRAtndc!employeenum)
        list.SubItems(1) = IIf(IsNull(rsHRAtndc!shiftcode), 0, rsHRAtndc!shiftcode)
        list.SubItems(2) = IIf(IsNull(rsHRAtndc!chronoid), 0, rsHRAtndc!chronoid)
        list.SubItems(3) = IIf(IsNull(rsHRAtndc!unitid), 0, rsHRAtndc!unitid)
        list.SubItems(4) = IIf(IsNull(rsHRAtndc!attendancedate), "", rsHRAtndc!attendancedate)
        list.SubItems(5) = IIf(IsNull(rsHRAtndc!Day), 0, rsHRAtndc!Day)
        list.SubItems(6) = IIf(IsNull(rsHRAtndc!attendancetime), "", rsHRAtndc!attendancetime)
        list.SubItems(7) = IIf(IsNull(rsHRAtndc!LogMode), "", rsHRAtndc!LogMode)
        list.SubItems(8) = IIf(IsNull(rsHRAtndc!FileName), "", rsHRAtndc!FileName)
        list.SubItems(9) = IIf(IsNull(rsHRAtndc!extracteddate), "", rsHRAtndc!extracteddate)
        list.SubItems(10) = IIf(IsNull(rsHRAtndc!DeptCode), "", rsHRAtndc!DeptCode)
End Sub


Private Sub LoadSearch()
    Dim str As String
    str = "Select * From HREmployeesAttendance where employeenum='" & empno & "'"
    
    Set rsHRAtndc = cnDCPI.Execute(str)
    lvwAttendance.ListItems.Clear
    Do While Not rsHRAtndc.EOF
        ListAttendance
    rsHRAtndc.MoveNext
    Loop
    rsHRAtndc.Close
    serch = False
End Sub
