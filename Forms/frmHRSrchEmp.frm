VERSION 5.00
Begin VB.Form frmHRSrchEmp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1335
   ScaleWidth      =   3210
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   1245
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1155
   End
   Begin VB.ComboBox cmbemp 
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Employee No."
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
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1110
   End
End
Attribute VB_Name = "frmHRSrchEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim searchempno As String
Dim res
Dim emp As String
Dim rsemp As ADODB.Recordset

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdsearch_Click()
    serch = True
    empno = Trim(Me.cmbemp)
    Unload Me
 
End Sub
Private Sub Form_Load()
   Me.Top = (mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
Set rsemp = New ADODB.Recordset
Set rsemp = cnDCPI.Execute("Select distinct employeenum From HREmployeesAttendance")
Do While Not rsemp.EOF
        cmbemp.AddItem rsemp!employeenum
        rsemp.MoveNext
    Loop
    rsemp.Close
cmbemp.ListIndex = -1
'empno = m_empno '.txtEmployeeNo.Text 'to be removed
End Sub
