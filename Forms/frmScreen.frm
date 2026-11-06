VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScreen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3585
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5430
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4920
      Top             =   1080
   End
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   225
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Image CSPI 
      Height          =   1695
      Left            =   960
      Picture         =   "frmScreen.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Image SAPI 
      Height          =   1335
      Left            =   1170
      Picture         =   "frmScreen.frx":08CA
      Stretch         =   -1  'True
      Top             =   1125
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.Image JCI 
      Height          =   1740
      Left            =   1170
      Picture         =   "frmScreen.frx":B090
      Top             =   900
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2385
      TabIndex        =   4
      Top             =   270
      Width           =   900
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Suite 1404, Summit one Tower, 530 Shaw Blvd., Mandaluyong City 1552 Philippines"
      Height          =   600
      Left            =   2565
      TabIndex        =   2
      Top             =   2835
      Width           =   2715
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      X1              =   -45
      X2              =   9795
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) January 2010"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   3150
      Width           =   2535
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   2925
      Width           =   810
   End
End
Attribute VB_Name = "frmScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------
'PROGRAM : SEARCH ASIA PAYROLL
'DEVELOPER : IT PROS
'DATE DEVELOPED : 052505
'STATUS: TO ENHANCE
'------------------------------------

Option Explicit

Dim counter As Integer
   
'CompanyCode = 1 SAPI
'CompanyCode = 2 JCI
'CompanyCode = 3 CSPI

'TODO: To read company code from database 1272008

Private Sub Form_Load()
Dim rs_log As New ADODB.Recordset

On Error GoTo ErrConnection

CompanyName = "TAS Trans Group of Companies (ADMIN)"
CSPI.Visible = True

Me.Picture = LoadPicture(App.Path & "\pics\tas.gif")
counter = 0
Timer1.Enabled = True
pbar.VALUE = 0
Me.MousePointer = vbHourglass

StrQuery = "Select Username, Password, operatorid, superuser, LoginStatus, AccountType from Userslist"
Set rs_log = CheckConnection(StrQuery, cnnstr, mstatic)

Set rs_log = Nothing

Exit Sub

ErrConnection:

    MsgBox "Cannot establish connection to SQL Server. " & Err.Description & " The system will exit.", vbInformation, PayrollAlert
    Err.Clear
    End
    
End Sub

Private Function CheckConnection(ByVal strSQL As String, _
                        ByVal strConString As String, _
                        Optional rInt As mInt = 0) As ADODB.Recordset
    Set cnnAction = New ADODB.Connection
    Dim cmdcomm As ADODB.Command
    Dim rs As New ADODB.Recordset

    cnnAction.Open strConString
    
    Set cmdcomm = New ADODB.Command
    cmdcomm.ActiveConnection = cnnAction
    cmdcomm.CommandType = adCmdText
    cmdcomm.CommandText = strSQL
    
    rs.CursorLocation = adUseClient    'Set rs.ActiveConnection = cnnAction
    Select Case rInt
        Case mInt.mread
            Set rs = cmdcomm.Execute
        Case mInt.mstatic
            rs.Open cmdcomm, , adOpenStatic, adLockReadOnly, adCmdText
        Case mInt.mKeySet
            rs.Open cmdcomm, , adOpenKeyset, adLockReadOnly, adCmdText
        Case mInt.mOptim
            rs.Open cmdcomm, , adOpenDynamic, adLockOptimistic, adCmdText
    End Select
     
    Set CheckConnection = rs
    Set cnnAction = Nothing
    Set rs = Nothing
    
End Function

Private Sub Timer1_Timer()

If counter <= 100 Then
  '
    If pbar.VALUE = 100 Then
        Timer1.Interval = 0
        Timer1.Enabled = False
          
        pbar.VALUE = 100
    
        Me.Hide
        frmLogin.Show
    Else
        If counter < 100 Then counter = counter + 4 Else Timer1.Interval = 0
        pbar.VALUE = counter
    End If
  
End If
  
Me.MousePointer = vbDefault
  
End Sub

