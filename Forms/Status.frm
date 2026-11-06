VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStatus 
   BackColor       =   &H80000004&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Executing action..."
   ClientHeight    =   990
   ClientLeft      =   4875
   ClientTop       =   2085
   ClientWidth     =   5400
   ControlBox      =   0   'False
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   255
      Left            =   420
      TabIndex        =   2
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   180
      Top             =   780
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   180
      Top             =   240
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   600
      Top             =   240
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   780
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4500
      TabIndex        =   0
      Top             =   600
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4260
      TabIndex        =   1
      Top             =   600
      Width           =   675
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
' Status.Animation1.Open App.Path + "\filecopy.avi"
'Status.Animation1.Play
frmStatus.Timer2.Enabled = True
frmStatus.Timer3.Enabled = True
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Bar1.VALUE = Bar1.VALUE + 2
Label1.Caption = Bar1.VALUE

If Bar1.VALUE = 100 Then
Animation1.Close
Timer1.Enabled = False
Unload Me

Timer2.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
Bar1.VALUE = Bar1.VALUE + 5
Label1.Caption = Bar1.VALUE
If Bar1.VALUE = 100 Then
'Animation1.Close
Timer1.Enabled = False
Unload Me
End If
End Sub

Private Sub Timer4_Timer()
Status.Caption = "Save"
Timer4.Enabled = False
End Sub

