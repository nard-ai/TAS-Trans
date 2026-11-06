VERSION 5.00
Begin VB.Form frmHRMainLeave 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leave Application"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2850
   ScaleWidth      =   3810
   Begin VB.CommandButton btnExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   938
      TabIndex        =   3
      Top             =   2048
      Width           =   1935
   End
   Begin VB.CommandButton btnNoted 
      Caption         =   "&Note"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton btnApproved 
      Caption         =   "A&pprove"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   938
      TabIndex        =   1
      Top             =   848
      Width           =   1935
   End
   Begin VB.CommandButton btnApply 
      Caption         =   "&Apply"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   248
      Width           =   1935
   End
End
Attribute VB_Name = "frmHRMainLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnApply_Click()
    frmHRLeave.Show
    
End Sub

Private Sub btnApproved_Click()
    frmHRLeaveAprvd.Show vbModal
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnNoted_Click()
frmHRLeaveNoted.Show vbModal
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\pics\bg.gif")

   Me.Top = (mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
    
 If UL = "" Then
    Select Case HR
        Case "MG"
            btnApply.Enabled = True
            btnApproved.Enabled = True
            btnNoted.Enabled = True
    
        Case "SU"
            btnApply.Enabled = True
            btnApproved.Enabled = True
            btnNoted.Enabled = False
            
        Case "EN"
            btnApply.Enabled = True
            btnApproved.Enabled = False
            btnNoted.Enabled = False
 
    End Select
 
 ElseIf UL = "AD" Then
            btnApply.Enabled = True
            btnApproved.Enabled = False
            btnNoted.Enabled = False
 End If
 
End Sub
