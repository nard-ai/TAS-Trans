VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1200
   ClientLeft      =   6885
   ClientTop       =   6540
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   1200
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   5175
      Begin VB.Label msgDisplay 
         Alignment       =   2  'Center
         Caption         =   "Message"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   400
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Top = (mdiMain.Height - Me.Height) / 2
    Me.Left = (mdiMain.Width - Me.Width) / 2
End Sub
