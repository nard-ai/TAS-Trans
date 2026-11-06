VERSION 5.00
Begin VB.Form frmChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change password"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   Icon            =   "frmchange.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4035
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   1710
      TabIndex        =   5
      Top             =   1395
      Width           =   960
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   330
      Left            =   2880
      TabIndex        =   2
      Top             =   1395
      Width           =   960
   End
   Begin VB.TextBox TxtNew 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1485
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   900
      Width           =   2355
   End
   Begin VB.TextBox TxtOld 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1485
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   540
      Width           =   2355
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "New password"
      Height          =   240
      Left            =   135
      TabIndex        =   4
      Top             =   945
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current password"
      Height          =   240
      Left            =   135
      TabIndex        =   3
      Top             =   540
      Width           =   1365
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
    
    If TxtOld.Text = "" Then
        MsgBox "Enter old password", vbInformation
        TxtOld.SetFocus
        Exit Sub
    End If
    
    If TxtNew.Text = "" Then
        MsgBox "Enter new password", vbInformation
        TxtNew.SetFocus
        Exit Sub
    End If
    
    Dim rst As New ADODB.Recordset
    
    If rst.State = 1 Then rst.Close
    
    Set rst = ReturnRst("Select * from userslist where operatorid= " & operatorID & "", cnnstr, mstatic)
    If rst.EOF = False Then
        If UCase(rst.Fields!Password) = UCase(TxtOld.Text) Then
            Call Action_Query("Update userslist set [password]= '" & TxtNew.Text & "' where OPERATORid=" & operatorID & "", cnnstr)
            MsgBox "Password successfully changed", vbInformation
            Unload Me
        Else
            MsgBox "Current password is incorrect!", vbInformation
            TxtOld.SetFocus
        End If
    End If

End Sub

Private Sub Form_Load()
Me.Top = 0
Left = 0
End Sub

