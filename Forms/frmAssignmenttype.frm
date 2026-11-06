VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAssignType 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assignment type"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   3675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdlast 
      Height          =   255
      Left            =   3240
      Picture         =   "frmAssignmenttype.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Last record"
      Top             =   450
      Width           =   255
   End
   Begin VB.CommandButton cmdNext 
      Height          =   255
      Left            =   3000
      Picture         =   "frmAssignmenttype.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Next record"
      Top             =   450
      Width           =   255
   End
   Begin VB.CommandButton cmdPrev 
      Height          =   255
      Left            =   2760
      Picture         =   "frmAssignmenttype.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Previous record"
      Top             =   450
      Width           =   255
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   255
      Left            =   2520
      Picture         =   "frmAssignmenttype.frx":03DE
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Top record"
      Top             =   450
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Height          =   1200
      Left            =   45
      TabIndex        =   0
      Top             =   495
      Width           =   3570
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   1260
         MaxLength       =   25
         TabIndex        =   2
         Top             =   585
         Width           =   2205
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   1260
         TabIndex        =   1
         Top             =   270
         Width           =   2205
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Branch name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   675
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Branch ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   330
         Width           =   1695
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   767
      ButtonWidth     =   609
      ButtonHeight    =   609
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton cmdDelete 
         Height          =   405
         Left            =   1410
         Picture         =   "frmAssignmenttype.frx":0528
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   495
         Picture         =   "frmAssignmenttype.frx":062A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   945
         Picture         =   "frmAssignmenttype.frx":0774
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   45
         Picture         =   "frmAssignmenttype.frx":0CA6
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "New"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   1845
         Picture         =   "frmAssignmenttype.frx":11D8
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmAssignType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsT As New ADODB.Recordset

Private Sub Form_Load()
  Set rsT = ReturnRst("Select * from HRAssignmentType Order by AssignType ASC", Cnnstr, mstatic)
  With rsT
    If .EOF = False Then
        Showdata
        Call EnableNavigator(False, True, True, True)
    Else
        txtID.Text = ""
        txtName.Text = ""
    End If
  End With
End Sub
Private Sub cmdCancel_Click()
    Call EnableButton(True, True, False, True, False)
    Call EnableNavigator(True, True, True, True)
    txtName.Enabled = False
    Showdata
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Are you sure want to remove this data?", vbQuestion + vbYesNo) = vbYes Then
       Call Action_Query(" DELETE FROM HRAssignmentType where AssignType= " & txtID & "", Cnnstr)
       Call EnableButton(True, True, False, True, False)
       Call EnableNavigator(True, True, True, True)
       rsT.Requery
       Call Form_Load
       Showdata
    Else
    End If
End Sub

Private Sub cmdEdit_Click()
    Call EnableButton(False, False, True, False, True)
    txtName.Enabled = True
    txtName.SetFocus
End Sub

Private Sub cmdfirst_Click()
   rsT.MoveFirst
   If rsT.BOF = True Then
       MsgBox "record not found"
       Call EnableNavigator(False, False, False, False)
   Else
       Showdata
       Call EnableNavigator(False, False, True, True)
   End If
End Sub
Private Sub cmdlast_Click()
   rsT.MoveLast
   If rsT.EOF = True Then
       MsgBox "record not found"
       Call EnableNavigator(False, False, False, False)
   Else
       Showdata
       Call EnableNavigator(True, True, False, False)
   End If
End Sub
Private Sub cmdnext_Click()
   rsT.MoveNext
   If rsT.EOF = True Then
      MsgBox "end of file"
      Call EnableNavigator(True, True, False, False)
      rsT.MoveLast
   Else
      Showdata
      Call EnableNavigator(True, True, True, True)
   End If
End Sub

Private Sub cmdprev_Click()
   rsT.MovePrevious
   If rsT.BOF = True Then
       MsgBox "begin of file"
       Call EnableNavigator(False, False, True, True)
       rsT.MoveFirst
   Else
       Showdata
       Call EnableNavigator(True, True, True, True)
   End If
End Sub
Private Sub cmdAdd_Click()
    txtID.Text = "<auto number>"
    txtName.Enabled = True
    txtName.Text = ""
    txtName.SetFocus
    Call EnableButton(False, False, True, False, True)
    Call EnableNavigator(False, False, False, False)
End Sub
Private Sub cmdSave_Click()
 If Trim(txtName.Text) = "" Then
        MsgBox "Enter name.", vbInformation
        txtName.Enabled = True
        txtName.SetFocus
 Else
    If IsNumeric(txtID) = True Then
        Call Action_Query(" update HRAssignmentType set AssignTypeDesc='" & txtName & "' where AssignType= " & txtID & "", Cnnstr)
        MsgBox "Updated.", vbInformation
        rsT.Requery
        rsT.Find "AssignType = " & txtID & ""
    Else
        Call Action_Query("iNSERT into HRAssignmentType(AssignTypeDesc) values ('" & txtName & "')", Cnnstr)
        MsgBox "saved", vbInformation
        Call EnableNavigator(True, True, False, False)
        rsT.Requery
        Call Form_Load
        rsT.MoveLast
        ' rst.Find "AssignType = " & txtid & ""
    End If
    Showdata
    txtName.Enabled = False
    Call EnableButton(True, True, False, True, False)
 End If
End Sub

Private Sub EnableNavigator(Optional ByVal Blcmdfirst As Boolean, Optional ByVal BlcmdPrev As Boolean, Optional ByVal BlcmdNext As Boolean, Optional ByVal Blcmdlast As Boolean)
    cmdfirst.Enabled = Blcmdfirst
    cmdPrev.Enabled = BlcmdPrev
    cmdNext.Enabled = BlcmdNext
    cmdlast.Enabled = Blcmdlast
End Sub

Private Sub EnableButton(Optional ByVal add As Boolean, Optional ByVal edit As Boolean, Optional ByVal save As Boolean, Optional ByVal Delete As Boolean, Optional ByVal Cancel As Boolean)
    cmdadd.Enabled = add
    cmdedit.Enabled = edit
    cmdsave.Enabled = save
    cmdDelete.Enabled = Delete
    cmdcancel.Enabled = Cancel
End Sub

Private Sub Showdata()
 With rsT
    txtID.Text = .Fields!AssignType
    txtName.Text = .Fields!AssignTypeDesc
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
If frmAssignType.Visible = True Then
Call frmAssignType.RefreshAssignType
End If

End Sub

Private Sub txtName_GotFocus()
    txtName.BackColor = &HFFFFC0
End Sub

Private Sub txtName_LostFocus()
    txtName.BackColor = vbWhite
End Sub
