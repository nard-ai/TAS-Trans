VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPositions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HR Position"
   ClientHeight    =   1785
   ClientLeft      =   5925
   ClientTop       =   5025
   ClientWidth     =   4410
   Icon            =   "frmPositions.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4410
   Begin VB.CommandButton cmdlast 
      Height          =   255
      Left            =   3840
      Picture         =   "frmPositions.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Last record"
      Top             =   1410
      Width           =   255
   End
   Begin VB.CommandButton cmdNext 
      Height          =   255
      Left            =   3600
      Picture         =   "frmPositions.frx":0A14
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Next record"
      Top             =   1410
      Width           =   255
   End
   Begin VB.CommandButton cmdPrev 
      Height          =   255
      Left            =   3360
      Picture         =   "frmPositions.frx":0B5E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Previous record"
      Top             =   1410
      Width           =   255
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   255
      Left            =   3120
      Picture         =   "frmPositions.frx":0CA8
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Top record"
      Top             =   1410
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Height          =   1080
      Left            =   45
      TabIndex        =   0
      Top             =   495
      Width           =   4290
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
         Width           =   2805
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
         Width           =   1005
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Name"
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
         Caption         =   "Position ID"
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
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   767
      ButtonWidth     =   609
      ButtonHeight    =   609
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton cmdDelete 
         Height          =   405
         Left            =   1410
         Picture         =   "frmPositions.frx":0DF2
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   495
         Picture         =   "frmPositions.frx":0EF4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   945
         Picture         =   "frmPositions.frx":103E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   45
         Picture         =   "frmPositions.frx":1570
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "New"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   1845
         Picture         =   "frmPositions.frx":1AA2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmPositions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New ADODB.Recordset

Private Sub Form_Load()
  Set rst = ReturnRst("Select * from HRPositions Order by Positioncode ASC", cnnstr, mstatic)
  With rst
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
    If MsgBox("Delete position " & txtName.Text & "?", vbQuestion + vbYesNo) = vbYes Then
        If CheckHRCodes("SELECT * FROM HREmployees WHERE positioncode = " & txtID) = False Then
           If Action_Query(" DELETE FROM HRPositions where Positioncode = " & txtID & "", cnnstr) = True Then
                Call EnableButton(True, True, False, True, False)
                Call EnableNavigator(True, True, True, True)
                rst.Requery
                Call Form_Load
                Showdata
            Else
                MsgBox "Cannot delete position this time, please close the system and try again!", vbInformation, PayrollAlert
            End If
        Else
            MsgBox "Position " & txtName.Text & " still used by HR Employees, cannot delete this time!", vbInformation, PayrollAlert
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    Call EnableButton(False, False, True, False, True)
    txtName.Enabled = True
    txtName.SetFocus
End Sub

Private Sub cmdfirst_Click()
   rst.MoveFirst
   If rst.BOF = True Then
       MsgBox "record not found"
       Call EnableNavigator(False, False, False, False)
   Else
       Showdata
       Call EnableNavigator(False, False, True, True)
   End If
End Sub
Private Sub cmdlast_Click()
   rst.MoveLast
   If rst.EOF = True Then
       MsgBox "record not found"
       Call EnableNavigator(False, False, False, False)
   Else
       Showdata
       Call EnableNavigator(True, True, False, False)
   End If
End Sub
Private Sub cmdNext_Click()
   rst.MoveNext
   If rst.EOF = True Then
      MsgBox "End of File."
      Call EnableNavigator(True, True, False, False)
      rst.MoveLast
   Else
      Showdata
      Call EnableNavigator(True, True, True, True)
   End If
End Sub

Private Sub cmdPrev_Click()
   rst.MovePrevious
   If rst.BOF = True Then
       MsgBox "Beginning of file."
       Call EnableNavigator(False, False, True, True)
       rst.MoveFirst
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
        txtName.SetFocus
 Else
    If IsNumeric(txtID) = True Then
        Call Action_Query(" update HRPositions set Positionname='" & txtName & "' where Positioncode= " & txtID & "", cnnstr)
        MsgBox "HR Department was Updated.", vbInformation
        rst.Requery
            rst.Find "Positioncode = " & txtID & ""
    Else
        Call Action_Query("Insert into HRPositions(Positionname) values ('" & txtName & "')", cnnstr)
        MsgBox "New Department was Saved!", vbInformation
        Call EnableNavigator(True, True, False, False)
        rst.Requery
        Call Form_Load
        rst.MoveLast
        ' rst.Find "hrcode = " & txtid & ""
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

Private Sub EnableButton(Optional ByVal Add As Boolean, Optional ByVal Edit As Boolean, Optional ByVal save As Boolean, Optional ByVal Delete As Boolean, Optional ByVal Cancel As Boolean)
    cmdadd.Enabled = Add
    cmdedit.Enabled = Edit
    cmdsave.Enabled = save
    cmdDelete.Enabled = Delete
    cmdcancel.Enabled = Cancel
End Sub

Private Sub Showdata()
 With rst
    txtID.Text = .Fields!Positioncode
    txtName.Text = .Fields!Positionname
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

'If frmHRFinalEmployees.Visible = True Then
If FeMP = True Then
    Call frmHRFinalEmployees.RefreshPosition
End If

'End If

End Sub

Private Sub txtName_GotFocus()
    txtName.BackColor = &HFFFFC0
End Sub

Private Sub txtName_LostFocus()
    txtName.BackColor = vbWhite
End Sub
