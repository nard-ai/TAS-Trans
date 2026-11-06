VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBank 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank  Details"
   ClientHeight    =   3645
   ClientLeft      =   5475
   ClientTop       =   4155
   ClientWidth     =   4380
   Icon            =   "frmBank.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4380
   Begin VB.CommandButton cmdlast 
      Height          =   255
      Left            =   3960
      Picture         =   "frmBank.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Last record"
      Top             =   450
      Width           =   255
   End
   Begin VB.CommandButton cmdNext 
      Height          =   255
      Left            =   3720
      Picture         =   "frmBank.frx":0A14
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Next record"
      Top             =   450
      Width           =   255
   End
   Begin VB.CommandButton cmdPrev 
      Height          =   255
      Left            =   3480
      Picture         =   "frmBank.frx":0B5E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Previous record"
      Top             =   450
      Width           =   255
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   255
      Left            =   3240
      Picture         =   "frmBank.frx":0CA8
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Top record"
      Top             =   450
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Height          =   3045
      Left            =   45
      TabIndex        =   0
      Top             =   495
      Width           =   4245
      Begin VB.TextBox TxtTelno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1260
         MaxLength       =   25
         TabIndex        =   21
         Top             =   1575
         Width           =   2835
      End
      Begin VB.TextBox TxtPosition 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1260
         MaxLength       =   25
         TabIndex        =   19
         Top             =   2295
         Width           =   2835
      End
      Begin VB.TextBox TxtAddress 
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   555
         Left            =   1260
         TabIndex        =   16
         Top             =   945
         Width           =   2835
      End
      Begin VB.TextBox TxtContactPerson 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1260
         MaxLength       =   25
         TabIndex        =   15
         Top             =   1935
         Width           =   2835
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1260
         TabIndex        =   2
         Top             =   585
         Width           =   2835
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1260
         TabIndex        =   1
         Top             =   270
         Width           =   1800
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. No"
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
         Left            =   135
         TabIndex        =   22
         Top             =   1620
         Width           =   1290
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   135
         TabIndex        =   20
         Top             =   1260
         Width           =   1020
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact person"
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
         Left            =   135
         TabIndex        =   18
         Top             =   2025
         Width           =   1290
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
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
         Left            =   135
         TabIndex        =   17
         Top             =   2385
         Width           =   1140
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank name"
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
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank ID"
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
         Top             =   375
         Width           =   1065
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   767
      ButtonWidth     =   609
      ButtonHeight    =   609
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton cmdDelete 
         Height          =   405
         Left            =   1410
         Picture         =   "frmBank.frx":0DF2
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   495
         Picture         =   "frmBank.frx":0EF4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   945
         Picture         =   "frmBank.frx":103E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   45
         Picture         =   "frmBank.frx":1570
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "New"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   1845
         Picture         =   "frmBank.frx":1AA2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New ADODB.Recordset

Private Sub Form_Load()
  Set rst = ReturnRst("Select * from tblbank Order by bankid ASC", cnnStr, mstatic)
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
    If MsgBox("Are you sure you want to remove this data?", vbQuestion + vbYesNo) = vbYes Then
       Call Action_Query(" DELETE FROM tblbank where bankid= " & txtID & "", cnnStr)
       Call EnableButton(True, True, False, True, False)
       Call EnableNavigator(True, True, True, True)
       rst.Requery
       Call Form_Load
       Showdata
    Else
    End If
End Sub

Private Sub cmdEdit_Click()
    Call EnableButton(False, False, True, False, True)
    TextBox (True)
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
       MsgBox "Record not Found."
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
       MsgBox "Beginning of File."
       Call EnableNavigator(False, False, True, True)
       rst.MoveFirst
   Else
       Showdata
       Call EnableNavigator(True, True, True, True)
   End If
End Sub
Private Sub cmdAdd_Click()
    txtID.Text = "<auto number>"
    TextBox (True)
    
    txtName.Text = ""
    TxtAddress.Text = ""
    TxtPosition.Text = ""
    TxtContactPerson.Text = ""
    TxtTelno.Text = ""
    
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
        Call Action_Query(" update tblbank set bankname='" & txtName & "',address='" & TxtAddress & "',contactperson='" & TxtContactPerson & "',[position]='" & TxtPosition & "',telno='" & TxtTelno & "' where bankid= " & txtID & "", cnnStr)
        MsgBox "Updated.", vbInformation
        rst.Requery
        rst.Find "bankid = " & txtID & ""
    Else
        Call Action_Query("iNSERT into tblbank(bankname,address,contactperson,[position],telno) values ('" & txtName & "','" & TxtAddress & "','" & TxtContactPerson & "','" & TxtPosition & "','" & TxtTelno & "')", cnnStr)
        MsgBox "saved", vbInformation
        Call EnableNavigator(True, True, False, False)
        rst.Requery
        Call Form_Load
        rst.MoveLast
        ' rst.Find "brcode = " & txtid & ""
    End If
    Showdata
   TextBox (False)
    
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
    cmddelete.Enabled = Delete
    cmdCancel.Enabled = Cancel
End Sub

Private Sub Showdata()
 With rst
    txtID.Text = .Fields!BANKID
    txtName.Text = .Fields!bankname
    TxtAddress.Text = .Fields!Address
    TxtContactPerson.Text = .Fields!Contactperson
        TxtPosition.Text = .Fields!Position
        TxtTelno.Text = .Fields!telno
        
    
 End With
End Sub
Private Function TextBox(ByVal txtval As Boolean)
    txtName.Enabled = txtval
    TxtAddress.Enabled = txtval
    TxtContactPerson.Enabled = txtval
    TxtPosition.Enabled = txtval
    TxtTelno.Enabled = txtval
End Function
Private Sub Form_Unload(Cancel As Integer)

'If ISVisibleHRF = True Then
'If FeMP = True Then
'Call frmHRFinalEmployees.RefreshBranch
'End If
'End If
D:
End Sub

Private Sub TxtAddress_GotFocus()
TxtAddress.BackColor = &HFFFFC0
End Sub

Private Sub TxtAddress_LostFocus()
TxtAddress.BackColor = vbWhite
End Sub

Private Sub TxtContactPerson_GotFocus()
TxtContactPerson.BackColor = &HFFFFC0
End Sub

Private Sub TxtContactPerson_LostFocus()
TxtContactPerson.BackColor = vbWhite
End Sub

Private Sub txtName_GotFocus()
    txtName.BackColor = &HFFFFC0
End Sub

Private Sub txtName_LostFocus()
    txtName.BackColor = vbWhite
End Sub


Private Sub TxtPosition_GotFocus()
TxtPosition.BackColor = &HFFFFC
End Sub



Private Sub TxtPosition_LostFocus()
TxtPosition.BackColor = vbWhite
End Sub

Private Sub TxtTelno_GotFocus()
TxtTelno.BackColor = &HFFFFC
End Sub

Private Sub TxtTelno_LostFocus()
TxtTelno.BackColor = vbWhite
End Sub
