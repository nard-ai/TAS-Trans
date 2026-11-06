VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeductionTypes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Deduction Types"
   ClientHeight    =   3870
   ClientLeft      =   9285
   ClientTop       =   7170
   ClientWidth     =   6360
   Icon            =   "frmDeductionTypes.frx":0000
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6360
   Begin VB.CommandButton cmdlast 
      Height          =   255
      Left            =   5880
      Picture         =   "frmDeductionTypes.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Last record"
      Top             =   450
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdNext 
      Height          =   255
      Left            =   5640
      Picture         =   "frmDeductionTypes.frx":0A14
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Next record"
      Top             =   450
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdPrev 
      Height          =   255
      Left            =   5400
      Picture         =   "frmDeductionTypes.frx":0B5E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Previous record"
      Top             =   450
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   255
      Left            =   5160
      Picture         =   "frmDeductionTypes.frx":0CA8
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Top record"
      Top             =   450
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Height          =   3285
      Left            =   45
      TabIndex        =   0
      Top             =   495
      Width           =   6240
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Incentives"
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Deductions"
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   120
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkHide 
         BackColor       =   &H80000009&
         Caption         =   "Hide When Zero"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4560
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chkType 
         BackColor       =   &H80000009&
         Caption         =   "Active"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtHeading 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1380
         MaxLength       =   12
         TabIndex        =   16
         Top             =   1095
         Width           =   3315
      End
      Begin VB.TextBox txtEntryDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4740
         MaxLength       =   25
         TabIndex        =   15
         Top             =   1095
         Width           =   1395
      End
      Begin VB.TextBox txtEntryBy 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4740
         MaxLength       =   25
         TabIndex        =   14
         Top             =   735
         Width           =   1395
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   2
         Top             =   735
         Width           =   3315
      End
      Begin VB.TextBox txtTypes 
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1140
         TabIndex        =   1
         Top             =   315
         Width           =   280
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1590
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   2805
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Active"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descriptions"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Hide When Zero"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Active"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sequence#:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Report Header:"
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
         TabIndex        =   17
         Top             =   1140
         Width           =   1290
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   705
         Width           =   1320
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   767
      ButtonWidth     =   609
      ButtonHeight    =   609
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton cmdDelete 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1410
         Picture         =   "frmDeductionTypes.frx":0DF2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   495
         Picture         =   "frmDeductionTypes.frx":0EF4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   945
         Picture         =   "frmDeductionTypes.frx":103E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdadd 
         Enabled         =   0   'False
         Height          =   405
         Left            =   45
         Picture         =   "frmDeductionTypes.frx":1570
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "New"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   1845
         Picture         =   "frmDeductionTypes.frx":1AA2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmDeductionTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New ADODB.Recordset
Dim Filterkey As Integer

Private Sub Form_Load()

If Option1.VALUE = True Then
Filterkey = 1
Else
Filterkey = 2
End If

    Set rst = ReturnRst("Select * from DeductionTypes2 where filtertype = " & Filterkey & " Order by DeductionCode", CnnstrHR, mstatic)
    Call FillListView
    rst.MoveFirst
    With rst
        If .EOF = False Then
            Showdata
            Call EnableNavigator(False, True, True, True)
            Call EnableButton(True, True, False, True, False)
        End If
    End With
End Sub

Private Sub FillListView()
Dim litem2 As ListItem

ListView1(0).ListItems.Clear
    While Not (rst.EOF Or rst.BOF)
        Set litem2 = ListView1(0).ListItems.Add(, , 1)
        litem2.SubItems(1) = IIf(IsNull(rst!DeductionDescription), "Blank Description", rst!DeductionDescription)
        litem2.SubItems(2) = rst!HideWhenZero
        litem2.SubItems(3) = rst!DeductionFlag
        litem2.SubItems(4) = rst!DeductionCode
        rst.MoveNext
    Wend
End Sub

Private Sub cmdCancel_Click()
    Call EnableButton(True, True, False, True, False)
    Call EnableNavigator(True, True, True, True)
    TextBox (False)
    Option1.Enabled = True
    Option2.Enabled = True
Showdata
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Do you want to delete this record?", vbQuestion + vbYesNo) = vbYes Then
       Call Action_Query(" DELETE FROM DeductionTypes2 where DeductionCode = " & txtTypes & "", CnnstrHR)
       Call EnableButton(True, True, False, True, False)
       Call EnableNavigator(True, True, True, True)
       rst.Requery
       Call Form_Load
       Showdata
    Else
    End If
End Sub

Private Sub cmdEdit_Click()
    If InStr(1, txtDesc.Text, "Reserve", vbTextCompare) = 0 Then
        'MsgBox "Description is not allow to edit! You can only hide or make visible the Deduction list.", vbInformation, PayrollAlert
        Call EnableButton(False, False, True, False, True)
        Call EnableNavigator(False, False, False, False)
        chkType.Enabled = True
        chkHide.Enabled = True
        txtDesc.Enabled = False
        'Option1.Enabled = False
        'Option2.Enabled = False
        txtHeading.Enabled = False
        chkType.SetFocus
    Else
        Call EnableButton(False, False, True, False, True)
        Call EnableNavigator(False, False, False, False)
        chkType.Enabled = True
        chkHide.Enabled = True
        'Option1.Enabled = False
        'Option2.Enabled = False
        txtDesc.Enabled = True
        txtHeading.Enabled = True
        txtDesc.SetFocus
    End If
End Sub

Private Sub cmdfirst_Click()
   rst.MoveFirst
   If rst.BOF = True Then
       MsgBox "Record not Found!"
       Call EnableNavigator(False, False, False, False)
   Else
       Showdata
       Call EnableNavigator(False, False, True, True)
   End If
End Sub

Private Sub cmdlast_Click()
   rst.MoveLast
   If rst.EOF = True Then
       MsgBox "Record not Found!"
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
    txtTypes.Text = "<auto number>"
    txtDesc.Enabled = True
    txtHeading.Enabled = True
    chkType.Enabled = True
    chkHide.Enabled = True
    Option1.Enabled = False
    Option2.Enabled = False
    
    txtDesc.Text = ""
    txtHeading.Text = ""
    chkType.VALUE = 1
    chkType.VALUE = 0
    txtEntryBy.Text = mUserID
    txtEntryDate.Text = Date
    
    txtDesc.SetFocus
    Call EnableButton(False, False, True, False, True)
    Call EnableNavigator(False, False, False, False)
End Sub

Private Sub cmdSave_Click()
 If Trim(txtDesc.Text) = "" Then
        MsgBox "Pleaser Enter Description.", vbInformation
        txtDesc.SetFocus
 Else
    If IsNumeric(txtTypes) = True Then
        Call Action_Query("UPDATE DeductionTypes2 SET DeductionDescription = '" & txtDesc & "', DeductionHeader = '" & txtHeading & "', DeductionFlag = " & chkType.VALUE & ", HideWhenZero = " & chkHide.VALUE & ", EntryBy = '" & IIf(txtEntryBy = "", mUserID, txtEntryBy) & "', EntryDate = '" & Format(Now, "mm/dd/yyyy") & "' WHERE DeductionCode = " & txtTypes, CnnstrHR)
        MsgBox "Successfully Updated.", vbInformation
        Set rst = ReturnRst("Select * from DeductionTypes2 where filtertype = " & Filterkey & " Order by DeductionCode", CnnstrHR, mstatic)
        rst.Requery
Call FillListView
        If txtTypes <> "<auto number>" Then
            rst.Find "DeductionCode = " & txtTypes & ""
        End If
    Else
        Call Action_Query("INSERT into DeductionTypes2 (DeductionDescription, DeductionHeader, DeductionFlag, HideWhenZero, FilterType, EntryBy, EntryDate) Values ('" & txtDesc & "','" & txtHeading & "'," & chkType.VALUE & ", " & chkHide.VALUE & ", " & Filterkey & ", '" & txtEntryBy & "','" & Format(Now, "mm/dd/yyyy") & "')", CnnstrHR)
        MsgBox "Successfully Saved.", vbInformation
        rst.Requery
        Call Form_Load
        rst.MoveLast
    End If
    Showdata
    Option1.Enabled = True
    Option2.Enabled = True
    TextBox (False)
    Call EnableButton(True, True, False, True, False)
    Call EnableNavigator(True, True, False, False)
 End If
End Sub

Private Sub EnableNavigator(Optional ByVal Blcmdfirst As Boolean, Optional ByVal BlcmdPrev As Boolean, Optional ByVal BlcmdNext As Boolean, Optional ByVal Blcmdlast As Boolean)
    cmdfirst.Enabled = Blcmdfirst
    cmdPrev.Enabled = BlcmdPrev
    cmdNext.Enabled = BlcmdNext
    cmdlast.Enabled = Blcmdlast
End Sub

Private Sub EnableButton(Optional ByVal Add As Boolean, Optional ByVal Edit As Boolean, Optional ByVal save As Boolean, Optional ByVal Delete As Boolean, Optional ByVal Cancel As Boolean)
    cmdadd.Enabled = False 'Add
    cmdedit.Enabled = Edit
    cmdsave.Enabled = save
    cmdDelete.Enabled = False 'Delete
    cmdcancel.Enabled = Cancel
End Sub

Private Sub Showdata()
With rst
If .EOF Then
.MoveLast
End If
    txtTypes.Text = .Fields!DeductionCode
    txtDesc.Text = .Fields!DeductionDescription
    txtHeading.Text = .Fields!DeductionHeader
    chkType.VALUE = IIf(.Fields!DeductionFlag = True, 1, 0)
    chkHide.VALUE = IIf(.Fields!HideWhenZero = True, 1, 0)
    txtEntryBy.Text = IIf(IsNull(.Fields!Entryby), "", .Fields!Entryby)
    txtEntryDate.Text = IIf(IsNull(.Fields!EntryDate), Format(Now, "mm/dd/yyyy"), .Fields!EntryDate)
End With
End Sub

Private Function TextBox(ByVal txtval As Boolean)
    txtDesc.Enabled = txtval
    txtHeading.Enabled = txtval
    chkType.Enabled = txtval
    chkHide.Enabled = txtval
End Function

Private Sub Form_Unload(Cancel As Integer)
'If ISVisibleHRF = True Then
'If FeMP = True Then
'Call frmHRFinalEmployees.RefreshBranch
'End If
'End If
D:
End Sub

Private Sub ListView1_Click(Index As Integer)
    rst.Filter = "DeductionCode = " & ListView1(0).SelectedItem.ListSubItems(4).Text
    Showdata
End Sub

Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    rst.Filter = "DeductionCode = " & ListView1(0).SelectedItem.ListSubItems(4).Text
    Showdata
End Sub

Private Sub Option1_Click()
If Option1.VALUE = True Then
Filterkey = 1
Else
Filterkey = 2
End If
Set rst = ReturnRst("Select * from DeductionTypes2 where filtertype = " & Filterkey & " Order by DeductionCode", CnnstrHR, mstatic)
If rst.EOF = False Then
Call FillListView
rst.MoveFirst
End If
End Sub

Private Sub Option2_Click()
If Option2.VALUE = True Then
Filterkey = 2
Else
Filterkey = 1
End If
Set rst = ReturnRst("Select * from DeductionTypes2 where filtertype = " & Filterkey & " Order by DeductionCode", CnnstrHR, mstatic)
If rst.EOF = False Then
Call FillListView
rst.MoveFirst
End If
End Sub

Private Sub txtDesc_GotFocus()
    txtDesc.BackColor = &HFFFFC0
End Sub

Private Sub txtDesc_LostFocus()
    txtDesc.BackColor = vbWhite
End Sub

Private Sub TxtHeading_GotFocus()
    txtHeading.BackColor = &HFFFFC0
End Sub

Private Sub TxtHeading_LostFocus()
    txtHeading.BackColor = vbWhite
End Sub
