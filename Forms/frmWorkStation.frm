VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmWorkStation 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Work Station"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   Icon            =   "frmWorkStation.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdlast 
      Height          =   315
      Left            =   4680
      Picture         =   "frmWorkStation.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Last record"
      Top             =   1530
      Width           =   255
   End
   Begin VB.CommandButton cmdNext 
      Height          =   315
      Left            =   4440
      Picture         =   "frmWorkStation.frx":0A14
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Next record"
      Top             =   1530
      Width           =   255
   End
   Begin VB.CommandButton cmdPrev 
      Height          =   315
      Left            =   4200
      Picture         =   "frmWorkStation.frx":0B5E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Previous record"
      Top             =   1530
      Width           =   255
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   315
      Left            =   3960
      Picture         =   "frmWorkStation.frx":0CA8
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Top record"
      Top             =   1530
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Height          =   1200
      Left            =   45
      TabIndex        =   0
      Top             =   495
      Width           =   5010
      Begin VB.ComboBox cmbArea 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   315
         ItemData        =   "frmWorkStation.frx":0DF2
         Left            =   2400
         List            =   "frmWorkStation.frx":0DFF
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "MovementTypeDesc"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   2
         Top             =   705
         Width           =   3645
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         DataField       =   "MovementType"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   960
         TabIndex        =   1
         Top             =   267
         Width           =   885
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Area:"
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
         Left            =   1920
         TabIndex        =   16
         Top             =   285
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Work Station:"
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
         Top             =   723
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Work ID:"
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
         Top             =   300
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   767
      ButtonWidth     =   609
      ButtonHeight    =   609
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton cmdDelete 
         Height          =   405
         Left            =   1410
         Picture         =   "frmWorkStation.frx":0E22
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   495
         Picture         =   "frmWorkStation.frx":0F24
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   945
         Picture         =   "frmWorkStation.frx":106E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   45
         Picture         =   "frmWorkStation.frx":15A0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "New"
         Top             =   0
         Width           =   420
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   1845
         Picture         =   "frmWorkStation.frx":1AD2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   420
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Index           =   0
      Left            =   45
      TabIndex        =   15
      Top             =   2040
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   3201
      SortKey         =   1
      View            =   3
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Branch Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Area"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "frmWorkStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New ADODB.Recordset
Dim firstRun As Boolean

Private Sub Form_Load()
Dim rsArea As ADODB.Recordset

firstRun = True
'StrQuery = "SELECT a.*, b.AreaName FROM HRWorkStation a INNER JOIN HRArea b ON a.AreaCode = b.AreaCode Order by a.BrName ASC"
StrQuery = "SELECT * FROM HRWorkStation Order by BrName ASC"
Set rst = ReturnRst(StrQuery, cnnstr, mstatic)
With rst
    If .EOF = False Then
        rst.MoveFirst
        Showdata
        Call EnableNavigator(False, True, True, True)
    Else
        txtID.Text = ""
        txtName.Text = ""
    End If
End With
  
StrQuery = "SELECT * FROM HRArea WHERE AreaFlag = 1 ORDER BY AreaCode"
Set rsArea = ReturnRst(StrQuery, CnnstrHR, mstatic)
If rsArea.RecordCount <> 0 Then
    Call FillComboData(cmbArea, rsArea)
    cmbArea.ListIndex = 0
End If
rsArea.Close
firstRun = False
Call EnableButton(True, True, False, True, False)

End Sub

Private Function FillComboData(ByRef mCmb As ComboBox, ByRef mRs As Recordset)
    mCmb.Clear
    While Not (mRs.EOF Or mRs.BOF)
        mCmb.AddItem Trim(Replace(IIf(IsNull(mRs.Fields(1)), "None", mRs.Fields(1)), "'", "''"))
        mCmb.ItemData(mCmb.NewIndex) = mRs.Fields(0)
        mRs.MoveNext
    Wend
End Function

Private Sub FillListView(ByRef mList As ListView, _
                        ByRef mRs As Recordset)
Dim litem1 As ListItem

    mList.View = lvwReport
    mList.ListItems.Clear
    While Not (mRs.EOF Or mRs.BOF)
        Set litem1 = mList.ListItems.Add(, , mRs(0))
        litem1.SubItems(1) = IIf(IsNull(mRs(1)), "Blank", mRs(1))
        'Set litem1 = mList.ListItems.Add(, , IIf(IsNull(mRs(1)), "Blank", mRs(1)))
        litem1.SubItems(2) = IIf(IsNull(mRs.Fields("AreaName")), "Blank", mRs.Fields("AreaName"))
        mRs.MoveNext
    Wend
End Sub

Private Sub cmdCancel_Click()
    Call EnableButton(True, True, False, True, False)
    Call EnableNavigator(True, True, True, True)
    ListView1(0).Enabled = True
    txtName.Enabled = False
    Showdata
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Delete Work Station " & txtName & "?", vbQuestion + vbYesNo) = vbYes Then
       'Call Action_Query("DELETE FROM HRWorkStation WHERE brcode= " & txtID & "", CnnstrHR)
        If CheckHRCodes("SELECT * FROM HREmployees WHERE brcode = " & txtID) = False Then
            If Action_Query("DELETE HRWorkStation WHERE brcode = " & txtID & "", cnnstr) = True Then
                MsgBox "Work Station " & txtName & " was successfully deleted.", vbInformation
            Else
                MsgBox "Work Station was not successfully deleted.", vbInformation
            End If
        Else
            MsgBox "Work Station " & txtName & " still used by HR employees, cannot be deleted.", vbInformation
        End If
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
    Call EnableNavigator(False, False, False, False)
    ListView1(0).Enabled = False
    txtName.Enabled = True
    cmbArea.Enabled = True
    txtName.SetFocus
End Sub

Private Sub cmdfirst_Click()
   rst.MoveFirst
   If rst.BOF = True Then
       MsgBox "Record not Found."
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
    txtName.Enabled = True
    cmbArea.Enabled = True
    txtName.Text = ""
    txtName.SetFocus
    Call EnableButton(False, False, True, False, True)
    Call EnableNavigator(False, False, False, False)
End Sub

Private Sub cmdSave_Click()
Dim mRSNew As ADODB.Recordset
Dim NewHRID As Integer

 If Trim(txtName.Text) = "" Then
        MsgBox "Please Enter Work Station name first.", vbInformation
        txtName.SetFocus
 Else
'    If DuplicateBranch(txtName) = True Then
'        'MsgBox "Duplicate branch, please change.", vbInformation
'        'txtName.SetFocus
'    Else
        If IsNumeric(txtID) = True Then
            If Action_Query("UPDATE HRWorkStation SET brname = '" & txtName & "', AreaCode = " & cmbArea.ListIndex + 1 & " WHERE brcode = " & txtID & "", cnnstr) = True Then
                If Action_Query("UPDATE HRWorkStation SET brname = '" & txtName & "', AreaCode = " & cmbArea.ListIndex + 1 & " WHERE brcode = " & txtID & "", CnnstrHR) = True Then
                    MsgBox "Work Station successfully updated.", vbInformation
                Else
                    MsgBox "Work Station was not successfully updated. Please repeat updates.", vbInformation
                End If
                'rst.Requery
                Call Form_Load
                rst.Find "brcode = " & txtID & ""
            Else
                MsgBox "Work Station was not successfully updated. Please repeat updates.", vbInformation
            End If
        Else
            ' Save first to Master HR DB.
'            If Action_Query("INSERT INTO HRWorkStation (brname, areacode) values ('" & txtName & "', " & cmbArea.ListIndex + 1 & ")", CnnstrHR) = True Then
'                ' Get the HR ID number.
'                StrQuery = "SELECT * FROM HRWorkStation WHERE BrName = '" & txtName & "' and areacode = " & cmbArea.ListIndex + 1
'                Set mRSNew = ReturnRst(StrQuery, CnnstrHR, mstatic)
'                If mRSNew.EOF = True Then ' was not successfully saved
'                    MsgBox "New Work Station was not successfully saved! Please repeat entry.", vbInformation
'                    txtName.SetFocus
'                Else
'                    NewHRID = mRSNew!BrCode
'                End If
'                Set mRSNew = Nothing
                
            If Action_Query("INSERT INTO HRWorkStation (brName, AreaCode) values ( '" & txtName & "', " & cmbArea.ListIndex + 1 & ")", cnnstr) = True Then
                MsgBox "New Work Station successfully saved!", vbInformation
                Call EnableNavigator(True, True, False, False)
                rst.Requery
                Call Form_Load
                rst.MoveLast
            Else
                MsgBox "New Work Station was not successfully saved! Please repeat entry.", vbInformation
            End If
        End If
    'End If
    Showdata
    txtName.Enabled = False
    cmbArea.Enabled = False
    Call EnableButton(True, True, False, True, False)
    ListView1(0).Enabled = True
 End If
End Sub

Private Function DuplicateBranch(ByVal BName As String) As Boolean
Dim mRSDup As ADODB.Recordset

StrQuery = "SELECT * FROM HRWorkStation WHERE BrName = '" & BName & "'"
Set mRSDup = ReturnRst(StrQuery, CnnstrHR, mstatic)
If mRSDup.EOF = False Then
    DuplicateBranch = True
Else
    DuplicateBranch = False
End If
Set mRSDup = Nothing

End Function

Private Sub EnableNavigator(Optional ByVal Blcmdfirst As Boolean, Optional ByVal BlcmdPrev As Boolean, Optional ByVal BlcmdNext As Boolean, Optional ByVal Blcmdlast As Boolean)
    cmdfirst.Enabled = Blcmdfirst
    cmdPrev.Enabled = BlcmdPrev
    cmdNext.Enabled = BlcmdNext
    cmdlast.Enabled = Blcmdlast
End Sub

Private Sub EnableButton(Optional ByVal Add As Boolean, Optional ByVal Edit As Boolean, Optional ByVal Save As Boolean, Optional ByVal Delete As Boolean, Optional ByVal Cancel As Boolean)
    cmdadd.Enabled = Add
    cmdedit.Enabled = Edit
    cmdsave.Enabled = Save
    cmddelete.Enabled = Delete
    cmdcancel.Enabled = Cancel
End Sub

Private Sub Showdata()

With rst
    '.MoveFirst
    txtID.Text = .Fields!BrCode
    txtName.Text = .Fields!BrName
'    If firstRun = False Then
'        cmbArea = LTrim(RTrim(.Fields!AreaName))
'    End If
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

'If frmHRFinalEmployees.Visible = True Then
If FeMP = True Then
    Call frmHRFinalEmployees.RefreshBranch
End If
'End If

End Sub

Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    rst.MoveFirst
    rst.Find "brcode = " & Item.Text
    Showdata
End Sub

Private Sub txtName_GotFocus()
    txtName.BackColor = &HFFFFC0
End Sub

Private Sub txtName_LostFocus()
    txtName.BackColor = vbWhite
End Sub
