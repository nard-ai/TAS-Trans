VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUsers 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Accounts"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmUsers.frx":08CA
   ScaleHeight     =   4215
   ScaleWidth      =   7215
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5535
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsers.frx":16020C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      Height          =   3570
      Left            =   90
      OLEDropMode     =   1  'Manual
      TabIndex        =   12
      Top             =   540
      Width           =   7065
      Begin TabDlg.SSTab SSTab2 
         Height          =   3210
         Left            =   135
         TabIndex        =   16
         Top             =   270
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   5662
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         BackColor       =   -2147483639
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Users list"
         TabPicture(0)   =   "frmUsers.frx":16065E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "List1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2565
            Left            =   135
            TabIndex        =   5
            ToolTipText     =   "User list view"
            Top             =   405
            Width           =   1695
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3210
         Left            =   2205
         TabIndex        =   17
         Top             =   270
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   5662
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   -2147483639
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "User"
         TabPicture(0)   =   "frmUsers.frx":16067A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label1(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "TxtOperator"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmbType"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Roles"
         TabPicture(1)   =   "frmUsers.frx":160696
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "LV"
         Tab(1).ControlCount=   1
         Begin VB.ComboBox cmbType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "frmUsers.frx":1606B2
            Left            =   1560
            List            =   "frmUsers.frx":1606CE
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   960
            Width           =   2775
         End
         Begin VB.TextBox TxtOperator 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1575
            TabIndex        =   0
            ToolTipText     =   "Username Entry"
            Top             =   525
            Width           =   2790
         End
         Begin VB.Frame Frame2 
            Height          =   1365
            Left            =   135
            TabIndex        =   22
            Top             =   1410
            Width           =   4380
            Begin VB.TextBox txtConfirm 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1455
               MaxLength       =   15
               PasswordChar    =   "*"
               TabIndex        =   3
               ToolTipText     =   "Confirmation of password"
               Top             =   945
               Width           =   2790
            End
            Begin VB.TextBox txtpassword 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1455
               MaxLength       =   15
               PasswordChar    =   "*"
               TabIndex        =   2
               ToolTipText     =   "Password Entry"
               Top             =   585
               Width           =   2790
            End
            Begin VB.TextBox txtusername 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1440
               MaxLength       =   15
               TabIndex        =   1
               ToolTipText     =   "Username Entry"
               Top             =   225
               Width           =   2790
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Confirm Password"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   390
               Left            =   150
               TabIndex        =   25
               Top             =   900
               Width           =   1125
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Password"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   150
               TabIndex        =   24
               Top             =   630
               Width           =   885
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Username"
               BeginProperty Font 
                  Name            =   "Verdana"
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
               TabIndex        =   23
               Top             =   270
               Width           =   885
            End
         End
         Begin MSComctlLib.ListView LV 
            Height          =   2625
            Left            =   -74730
            TabIndex        =   4
            Top             =   405
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   4630
            View            =   3
            Arrange         =   1
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            FlatScrollBar   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "ImageList1"
            SmallIcons      =   "ImageList1"
            ColHdrIcons     =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   7232
            EndProperty
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Account Type"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   27
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Operator Name"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   18
            Top             =   570
            Width           =   1380
         End
      End
      Begin VB.Label lblMI 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4680
         TabIndex        =   15
         Top             =   1740
         Width           =   195
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "M.I."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   9630
         TabIndex        =   14
         Top             =   1920
         Width           =   315
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9630
         TabIndex        =   13
         Top             =   1710
         Width           =   195
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0080FFFF&
         BorderWidth     =   3
         X1              =   7830
         X2              =   7620
         Y1              =   2010
         Y2              =   4500
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   873
      ButtonWidth     =   609
      ButtonHeight    =   714
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmddelete 
         Height          =   405
         Left            =   3015
         Picture         =   "frmUsers.frx":16073F
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   510
      End
      Begin VB.CommandButton cmdrefresh 
         Height          =   405
         Left            =   2400
         Picture         =   "frmUsers.frx":160841
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   510
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   0
         Picture         =   "frmUsers.frx":16098B
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "New"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   1215
         Picture         =   "frmUsers.frx":160EBD
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   510
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   1800
         Picture         =   "frmUsers.frx":1613EF
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   510
      End
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   600
         Picture         =   "frmUsers.frx":161921
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   510
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   2
         X1              =   0
         X2              =   800
         Y1              =   1500
         Y2              =   1200
      End
   End
   Begin VB.ComboBox txtcombo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmUsers.frx":161A6B
      Left            =   1305
      List            =   "frmUsers.frx":161A75
      Style           =   2  'Dropdown List
      TabIndex        =   20
      ToolTipText     =   "Use to select level of user"
      Top             =   0
      Width           =   2790
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   135
      Width           =   615
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NOTE:
'LEVEL=1 IS ONLY FOR ADMINISTRATORS
'LEVEL=2 IS ONLY FOR USERS

Dim save As Integer
Dim username As String

Public operatorID As Integer

Private Sub enabledtext()
    TxtOperator.Enabled = True
    txtusername.Enabled = True
    txtConfirm.Enabled = True
    txtpassword.Enabled = True
    cmbType.Enabled = True
    LV.Enabled = True
End Sub

Private Sub disabledtext()
TxtOperator.Enabled = False
txtusername.Enabled = False
txtConfirm.Enabled = False
txtpassword.Enabled = False
cmbType.Enabled = False
LV.Enabled = False
End Sub
    
Private Sub Delete()
Dim Response As String
        Response = MsgBox("Are you sure you want to delete '" & List1.Text & "'?", vbYesNo + vbQuestion, "Delete")
        If Response = vbYes Then
            Call Action_Query("Delete from userslist where USERname='" & txtusername.Text & "'", cnnstr)
            List1.RemoveItem List1.ListIndex
          '  MsgBox "User deleted", vbInformation, "Deleted"
            Call disabledtext
        End If
End Sub

Private Function checkifCheck() As Boolean
Dim X As Integer
X = 1
  Do While X <= LV.ListItems.Count
        If LV.ListItems(X).Checked = True Then
        checkifCheck = True
        
        Exit Function
        End If
        X = X + 1
        
Loop
         checkifCheck = False
         MsgBox "Select roles of User", vbInformation
         SSTab1.Tab = 1
End Function

Private Sub SaveUpdate()

    Dim a, b, c, D, e, f, g, X As Integer
    Dim RsUsers As ADODB.Recordset
    
   If checkifCheck = False Then Exit Sub
   
    a = IIf(LV.ListItems(1).Checked = True, 1, 0) ' Superuser
    b = IIf(LV.ListItems(2).Checked = True, 1, 0) ' Admin
    c = IIf(LV.ListItems(3).Checked = True, 1, 0) ' Client Master
    D = IIf(LV.ListItems(4).Checked = True, 1, 0) ' Employee Master
    e = IIf(LV.ListItems(5).Checked = True, 1, 0) ' User Accounts
    f = IIf(LV.ListItems(6).Checked = True, 1, 0) ' Maintenance
   g = IIf(LV.ListItems(7).Checked = True, 1, 0)  ' Payroll Rights
   h = IIf(LV.ListItems(8).Checked = True, 1, 0)  ' Employee Info
   i = IIf(LV.ListItems(9).Checked = True, 1, 0)  ' HR Report
   j = IIf(LV.ListItems(10).Checked = True, 1, 0) ' TimeKeeping
    
    
    If CheckText = False Then Exit Sub
    
    If txtpassword = txtConfirm Then
        Set RsUsers = New ADODB.Recordset
        StrQuery = "Select * FROM userslist where USERNAME = '" & txtusername.Text & "' OR OPERATOR_NAME = '" & Me.TxtOperator & "'"
        Set RsUsers = ReturnRst(StrQuery, cnnstr, mstatic)
        'RsUsers.Open "select * from userslist where USERNAME= '" & txtusername.Text & "' OR OPERATOR_NAME= '" & Me.TxtOperator & "'", cnnstr, adOpenDynamic, adLockOptimistic
        If save = 1 Then 'FOR ADD
            If RsUsers.EOF = True Then
                Call Action_Query("Insert into usersList ([operator_name],[username],[password],[AccountType],[superuser],[administrativerights],[clientmaster],[employeemaster],[useraccounts],[maintenancerights],[payrollrights],[employeeinforights],[reportrights],[TimeKeeping],[createdby],[createddate]) values ('" & TxtOperator & "','" & txtusername & "','" & txtpassword & "','" & cmbType.Text & "'," & a & "," & b & "," & c & "," & D & "," & e & "," & f & "," & g & "," & h & "," & i & "," & j & ",'Bernard','" & Date & "')", cnnstr)
            Else
                MsgBox "Duplicate operator name/username found. Pls. check the ", vbCritical
                Exit Sub
            End If
        ElseIf save = 2 Then 'FOR EDIT
            If RsUsers.EOF = False Then
                Call Action_Query("Update usersList set [operator_name]= '" & TxtOperator & "',[username]='" & txtusername & "',[password]='" & txtpassword & "', [AccountType] = '" & cmbType.Text & "',[superuser]=" & a & ",[administrativerights]=" & b & ",[clientmaster]=" & c & ",[employeemaster]=" & D & ",[useraccounts]=" & e & ",[maintenancerights]=" & f & ",[payrollrights]=" & g & ",[employeeinforights]=" & h & ",[reportrights]=" & i & ",[TimeKeeping]=" & j & ",[editedby]='Bernard',[editeddate]='" & Date & "' where operatorid=" & operatorID & "", cnnstr)
            End If
        End If
        Call Form_Load
        MsgBox "User account was updated!", vbInformation, PayrollAlert
        Call disabledtext
        Call ValidateButton(Me, True, True, False, False, False, True)
        List1.Enabled = True
       
    Else
        MsgBox "Passwords do not match.", vbOKOnly + vbInformation, "Edit"
        txtConfirm.SetFocus
        'cmdSave.Enabled = True
    End If

End Sub

Public Function CheckText() As Boolean
    If txtusername = "" Or txtpassword = "" Or txtConfirm = "" Then
        MsgBox "You must fill in all fields", vbCritical
        GoTo exitFunct
    End If
    usertxt = txtusername Like "* *"
    If usertxt Then
        MsgBox "Invalid username.", vbCritical
        GoTo exitFunct
    End If
    Passwordtxt = txtpassword Like "* *"
    If Passwordtxt Then
        MsgBox "Invalid password.", vbCritical
        GoTo exitFunct
    End If
    CheckText = True
    Exit Function
exitFunct:
    CheckText = False
End Function

Private Sub cmbType_GotFocus()
cmbType.BackColor = &HFFFFC0
End Sub

Private Sub cmbType_LostFocus()
cmbType.BackColor = vbWhite
Call UpdateRights
End Sub

Private Sub UpdateRights()
LV.ListItems(1).Checked = False
LV.ListItems(2).Checked = False
LV.ListItems(3).Checked = False
LV.ListItems(4).Checked = False  ' Employee Master
LV.ListItems(5).Checked = False  ' Employee Master
LV.ListItems(6).Checked = False  ' Maintenance
LV.ListItems(7).Checked = False  ' Payroll
LV.ListItems(8).Checked = False  ' Employee Rights
LV.ListItems(9).Checked = False  ' Reports
LV.ListItems(10).Checked = False ' Time Keep

If UCase(Left(cmbType.Text, 3)) = "ADM" Then
    LV.ListItems(1).Checked = True
    LV.ListItems(2).Checked = True
    LV.ListItems(3).Checked = True
    LV.ListItems(4).Checked = True  ' Employee Master
    LV.ListItems(5).Checked = True  ' Employee Master
    LV.ListItems(6).Checked = True  ' Maintenance
    LV.ListItems(7).Checked = True  ' Payroll
    LV.ListItems(8).Checked = True  ' Employee Rights
    LV.ListItems(9).Checked = True  ' Reports
    LV.ListItems(10).Checked = True ' Time Keep
ElseIf cmbType.Text = "FMS Timekeep" Then
    LV.ListItems(10).Checked = True
    LV.ListItems(8).Checked = True  ' Employee Rights
ElseIf cmbType.Text = "FMS Payroll" Then
    LV.ListItems(7).Checked = True
    LV.ListItems(8).Checked = True  ' Employee Rights
ElseIf cmbType.Text = "FMS Admin" Then
    LV.ListItems(10).Checked = True ' Time Keep
    LV.ListItems(7).Checked = True  ' Payroll
    LV.ListItems(8).Checked = True  ' Employee Rights
ElseIf cmbType.Text = "Staff Timekeep" Then
    LV.ListItems(10).Checked = True
    LV.ListItems(8).Checked = True  ' Employee Rights
ElseIf cmbType.Text = "Staff Payroll" Then
    LV.ListItems(7).Checked = True
    LV.ListItems(8).Checked = True  ' Employee Rights
ElseIf cmbType.Text = "Staff Admin" Then
    LV.ListItems(10).Checked = True ' Time Keep
    LV.ListItems(7).Checked = True  ' Payroll
    LV.ListItems(8).Checked = True  ' Employee Rights
ElseIf cmbType.Text = "Accounting" Then
    LV.ListItems(10).Checked = True ' Time Keep
    LV.ListItems(7).Checked = True  ' Payroll
    LV.ListItems(8).Checked = True  ' Employee Rights
    LV.ListItems(6).Checked = True  ' Maintenance
    LV.ListItems(9).Checked = True  ' Reports
    LV.ListItems(4).Checked = True  ' Employee Master
End If


'           If .Fields("superuser") = True Then LV.ListItems(1).Checked = True Else LV.ListItems(1).Checked = False
'            If .Fields("administrativerights") = True Then LV.ListItems(2).Checked = True Else LV.ListItems(2).Checked = False
'            If .Fields("clientmaster") = True Then LV.ListItems(3).Checked = True Else LV.ListItems(3).Checked = False
'            If .Fields("employeemaster") = True Then LV.ListItems(4).Checked = True Else LV.ListItems(4).Checked = False
'            If .Fields("useraccounts") = True Then LV.ListItems(5).Checked = True Else LV.ListItems(5).Checked = False
'            ' Maintenance Rights
'            If .Fields("maintenancerights") = True Then LV.ListItems(6).Checked = True Else LV.ListItems(6).Checked = False
'
'         If .Fields("payrollrights") = True Then LV.ListItems(7).Checked = True Else LV.ListItems(7).Checked = False
'         If .Fields("employeeinforights") = True Then LV.ListItems(8).Checked = True Else LV.ListItems(8).Checked = False
'         If .Fields("reportrights") = True Then LV.ListItems(9).Checked = True Else LV.ListItems(9).Checked = False
'         If .Fields("timekeeping") = True Then LV.ListItems(10).Checked = True Else LV.ListItems(10).Checked = False

End Sub


Private Sub cmdAdd_Click()
   Call ValidateButton(Me, False, False, True, True, False, False)
   LV.Enabled = True
   List1.Enabled = False
   save = 1
   Call enabledtext
   Call clearText
   
   SSTab1.Tab = 0
  TxtOperator.SetFocus
   
End Sub

Private Sub cmdCancel_Click()
    List1.Enabled = True
    Call disabledtext
    Call Form_Load
End Sub

Private Sub cmdDelete_Click()

If DeleteRights = False Then
    MsgBox "You do not have rights to Delete Record!", vbInformation, PayrollAlert
    Exit Sub
End If
Call Delete

End Sub

Private Sub cmdEdit_Click()
    Call ValidateButton(Me, False, False, True, True, False, False)
    LV.Enabled = True
    List1.Enabled = False
        
    save = 2
    List1.Enabled = False
    Call enabledtext
End Sub

Private Sub cmdRefresh_Click()
    Call LoadUsers
Call ValidateButton(Me, True, True, False, False, False, True)
    LV.Enabled = True
    LV.SelectedItem.SELECTED = True

End Sub

Private Sub cmdSave_Click()
Call SaveUpdate
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Call ValidateButton(Me, True, True, False, False, False, False)

LV.ListItems.Clear
Dim litm As ListItem

 Set litm = LV.ListItems.Add(, , "Superuser", , 1)
 Set litm = LV.ListItems.Add(, , "Administrative Rights", , 1)
 Set litm = LV.ListItems.Add(, , "      Client Master", , 1)
 Set litm = LV.ListItems.Add(, , "      Employee Master", , 1)
 Set litm = LV.ListItems.Add(, , "      User Accounts", , 1)
 Set litm = LV.ListItems.Add(, , "      Table Maintenance", , 1) ' Connection Settings
 
 Set litm = LV.ListItems.Add(, , "Payroll Rights ", , 1)
 Set litm = LV.ListItems.Add(, , "Employee Info. Rights", , 1)
 Set litm = LV.ListItems.Add(, , "HR Report Rights", , 1)
 Set litm = LV.ListItems.Add(, , "Timekeeping Rights", , 1)    ' Maintenance Settings

Call LoadUsers
List1.Enabled = True
If List1.ListCount >= 1 Then
List1.ListIndex = 0
Else
cmddelete.Enabled = False
cmdedit.Enabled = False
End If

End Sub

Private Sub LoadUsers()
Dim RsUsers As ADODB.Recordset
 
Set RsUsers = New ADODB.Recordset
List1.Clear
StrQuery = "select * from userslist order by operator_name"
Set RsUsers = ReturnRst(StrQuery, cnnstr, mstatic)
'RsUsers.Open "select * from userslist order by operator_name", cnnstr, adOpenDynamic, adLockOptimistic
   
While Not RsUsers.EOF
    List1.AddItem RsUsers.Fields("operator_name")
    RsUsers.MoveNext
Wend
End Sub

Private Sub List1_Click()
Dim RsUsers As ADODB.Recordset
 
On Error GoTo ErrorRtn2

Set RsUsers = New ADODB.Recordset
 
StrQuery = "select * from userslist where Operator_name='" & List1.Text & "'"
Set RsUsers = ReturnRst(StrQuery, cnnstr, mstatic)
'RsUsers.Open "select * from userslist where Operator_name='" & List1.Text & "'", cnnstr, adOpenDynamic, adLockOptimistic
With RsUsers
   
If RsUsers.EOF = False Then
            TxtOperator = IIf(IsNull(RsUsers.Fields("operator_name")) = True, "", RsUsers.Fields("Operator_name"))
            txtusername = RsUsers.Fields("username")
            txtpassword = RsUsers.Fields("password")
            txtConfirm = RsUsers.Fields("password")
            cmbType.Text = RsUsers.Fields("AccountType")
            'OperatorID = RsUsers.Fields("operatorID")
            
            If .Fields("superuser") = True Then LV.ListItems(1).Checked = True Else LV.ListItems(1).Checked = False
            If .Fields("administrativerights") = True Then LV.ListItems(2).Checked = True Else LV.ListItems(2).Checked = False
            If .Fields("clientmaster") = True Then LV.ListItems(3).Checked = True Else LV.ListItems(3).Checked = False
            If .Fields("employeemaster") = True Then LV.ListItems(4).Checked = True Else LV.ListItems(4).Checked = False
            If .Fields("useraccounts") = True Then LV.ListItems(5).Checked = True Else LV.ListItems(5).Checked = False
            ' Maintenance Rights
            If .Fields("maintenancerights") = True Then LV.ListItems(6).Checked = True Else LV.ListItems(6).Checked = False
           
         If .Fields("payrollrights") = True Then LV.ListItems(7).Checked = True Else LV.ListItems(7).Checked = False
         If .Fields("employeeinforights") = True Then LV.ListItems(8).Checked = True Else LV.ListItems(8).Checked = False
         If .Fields("reportrights") = True Then LV.ListItems(9).Checked = True Else LV.ListItems(9).Checked = False
         If .Fields("timekeeping") = True Then LV.ListItems(10).Checked = True Else LV.ListItems(10).Checked = False
         
         Call disabledtext
             operatorID = .Fields!operatorID
End If
End With
cmddelete.Enabled = True

Exit Sub

ErrorRtn2:

End Sub

Public Sub clearText()
    txtpassword = ""
    txtusername = ""
    txtConfirm = ""
    TxtOperator = ""
    
    cmbType.Text = "Accounting"
    Dim X As Integer
    X = 1
    Do While X <= LV.ListItems.Count
    LV.ListItems(X).Checked = False
    X = X + 1
    Loop
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim X As Integer
    X = 1
    If Item.Index = 1 Then
        If LV.ListItems(1).Checked = True Then
            Do While X <= LV.ListItems.Count
            LV.ListItems(X).Checked = True
            X = X + 1
            Loop
        Else
            Do While X <= LV.ListItems.Count
            LV.ListItems(X).Checked = False
            X = X + 1
            Loop
        End If
    ElseIf Item.Index = 2 Then
        If LV.ListItems(2).Checked = True Then
            X = 3
            Do While X <= 6
            LV.ListItems(X).Checked = True
            X = X + 1
            Loop
        Else
            X = 3
            Do While X <= 6
            LV.ListItems(X).Checked = False
            X = X + 1
            Loop
            
        End If
    
    Else
        If LV.ListItems(Item.Index).Checked = False Then
        LV.ListItems(1).Checked = False
        End If
        
    End If
End Sub

Private Sub txtConfirm_GotFocus()
txtConfirm.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtConfirm
End Sub

Private Sub txtConfirm_LostFocus()
txtConfirm.BackColor = vbWhite
End Sub

Private Sub TxtOperator_GotFocus()
TxtOperator.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = TxtOperator
End Sub

Private Sub TxtOperator_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub TxtOperator_LostFocus()
TxtOperator.BackColor = vbWhite
End Sub

Private Sub txtpassword_GotFocus()
txtpassword.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtpassword
End Sub

Private Sub txtpassword_LostFocus()
txtpassword.BackColor = vbWhite
End Sub

Private Sub txtusername_GotFocus()
txtusername.BackColor = &HFFFFC0
Set oCommonFunction.ValidAlpha = txtusername
End Sub

Private Sub txtusername_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtusername_LostFocus()
txtusername.BackColor = vbWhite
End Sub

