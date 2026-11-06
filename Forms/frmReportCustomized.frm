VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmReportCustomized 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Asia Report Manager"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   Icon            =   "frmReportHired.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   9465
   Begin VB.Frame FraCutOff 
      BackColor       =   &H00E0E0E0&
      Height          =   1815
      Left            =   2655
      TabIndex        =   0
      Top             =   1890
      Visible         =   0   'False
      Width           =   4200
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Close"
         Height          =   285
         Left            =   1125
         TabIndex        =   5
         Top             =   1350
         Width           =   915
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Height          =   285
         Left            =   2430
         TabIndex        =   4
         Top             =   1350
         Width           =   915
      End
      Begin VB.ComboBox cboCutOff 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmReportHired.frx":058A
         Left            =   495
         List            =   "frmReportHired.frx":058C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   810
         Width           =   3345
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "The DTR report will be based from the selected cutoff"
         Height          =   330
         Left            =   180
         TabIndex        =   2
         Top             =   225
         Width           =   3885
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select here:"
         Height          =   420
         Left            =   180
         TabIndex        =   3
         Top             =   585
         Width           =   1140
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4860
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   315
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   8573
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   -2147483638
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
      TabCaption(0)   =   "  "
      TabPicture(0)   =   "frmReportHired.frx":058E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSort(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSTab2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdShow"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtTitle"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   3060
         TabIndex        =   39
         Top             =   3870
         Width           =   4515
         Begin VB.CheckBox chkSort 
            Caption         =   "Descending"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   2835
            TabIndex        =   42
            Top             =   270
            Width           =   1230
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Sort by"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   41
            Top             =   270
            Width           =   1230
         End
         Begin VB.ComboBox cboSort 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   1080
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   225
            Width           =   1755
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3030
         Left            =   3060
         TabIndex        =   11
         Top             =   855
         Width           =   5865
         Begin VB.ComboBox cboEmpStatus 
            Height          =   315
            Left            =   3330
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1080
            Visible         =   0   'False
            Width           =   2445
         End
         Begin VB.TextBox lblCon 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   3
            Left            =   2295
            TabIndex        =   33
            Text            =   "Equal to"
            Top             =   1575
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.TextBox lblCon 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   2
            Left            =   2295
            TabIndex        =   32
            Text            =   "Equal to"
            Top             =   1170
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.TextBox lblCon 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   1
            Left            =   2295
            TabIndex        =   31
            Text            =   "Equal to"
            Top             =   765
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.ComboBox cboFilterby 
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            ItemData        =   "frmReportHired.frx":05AA
            Left            =   360
            List            =   "frmReportHired.frx":05AC
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1485
            Width           =   1950
         End
         Begin VB.ComboBox cboFilterby 
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            ItemData        =   "frmReportHired.frx":05AE
            Left            =   360
            List            =   "frmReportHired.frx":05B0
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1080
            Width           =   1950
         End
         Begin VB.ComboBox cboFilterby 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            ItemData        =   "frmReportHired.frx":05B2
            Left            =   360
            List            =   "frmReportHired.frx":05B4
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   675
            Width           =   1950
         End
         Begin VB.TextBox lblCon 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   0
            Left            =   2295
            TabIndex        =   24
            Text            =   "Equal to"
            Top             =   315
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.ComboBox cboFilterby 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            ItemData        =   "frmReportHired.frx":05B6
            Left            =   360
            List            =   "frmReportHired.frx":05B8
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   270
            Width           =   1950
         End
         Begin VB.ComboBox CboCompany 
            Height          =   315
            Left            =   3330
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   270
            Visible         =   0   'False
            Width           =   2445
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Contractual"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   20
            Top             =   2385
            Width           =   1590
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Project Base"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   19
            Top             =   2115
            Width           =   1590
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Probationary"
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   18
            Top             =   2655
            Width           =   1590
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Regular"
            Height          =   255
            Index           =   3
            Left            =   1980
            TabIndex        =   17
            Top             =   2115
            Width           =   1590
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Terminated"
            Height          =   255
            Index           =   4
            Left            =   1980
            TabIndex        =   16
            Top             =   2385
            Width           =   1590
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Resigned"
            Height          =   255
            Index           =   5
            Left            =   1980
            TabIndex        =   15
            Top             =   2655
            Width           =   1590
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Early EOC"
            Height          =   255
            Index           =   6
            Left            =   3690
            TabIndex        =   14
            Top             =   2115
            Width           =   1590
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "EOC"
            Height          =   255
            Index           =   7
            Left            =   3690
            TabIndex        =   13
            Top             =   2385
            Width           =   1590
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "AWOL"
            Height          =   255
            Index           =   8
            Left            =   3690
            TabIndex        =   12
            Top             =   2655
            Width           =   1590
         End
         Begin MSMask.MaskEdBox mskDate1 
            Height          =   285
            Index           =   0
            Left            =   3330
            TabIndex        =   35
            Top             =   675
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "mm/dd/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDate2 
            Height          =   285
            Index           =   0
            Left            =   3330
            TabIndex        =   36
            Top             =   1485
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   4210752
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "mm/dd/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDate1 
            Height          =   285
            Index           =   1
            Left            =   4590
            TabIndex        =   37
            Top             =   675
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   4210752
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "mm/dd/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDate2 
            Height          =   285
            Index           =   1
            Left            =   4590
            TabIndex        =   38
            Top             =   1485
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   4210752
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "mm/dd/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.CheckBox Check1 
            Caption         =   "(1) Filter by"
            Height          =   285
            Index           =   0
            Left            =   135
            TabIndex        =   23
            Top             =   270
            Width           =   1185
         End
         Begin VB.CheckBox Check1 
            Caption         =   "(3) Filter by"
            Height          =   285
            Index           =   1
            Left            =   135
            TabIndex        =   28
            Top             =   675
            Width           =   1230
         End
         Begin VB.CheckBox Check1 
            Caption         =   "(2) Filter by"
            Height          =   285
            Index           =   2
            Left            =   135
            TabIndex        =   29
            Top             =   1080
            Width           =   1230
         End
         Begin VB.CheckBox Check1 
            Caption         =   "(4) Filter by"
            Height          =   285
            Index           =   3
            Left            =   135
            TabIndex        =   30
            Top             =   1485
            Width           =   1230
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   5850
            Y1              =   1935
            Y2              =   1935
         End
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4275
         TabIndex        =   10
         Text            =   "S E A R C H   A S I A   R E P O R T "
         Top             =   585
         Width           =   4650
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "Sho&w Report"
         Height          =   315
         Left            =   7695
         TabIndex        =   9
         Top             =   4095
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "SELECT NAME"
         Height          =   420
         Left            =   3195
         TabIndex        =   7
         Top             =   4545
         Visible         =   0   'False
         Width           =   5865
         Begin VB.ComboBox CboName 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   990
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   270
            Width           =   4455
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4470
         Left            =   225
         TabIndex        =   43
         Top             =   180
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   7885
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Report List"
         TabPicture(0)   =   "frmReportHired.frx":05BA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstVw"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Selection"
         TabPicture(1)   =   "frmReportHired.frx":05D6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Lv"
         Tab(1).ControlCount=   1
         Begin VB.ListBox lstVw 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3900
            ItemData        =   "frmReportHired.frx":05F2
            Left            =   135
            List            =   "frmReportHired.frx":05F4
            TabIndex        =   44
            Top             =   450
            Width           =   2370
         End
         Begin MSComctlLib.ListView Lv 
            Height          =   3615
            Left            =   -74910
            TabIndex        =   45
            Top             =   450
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   6376
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
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
               Object.Width           =   3528
            EndProperty
         End
      End
      Begin VB.Label lblSort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Title :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3060
         TabIndex        =   46
         Top             =   630
         Width           =   1140
      End
   End
End
Attribute VB_Name = "FrmReportCustomized"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim i As Integer
Public rsItems As ADODB.Recordset
Public SELECT1, SELECT2, SELECT3, SELECT4, SELECT5, SELECTED, whereCon As String
Public SELECT1TITLE, SELECT2TITLE, SELECT3TITLE, SELECT4TITLE, SELECT5TITLE As String

Public Function SearchFilterby(ByVal ndx As Integer) As Boolean
    Dim X As Integer
    For X = 0 To cboFilterby.Count - 1
    If ndx <> X And cboFilterby(ndx).Text <> "" Then
     If cboFilterby(ndx).Text = cboFilterby(X).Text Then cboFilterby(ndx).ListIndex = -1: SearchFilterby = False: Exit Function
    End If
    Next X
    SearchFilterby = True
End Function

Private Sub cboFilterby_Click(Index As Integer)
    If SearchFilterby(Index) = False Then MsgBox "Filter field is already selected.", vbInformation: Exit Sub
    Hidefield (Index)
    If cboFilterby(Index).ListIndex = 0 Then 'CLIENT NAME
             CboCompany.Clear
             Dim rstclient As ADODB.Recordset
             Set rstclient = cnDCPI.Execute("select distinct clientname,clientcode from clienttable  WHERE clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' ORDER BY clientname")
             Do While Not rstclient.EOF
                    CboCompany.AddItem rstclient!ClientName
                      CboCompany.ItemData(CboCompany.NewIndex) = rstclient!ClientCode
                    rstclient.MoveNext
             Loop
             rstclient.Close
             CboCompany.Top = cboFilterby(Index).Top: lblCon(Index).Visible = True: CboCompany.Visible = True
             lblCon(Index).Visible = True: lblCon(Index).Text = "equal"
             
    ElseIf cboFilterby(Index).ListIndex = 1 Then 'DATEHIRED
         mskDate1(0).Visible = True: mskDate1(0).Top = cboFilterby(Index).Top
         mskDate1(1).Visible = True: mskDate1(1).Top = cboFilterby(Index).Top
         lblCon(Index).Visible = True: lblCon(Index).Text = "between"
    
    ElseIf cboFilterby(Index).ListIndex = 2 Then 'END CONTRACT
         mskDate2(0).Visible = True: mskDate2(0).Top = cboFilterby(Index).Top
         mskDate2(1).Visible = True: mskDate2(1).Top = cboFilterby(Index).Top
         lblCon(Index).Visible = True: lblCon(Index).Text = "between"
    
    ElseIf cboFilterby(Index).ListIndex = 3 Then 'STATUS
         Me.cboEmpStatus.Clear
       
         Dim rstStatus As ADODB.Recordset
         Set rstStatus = cnDCPI.Execute("select distinct TaxExemptioncode,DESCRIPTION from TaxExemptioncode")
         Do While Not rstStatus.EOF
                cboEmpStatus.AddItem rstStatus!TaxExemptioncode & "   - " & rstStatus!Description
                 ' cboEmpStatus.ItemData(cboEmpStatus.NewIndex) = rstClient!employmentstatus
                rstStatus.MoveNext
         Loop
         rstStatus.Close
         cboEmpStatus.Top = cboFilterby(Index).Top: lblCon(Index).Visible = True: cboEmpStatus.Visible = True
         lblCon(Index).Visible = True: lblCon(Index).Text = "equal"
    End If
End Sub
Private Function Hidefield(ByVal ndx As Integer)
    If mskDate1(0).Visible = True And mskDate1(0).Top = cboFilterby(ndx).Top Then mskDate1(0).Visible = False
    If mskDate1(1).Visible = True And mskDate1(1).Top = cboFilterby(ndx).Top Then mskDate1(1).Visible = False
    If mskDate2(0).Visible = True And mskDate2(0).Top = cboFilterby(ndx).Top Then mskDate2(0).Visible = False
    If mskDate2(1).Visible = True And mskDate2(1).Top = cboFilterby(ndx).Top Then mskDate2(1).Visible = False
    If CboCompany.Visible = True And CboCompany.Top = cboFilterby(ndx).Top Then CboCompany.Visible = False
    If cboEmpStatus.Visible = True And cboEmpStatus.Top = cboFilterby(ndx).Top Then cboEmpStatus.Visible = False
End Function
Private Sub HIDEALL()
    mskDate1(0).Visible = False
    mskDate1(1).Visible = False
    mskDate2(0).Visible = False
    mskDate2(1).Visible = False
    Dim X As Integer
    For X = 0 To lblCon.Count - 1
    lblCon(X).Visible = False
    Next X
    CboCompany.Visible = False
    cboEmpStatus.Visible = False
End Sub

Private Sub CboName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    Call FindMatch(CboName, KeyAscii)
End Sub

Private Sub Check1_Click(Index As Integer)
    If cboFilterby(Index).Enabled = True Then
        cboFilterby(Index).Enabled = False
        cboFilterby(Index).BackColor = &H8000000A
        If Index = 0 Then CHECKINDEX (Index)
        If Index = 1 Then CHECKINDEX (Index)
        If Index = 2 Then CHECKINDEX (Index)
        If Index = 3 Then CHECKINDEX (Index)
        cboFilterby(Index).ListIndex = -1
    Else
        If Check1(Index).Value = 1 Then cboFilterby(Index).Enabled = True
        If Check1(Index).Value = 1 Then cboFilterby(Index).BackColor = vbWhite
    End If
End Sub

Private Sub CHECKINDEX(ByVal Index As Integer)
    If cboFilterby(Index).ListIndex = 0 Then
        Me.CboCompany.Visible = False: lblCon(Index).Visible = False
    ElseIf cboFilterby(Index).ListIndex = 1 Then
        Me.mskDate1(0).Visible = False: Me.mskDate1(1).Visible = False: lblCon(Index).Visible = False
    ElseIf cboFilterby(Index).ListIndex = 2 Then
        Me.mskDate2(0).Visible = False: Me.mskDate2(1).Visible = False: lblCon(Index).Visible = False
    ElseIf cboFilterby(Index).ListIndex = 3 Then
        Me.cboEmpStatus.Visible = False: lblCon(Index).Visible = False
    End If
End Sub

Private Sub Check2_Click()
   If cboSort(0).Enabled = True Then
        cboSort(0).Enabled = False
        cboSort(0).BackColor = &H8000000A
        cboSort(0).ListIndex = -1
        chkSort(0).Enabled = False
    Else
        cboSort(0).Enabled = True
        cboSort(0).BackColor = vbWhite
        chkSort(0).Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
FraCutOff.Visible = False
  Me.SSTab1(0).Enabled = True
  
End Sub

Private Sub cmdOK_Click()
    If cboCutOff.ListIndex = -1 Then
        MsgBox "Please select cutoff for this report.", vbInformation
    Else
        Call CREATEdtrREPORT
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = UCase(CompanyName) & " REPORT MANAGER"
    txtTitle.Text = UCase(CompanyName)
    Me.Top = (mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Dim rsAssgnList As ADODB.Recordset
    Call LoadListReport
    Set rsAssgnList = New ADODB.Recordset
       sqlString = "SELECT FIELDNAME,FIELDOBJECT1 FROM tblCustomized"
       rsAssgnList.Open sqlString, Cnnstr, adOpenDynamic, adLockOptimistic
       PopulateList Lv, rsAssgnList
       Lv.ColumnHeaders(1).Width = 2000
       Lv.ColumnHeaders(2).Width = 0
       Lv.ListItems.Item(1).Checked = True
       Lv.ListItems.Item(2).Checked = True
     Dim ndx As Integer
     For ndx = 0 To cboFilterby.Count - 1
       cboFilterby(ndx).BackColor = &H8000000A
     Next ndx
End Sub

Private Sub cboSort_Click(Index As Integer)
  chkSort(Index).Enabled = cboSort(Index).ListIndex > 0
End Sub

Private Sub cboSort_GotFocus(Index As Integer)
    cboSort(0).Clear
    If SSTab2.Tab = 0 Then
        With cboSort(0)
        .AddItem "(None)"
        If SELECT1 = "" Then
           SELECT1 = "controlno"
        End If
        .AddItem SELECT1
        .ItemData(.NewIndex) = 1
        If SELECT2 = "" Then
             SELECT2 = "Empname"
        End If
         .AddItem SELECT2
        .ItemData(.NewIndex) = 2
        If SELECT3 <> "" Then
             .AddItem SELECT3
             .ItemData(.NewIndex) = 3
        End If
        If SELECT4 <> "" Then
             .AddItem SELECT4
             .ItemData(.NewIndex) = 4
        End If
          
        If SELECT5 <> "" Then
             .AddItem SELECT5
             .ItemData(.NewIndex) = 5
        End If
        End With
    Else
        Call CheckField
        For i = 0 To cboSort.Count - 1
            With cboSort(i)
              .AddItem "(None)"
              .AddItem SELECT1
             .ItemData(.NewIndex) = 1
              .AddItem SELECT2
             .ItemData(.NewIndex) = 2
              .AddItem SELECT3
              .ItemData(.NewIndex) = 3
              .AddItem SELECT4
              .ItemData(.NewIndex) = 4
              .AddItem SELECT5
              .ItemData(.NewIndex) = 5
              .ListIndex = 0
            End With
          Next
    End If
End Sub

'section 1 is detail
'section2 is header

Public Sub LoadListReport()
    lstVw.Clear
    lstVw.AddItem "Information Sheet"
    lstVw.AddItem "Employee List(Company,Branch,Dep't)"
    lstVw.AddItem "List of Pagibig Members"
    lstVw.AddItem "List of SSS Members"
    lstVw.AddItem "List of Philhealth Members"
    lstVw.AddItem "List of TIN Members"
    lstVw.AddItem "End of Contract Report"
    lstVw.AddItem "Resigned Employees"
    lstVw.AddItem "Employees Account nos."
    lstVw.AddItem "Employees birthday"
    lstVw.AddItem "Employees Rate"
    lstVw.AddItem "Tax Exception"
    lstVw.AddItem "Masterlist Excel Report "
    lstVw.AddItem "DTR Summary"
    lstVw.AddItem "3 months summary"
    lstVw.AddItem "Employees Count by client"
    ' lstVw.AddItem "Employment history"

    For i = 0 To cboFilterby.Count - 1
        cboFilterby(i).AddItem "Client name"
        cboFilterby(i).AddItem "Start of contract"
        cboFilterby(i).AddItem "End of contract"
        cboFilterby(i).AddItem "Civil status"
      Next
End Sub
Private Function WhereConn() As String
    whereCon = ""
    If CboCompany.Visible = True Then
        whereCon = " clientname= '" & CboCompany.Text & "'"
    Else
    End If
    
    If cboEmpStatus.Visible = True Then
        If whereCon = "" Then
            whereCon = "civilstatus= '" & Mid(cboEmpStatus.Text, 1, 3) & "'"
        Else
            whereCon = whereCon & " and civilstatus ='" & Trim(Mid(cboEmpStatus.Text, 1, 3)) & "'"
        End If
    End If

    If mskDate1(0).Visible = True And mskDate1(1).Visible = True Then
        If whereCon = "" Then
            whereCon = "  startcontract between  #" & mskDate1(0).Text & "# and #" & mskDate1(1).Text & "#"
        Else
            whereCon = whereCon & " and  startcontract between  #" & mskDate1(0).Text & "# and #" & mskDate1(1).Text & "#"
        End If
    End If
  
 If mskDate2(0).Visible = True And mskDate2(1).Visible = True Then
    If whereCon = "" Then
    whereCon = " endcontract between  #" & mskDate2(0).Text & "# and #" & mskDate2(1).Text & "#"
    Else
    whereCon = whereCon & " and endcontract between  #" & mskDate2(0).Text & "# and #" & mskDate2(1).Text & "#"
    End If
 End If
    
    Dim str As String
    str = ""
    If chkStatus(0).Value = 1 Then
    str = "'B'"
    End If
    
    If chkStatus(1).Value = 1 Then
    If str <> "" Then str = str & ",'C'" Else str = "'C'"
    End If
    
    If chkStatus(2).Value = 1 Then
    If str <> "" Then str = str & ",'P'" Else str = "'P'"
    End If
    
    If chkStatus(3).Value = 1 Then
    If str <> "" Then str = str & ",'R'" Else str = "'R'"
    End If
    
    If chkStatus(4).Value = 1 Then
    If str <> "" Then str = str & ",'T'" Else str = "'T'"
    End If
    
    
    If chkStatus(5).Value = 1 Then
    If str <> "" Then str = str & ",'X'" Else str = "'X'"
    End If
    
    If chkStatus(6).Value = 1 Then
    If str <> "" Then str = str & ",'Y'" Else str = "'Y'"
    End If
    
    If chkStatus(7).Value = 1 Then
        If str <> "" Then str = str & ",'Z'" Else str = "'Z'"
    End If
    
    If chkStatus(8).Value = 1 Then
        If str <> "" Then str = str & ",'W'" Else str = "'W'"
    End If
    
    If lstVw.ListIndex <> 7 Then
    If str = "" Then MsgBox "Pls select employment status.", vbInformation: whereCon = "EXIT": Exit Function
    End If
    
    If whereCon <> "" Then
        whereCon = whereCon & " AND HREmployees.EMPLOYMENTSTATUS IN " & "(" & str & ")"
    Else
        whereCon = " HREmployees.EMPLOYMENTSTATUS IN " & "(" & str & ")"
    End If
    
    If whereCon <> "" Then: WhereConn = " where " & whereCon: whereCon = " where " & whereCon
End Function

Public Function CheckFieldDAte() As Boolean
    If mskDate1(0).Visible = True And mskDate1(0).Text = "__/__/____" Then GoTo exx
    If mskDate1(1).Visible = True And mskDate1(1).Text = "__/__/____" Then GoTo exx
    If mskDate2(0).Visible = True And mskDate2(0).Text = "__/__/____" Then GoTo exx
    If mskDate2(1).Visible = True And mskDate2(1).Text = "__/__/____" Then GoTo exx
    CheckFieldDAte = True
    Exit Function
exx:
    MsgBox "Fill up the blank field.", vbInformation: CheckFieldDAte = False:
End Function

Private Function ChangeChkStatus(ByVal what As Boolean)
    Dim X As Integer
    For X = 0 To 8
    chkStatus(X).Enabled = what
    Next X
End Function
Private Sub cmdShow_Click()
    On Error Resume Next
    If CheckFieldDAte = False Then Exit Sub
        
    If cboFilterby(0).Enabled = False Or cboFilterby(0).ListIndex = -1 Then
        If MsgBox("Are you sure you want generate report without filtering?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
        End If
    End If
    
    Call WhereConn
    If whereCon = "EXIT" Then Exit Sub
        If SSTab2.Tab = 1 Then
            If isCheckExceed = False Then MsgBox "No. of fields selected maybe less than 5 or exceeded . 5 fields are required.", vbInformation: Exit Sub
            Call CheckField
        Else
            If lstVw.ListIndex = -1 Then MsgBox "Select type of report!", vbInformation: Exit Sub
            SELECT1 = "controlno"
            SELECT2 = "Empname"
            SELECT3 = "Empposition"
            SELECT4 = "ClientName"
            
            SELECT1TITLE = "Emp. No."
            SELECT2TITLE = "Name"
            SELECT3TITLE = "Position"
            SELECT4TITLE = "Client Name"
        
        
        If lstVw.ListIndex = 0 Then 'Information Sheet
         Dim Sortby, ByDirection As String
         ByDirection = IIf(chkSort(0).Value = 1, "Desc", "Asc")
      
      If cboSort(0).Text <> "(None)" And Trim(cboSort(0).Text) <> "" Then Sortby = " order by " & cboSort(0).Text & " " & ByDirection
            Set rsItems = ReturnRst("Select * from qryHRhiredREPORT " & whereCon & Sortby, Cnnstr, mstatic)
            If rsItems.EOF = True Then MsgBox "Record not found!", vbInformation: Exit Sub
            
              If CompanyCode = "0" Then
              With rptReportHired_SAPI
                Set .DataSource = rsItems
                 .DataMember = rsItems.DataMember
                .Sections("Section2").Controls("lblAddress").Caption = "Suite 1404, Summit One Tower" & vbCrLf & "530 Shaw Blvd., Mandaluyong City" & vbCrLf & "1552 Phillippines" & vbCrLf & "Tel. no. 02-5339984" & vbCrLf & "Fax. 02-5340082"
                 .Show
               End With
               ElseIf CompanyCode = "1" Then
              With rptReportHired_JCI
                Set .DataSource = rsItems
                 .DataMember = rsItems.DataMember
                .Sections("Section2").Controls("lblAddress").Caption = "Suite 1404, Summit One Tower" & vbCrLf & "530 Shaw Blvd., Mandaluyong City" & vbCrLf & "1552 Phillippines" & vbCrLf & "Tel. no. 02-5339984" & vbCrLf & "Fax. 02-5340082"
                 .Show
               End With
               End If
               
               
               Exit Sub
        ElseIf lstVw.ListIndex = 1 Then 'Employees List(Company,Branch,Dep't)
            SELECT3 = "ClientName"
            SELECT3TITLE = "Client Name"
            SELECT4 = "Brname"
            SELECT4TITLE = "Branch"
            SELECT5 = "Deptname"
            SELECT5TITLE = "Department"
        '    WhereCon = ""
        ElseIf lstVw.ListIndex = 2 Then 'List of Pagibig Members
            SELECT5 = "Pagibig"
            SELECT5TITLE = "Pagibig"
         '   WhereCon = ""
        ElseIf lstVw.ListIndex = 3 Then 'List of SSS Members
            SELECT5 = "SSS"
            SELECT5TITLE = "SSS"
          '  WhereCon = ""
      ElseIf lstVw.ListIndex = 4 Then 'List of Philhealth Members
            SELECT5 = "PHILHEALTH"
            SELECT5TITLE = "Philhealth"
           ' WhereCon = ""
        ElseIf lstVw.ListIndex = 5 Then 'List of TIN Members
            SELECT5 = "Taxid"
            SELECT5TITLE = "T.I.N"
            'WhereCon = ""
        ElseIf lstVw.ListIndex = 6 Then 'End of Contract Report
            SELECT3 = "ClientName"
            SELECT3TITLE = "Client Name"
            SELECT4 = "StartContract"
            SELECT4TITLE = "Start of Contract"
            SELECT5 = "Endcontract"
            SELECT5TITLE = "End of Contract"
            'WhereCon = ""
        ElseIf lstVw.ListIndex = 7 Then 'Resigned Employees
            SELECT3 = "ClientName"
            SELECT3TITLE = "Client Name"
            SELECT4 = "Dateresign"
            SELECT4TITLE = "Date resigned"
            SELECT5 = "ReasonRes"
            SELECT5TITLE = "Reason"
            chkStatus(5).Value = 1
           
            If whereCon = "" Then
             whereCon = " where employmentstatus= 'X' "
             Else
             whereCon = whereCon & " and employmentstatus= 'X' "
             End If
        ElseIf lstVw.ListIndex = 8 Then 'Employees Account nos.
            SELECT5 = "Acctno"
            SELECT5TITLE = "Account no."
            'WhereCon = ""
        ElseIf lstVw.ListIndex = 9 Then 'Employees birthday
            SELECT5 = "Birthdate"
            SELECT5TITLE = "Birthday"
           ' WhereCon = ""
         ElseIf lstVw.ListIndex = 10 Then 'Employees rate
            SELECT3 = "ClientName"
            SELECT3TITLE = "Client Name"
            SELECT4 = "Dailyrate"
            SELECT4TITLE = "Daily Rate"
            SELECT5 = "MonthlyRate"
            SELECT5TITLE = "Monthly Rate"
        ElseIf lstVw.ListIndex = 11 Then 'Tax Exception
            SELECT3 = "EmptaxCode"
            SELECT3TITLE = "Exemption code"
            SELECT4 = "Description"
            SELECT4TITLE = "Description"
            SELECT5 = "ExemptionAmount"
            SELECT5TITLE = "Exemption Amount"
        ElseIf lstVw.ListIndex = 12 Then 'Masterlist Excel
           Call CreateMasterList
           Exit Sub
        
       ElseIf lstVw.ListIndex = 13 Then 'DTR
           Me.SSTab1(0).Enabled = False
    
                 FraCutOff.Visible = True
                 Dim rstcutOff As ADODB.Recordset
                 Set rstcutOff = cnDCPI.Execute("SELECT distinct CutoffTable.StartCutoffDate,CutoffTable.StartCutoffDate &  ' - '  & CutoffTable.EndCutoffDate as CutOffDate FROM CutoffTable order by CutoffTable.StartCutoffDate")
    
                 Do While Not rstcutOff.EOF
                        Me.cboCutOff.AddItem rstcutOff!CutOffDate
                        '  CboCompany.ItemData(CboCompany.NewIndex) = rstcutOff!ClientCode
                        rstcutOff.MoveNext
                 Loop
                 
                 rstcutOff.Close
                 
            'Call CREATEdtrREPORT
            Exit Sub
        ElseIf lstVw.ListIndex = 14 Then '3 months summary
            SELECT3 = "ClientName"
            SELECT3TITLE = "Client Name"
            SELECT4 = "StartContract"
            SELECT4TITLE = "Start of Contract"
            SELECT5 = "Endcontract"
            SELECT5TITLE = "End of Contract"
            
            Dim startC, endC As String
            startC = Date - 92
            startC = month(startC) & "/1/" & Year(startC)
            endC = month(startC) & "/31/" & Year(startC)
            If IsDate(endC) = False Then
            endC = month(startC) & "/30/" & Year(startC)
            End If
                    
            If IsDate(endC) = False Then
            endC = month(startC) & "/29/" & Year(startC)
            End If
            
            If IsDate(endC) = False Then
            endC = month(startC) & "/28/" & Year(startC)
            End If
        
           If whereCon = "" Then
             whereCon = " where startcontract between # " & startC & "# and #" & endC & ""
             Else
             
             whereCon = whereCon & " and startcontract between #" & startC & "# and #" & endC & "#"
             End If
           ElseIf lstVw.ListIndex = 15 Then 'Masterlist Excel
              Call CreateEmployeesCount
              Exit Sub
            ElseIf lstVw.ListIndex = 16 Then 'Employment History
               Call CreateEmploymentHistory
            Exit Sub
           End If
        
        End If
         SELECTED = SELECT1 & "," & SELECT2 & "," & SELECT3 & "," & SELECT4 & "," & SELECT5
         If CompanyCode = "0" Then
         Call ShowReport_sapi
         ElseIf CompanyCode = "1" Then
         Call ShowReport_jci
        End If
     
End Sub
Private Function CreateEmploymentHistory()
Dim str As String
Dim rsTemp As ADODB.Recordset
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet1 As Excel.Worksheet
Dim ctr As Long
Dim i As Long
Dim Sortby, ByDirection As String
  ByDirection = IIf(chkSort(0).Value = 1, "Desc", "Asc")
    If cboSort(0).Text <> "(None)" And Trim(cboSort(0).Text) <> "" Then Sortby = " order by " & cboSort(0).Text & " " & ByDirection
      If Sortby = "" Then Sortby = " order by HREmploymentHistory.Surname"
       sqlString = "SELECT HREmploymentHistory.EmployeeNum, HREmploymentHistory.ControlNo, HREmploymentHistory.Surname+', '+HREmploymentHistory.FirstName AS EMPNAME, HREmploymentHistory.EmpPosition, HREmploymentHistory.ClientName, HREmploymentHistory.DeptName, HREmploymentHistory.BrName, HREmploymentHistory.Paymentbasis, HREmploymentHistory.MonthlyRate, HREmploymentHistory.DailyRate, HREmploymentHistory.StartContract, HREmploymentHistory.EndContract, HREmployees.EmploymentStatus" & _
                   " FROM HREmploymentHistory LEFT JOIN HREmployees ON HREmploymentHistory.EmployeeNum = HREmployees.EmployeeNum " & whereCon & Sortby
Sortby = ""
               
Set rsTemp = ReturnRst(sqlString, cnDCPI)
If rsTemp.EOF = False Then
    Call ShowEmploymentHistory(rsTemp, 4)
     Else
    MsgBox "No Records!!!", vbCritical, ""
End If
End Function



Private Sub ShowEmploymentHistory(ByRef rs As ADODB.Recordset, ByVal excelReportType As Integer)
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String
Dim sumString As String
Dim ctr As Long
Dim i As Long, j As Long
Dim cPos As Integer

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)
cPos = 7
ctr = 9
i = 1

'On Error GoTo Errh
' Print Header
'xlSheet.Range("A2") = "FOR THE PAYROLL PERIOD OF:"
'xlSheet.Range("A3") = "PAYROLL RELEASE DATE:"
xlSheet.Range("A4") = "COMPANY NAME:"

'xlSheet.Range("D2") = rs!CutoffPeriod 'Format(rs!StartPayDate, "mmm dd") & " - " & Format(rs!EndPayDate, "mmm dd yyyy")
'xlSheet.Range("D3") = rs!SalaryRelease  'Format(Date, "mmm dd yyyy")
xlSheet.Range("D4") = rs!ClientName


xlSheet.Cells(cPos - 1, 1).Value = "EmployeeNum"
xlSheet.Cells(cPos - 1, 2).Value = "EMPNAME"
xlSheet.Cells(cPos - 1, 3).Value = "CONTROLNO"
xlSheet.Cells(cPos - 1, 4).Value = "EMPPOSITION"
xlSheet.Cells(cPos - 1, 5).Value = "CLIENTNAME"
xlSheet.Cells(cPos - 1, 6).Value = "DEPT."
xlSheet.Cells(cPos - 1, 7).Value = "BRANCH"
xlSheet.Cells(cPos - 1, 8).Value = "PAYMENT BASIS"
xlSheet.Cells(cPos - 1, 9).Value = "MONTHLY RATE"
xlSheet.Cells(cPos - 1, 10).Value = "DAILY RATE"
xlSheet.Cells(cPos - 1, 11).Value = "START CONTRACT"
xlSheet.Cells(cPos - 1, 12).Value = "END CONTRACT"
cPos = cPos + 1

Do While Not rs.EOF
    'DoEvents
    xlSheet.Cells(ctr, 1).Value = rs!employeenum
    xlSheet.Cells(ctr, 2).Value = rs!empname
    xlSheet.Cells(ctr, 3).Value = rs!ControlNo
    xlSheet.Cells(ctr, 4).Value = rs!Empposition
    xlSheet.Cells(ctr, 5).Value = rs!ClientName
    xlSheet.Cells(ctr, 6).Value = rs!deptname
    xlSheet.Cells(ctr, 7).Value = rs!brname
       xlSheet.Cells(ctr, 8).Value = rs!PaymentBasis
    xlSheet.Cells(ctr, 9).Value = Format(rs!MonthlyRate, "##,##0.00")
    xlSheet.Cells(ctr, 10).Value = Format(rs!DailyRate, "##,#0.00")
    xlSheet.Cells(ctr, 11).Value = Format(rs!StartContract, "mm/dd/yyyy")
    xlSheet.Cells(ctr, 12).Value = Format(rs!endcontract, "mm/dd/yyyy")
 
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Cells.EntireRow.AutoFit
    i = i + 1
    ctr = ctr + 1
    rs.MoveNext
Loop


    xlSheet.Columns("A:A").ColumnWidth = 3.43
    xlSheet.Columns("B:B").ColumnWidth = 24.57
    xlSheet.Range("F6:I6").Select
    With xlApp1.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
   '  xlApp1.Selection.Merge
   ' xlApp1.ActiveCell.FormulaR1C1 = "DEDUCTIONS"
    Dim R
    R = "A6:L" & ctr
    xlSheet.Range("" & R & "").Select
     xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
     xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With xlApp1.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
        xlSheet.Range("A6:L7").Select
     xlApp1.Selection.Font.Bold = True
    With xlApp1.Selection.Font
        .Name = "Tahoma"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    xlSheet.Columns("L:L").ColumnWidth = 17.86
    With xlApp1.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
    End With
    xlSheet.Columns("F:I").Select
     xlApp1.Selection.ColumnWidth = 14.43
    xlSheet.Columns("E:E").ColumnWidth = 17.43
    xlSheet.Columns("D:D").ColumnWidth = 14.43
    
    xlSheet.Columns("K:K").Select
     xlApp1.Selection.ColumnWidth = 14.43
     xlApp1.Selection.ColumnWidth = 14.29
    
      xlApp1.Range("A6:L7").Select
    With xlApp1.Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
    End With
    xlApp1.Range("A2:G4").Select
    xlApp1.Selection.Font.Bold = True
    'xlApp1.Range("F9:H12").Select
    R = "F9:H" & ctr - 1
    xlApp1.Range("" & R & "").Select
    xlApp1.Selection.NumberFormat = "0.0"
    xlApp1.Selection.NumberFormat = "0.00"
    
    

    Set xlApp1 = Nothing
    Exit Sub
Errh:

    MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
    Err.Clear
     Set xlApp1 = Nothing
End Sub




Private Function CREATEdtrREPORT()

Dim str As String
Dim rsTemp As ADODB.Recordset
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet1 As Excel.Worksheet
Dim ctr As Long
Dim i As Long

Dim Sortby, ByDirection As String
  ByDirection = IIf(chkSort(0).Value = 1, "Desc", "Asc")
  
  If cboSort(0).Text <> "(None)" And Trim(cboSort(0).Text) <> "" Then Sortby = " order by " & cboSort(0).Text & " " & ByDirection
   
   If Sortby = "" Then Sortby = " order by Empname"
'SELECTED = " Empname,Controlno,Brname,Empposition,Telnum1,Telnum2,startcontract,endcontract,sss,philhealth,taxid,birthdate,civilstatus,Dailyrate,ecolaamount,employmentstatusdesc "
  'sqlString = "SELECT " & _
              "" & SELECTED & "" & _
               " From qryCustomized " & whereCon & Sortby
               sqlString = "sELECT empname,brname,cbamount,placementfee,miscfee,acctno,cutoffperiod,salaryrelease,clientname" & _
  " FROM QRYCUSTOMIZED " & whereCon & Sortby

       sqlString = "SELECT qryCustomized.*, HREmployees.CBstartpayment, HREmployees.CBendpayment, HREmployees.MFStartPayment, HREmployees.MFEndPayment, HREmployees.PFStartPayment, HREmployees.PFEndPayment" & _
                 " FROM qryCustomized LEFT JOIN HREmployees ON qryCustomized.EmployeeNum = HREmployees.EmployeeNum " & whereCon & Sortby


Sortby = ""
               
Set rsTemp = ReturnRst(sqlString, cnDCPI)
If rsTemp.EOF = False Then
    Call ShowDtrReport(rsTemp, 4)
Else
    MsgBox "No Records!!!", vbCritical, ""
End If
End Function



Private Sub ShowDtrReport(ByRef rs As ADODB.Recordset, ByVal excelReportType As Integer)
    Dim xlApp1 As Excel.Application
    Dim xlBook1 As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim mTotalPay As Double
    Dim mTotalDeduction As Double
    Dim mRange As String
    Dim sumString As String
    Dim ctr As Long
    Dim i As Long, j As Long
    Dim cPos As Integer
    
    Set xlApp1 = New Excel.Application
    Set xlBook1 = xlApp1.Workbooks.Add()
    xlApp1.Visible = True
    xlApp1.DisplayAlerts = False
    
    Set xlSheet = xlBook1.Worksheets(1)
    cPos = 7
    ctr = 9
    i = 1
    
    'On Error GoTo Errh
    ' Print Header
    xlSheet.Range("A2") = "FOR THE PAYROLL PERIOD OF:"
    xlSheet.Range("A3") = "PAYROLL RELEASE DATE:"
    xlSheet.Range("A4") = "COMPANY NAME:"
    
    xlSheet.Range("D2") = Me.cboCutOff.Text ' rs!CutoffPeriod
    xlSheet.Range("D3") = rs!SalaryRelease  'Format(Date, "mmm dd yyyy")
    xlSheet.Range("D4") = rs!ClientName
    
    'With xlSheet.Range("A2")
    '    .Font.Name = "Arial"  'Font.Name
    '    .Font.Size = 12 'font.size
    '    .Font.Bold = True ' Font.Bold
    '    .Font.Italic = False 'font.Italic
    'End With
    
    xlSheet.Cells(cPos - 1, 1).Value = "No."
    xlSheet.Cells(cPos - 1, 2).Value = "NAME"
    xlSheet.Cells(cPos - 1, 3).Value = "BRANCH"
    xlSheet.Cells(cPos - 1, 4).Value = "DTR"
    xlSheet.Cells(cPos - 1, 5).Value = "OT"
    xlSheet.Cells(cPos, 6).Value = "C/B"
    xlSheet.Cells(cPos, 7).Value = "P/F"
    xlSheet.Cells(cPos, 8).Value = "M/F"
    xlSheet.Cells(cPos, 9).Value = "C/A"
    xlSheet.Cells(cPos - 1, 10).Value = "ADD-ON"
    xlSheet.Cells(cPos, 10).Value = "OTHERS"
    xlSheet.Cells(cPos - 1, 11).Value = "ACCOUNT"
    xlSheet.Cells(cPos, 11).Value = "NUMBER"
    xlSheet.Cells(cPos, 12).Value = "REMARKS"
    cPos = cPos + 1
    
    Do While Not rs.EOF
        'DoEvents
        xlSheet.Cells(ctr, 1).Value = i
        xlSheet.Cells(ctr, 2).Value = rs!empname 'EMPLOYEE NAME
        xlSheet.Cells(ctr, 3).Value = rs!brname   'BRANCH NAME
        
        Dim ba 'variable for cut off
       ba = Split(Me.cboCutOff.Text, "-") 'split the cut off
       
       'cashbond
       
       'If (Format(ba(1), "yyyymmdd") >= Format(rs!cbstartpayment, "yyyymmdd")) And (Format(ba(1), "yyyymmdd") <= Format(rs!cbendpayment, "yyyymmdd")) Then
       If (Format(ba(1), "yyyymmdd") >= Format(rs!cbstartpayment, "yyyymmdd")) And (Format(ba(1), "yyyymmdd") <= Format(rs!cbendpayment, "yyyymmdd")) Or _
          (Format(rs!cbendpayment, "yyyymmdd") >= Format(ba(0), "yyyymmdd")) And (Format(rs!cbendpayment, "yyyymmdd") <= Format(ba(1), "yyyymmdd")) Or _
          (Format(rs!cbstartpayment, "yyyymmdd") >= Format(ba(0), "yyyymmdd")) And (Format(rs!cbstartpayment, "yyyymmdd") <= Format(ba(1), "yyyymmdd")) Then
           xlSheet.Cells(ctr, 6).Value = Format(rs!Cbamount, "##,##0.00")
       Else
           xlSheet.Cells(ctr, 6).Value = Format(0, "##,##0.00")
       End If
           
       'plcacement fee
       If (Format(ba(1), "yyyymmdd") >= Format(rs!PFstartpayment, "yyyymmdd")) And (Format(ba(1), "yyyymmdd") <= Format(rs!PFendpayment, "yyyymmdd")) Or _
          (Format(rs!PFendpayment, "yyyymmdd") >= Format(ba(0), "yyyymmdd")) And (Format(rs!PFendpayment, "yyyymmdd") <= Format(ba(1), "yyyymmdd")) Or _
          (Format(rs!PFstartpayment, "yyyymmdd") >= Format(ba(0), "yyyymmdd")) And (Format(rs!PFstartpayment, "yyyymmdd") <= Format(ba(1), "yyyymmdd")) Then
           xlSheet.Cells(ctr, 7).Value = Format(rs!PlacementFee, "##,#0.00")
       Else
           xlSheet.Cells(ctr, 7).Value = Format(0, "##,##0.00")
       End If
       
       'misc fee
       If (Format(ba(1), "yyyymmdd") >= Format(rs!MFstartpayment, "yyyymmdd")) And (Format(ba(1), "yyyymmdd") <= Format(rs!MFendpayment, "yyyymmdd")) Or _
          (Format(rs!MFendpayment, "yyyymmdd") >= Format(ba(0), "yyyymmdd")) And (Format(rs!MFendpayment, "yyyymmdd") <= Format(ba(1), "yyyymmdd")) Or _
          (Format(rs!MFstartpayment, "yyyymmdd") >= Format(ba(0), "yyyymmdd")) And (Format(rs!MFstartpayment, "yyyymmdd") <= Format(ba(1), "yyyymmdd")) Then
           xlSheet.Cells(ctr, 8).Value = Format(rs!MiscFee, "##,##0.00")
       Else
           xlSheet.Cells(ctr, 8).Value = Format(0, "##,##0.00")
       End If
         
         
        xlSheet.Cells(ctr, 11).Value = IIf(rs!AcctNo = "", "n/a", rs!AcctNo)
     
        xlSheet.Cells.EntireColumn.AutoFit
        xlSheet.Cells.EntireRow.AutoFit
        i = i + 1
        ctr = ctr + 1
        rs.MoveNext
    Loop
    
    
        xlSheet.Columns("A:A").ColumnWidth = 3.43
        xlSheet.Columns("B:B").ColumnWidth = 24.57
        xlSheet.Range("F6:I6").Select
        With xlApp1.Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = False
        End With
         xlApp1.Selection.Merge
        xlApp1.ActiveCell.FormulaR1C1 = "DEDUCTIONS"
        Dim R
        R = "A6:L" & ctr
        xlSheet.Range("" & R & "").Select
         xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
         xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With xlApp1.Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With xlApp1.Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With xlApp1.Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With xlApp1.Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With xlApp1.Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With xlApp1.Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
            xlSheet.Range("A6:L7").Select
         xlApp1.Selection.Font.Bold = True
        With xlApp1.Selection.Font
            .Name = "Tahoma"
            .Size = 10
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
        End With
        xlSheet.Columns("L:L").ColumnWidth = 17.86
        With xlApp1.Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
        End With
        xlSheet.Columns("F:I").Select
         xlApp1.Selection.ColumnWidth = 7
        xlSheet.Columns("E:E").ColumnWidth = 4.29
        xlSheet.Columns("D:D").ColumnWidth = 4.29
        
        xlSheet.Columns("K:K").Select
         xlApp1.Selection.ColumnWidth = 14.43
         xlApp1.Selection.ColumnWidth = 14.29
        
          xlApp1.Range("A6:L7").Select
        With xlApp1.Selection.Interior
            .ColorIndex = 15
            .Pattern = xlSolid
        End With
        xlApp1.Range("A2:G4").Select
        xlApp1.Selection.Font.Bold = True
        'xlApp1.Range("F9:H12").Select
        R = "F9:H" & ctr - 1
        xlApp1.Range("" & R & "").Select
        xlApp1.Selection.NumberFormat = "0.0"
        xlApp1.Selection.NumberFormat = "0.00"
        
        
    
        Set xlApp1 = Nothing
        Exit Sub
Errh:
    
        MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
        Err.Clear
         Set xlApp1 = Nothing
End Sub


Private Function CreateEmployeesCount()

Dim str As String
Dim rsTemp As ADODB.Recordset
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet1 As Excel.Worksheet
Dim ctr As Long
Dim i As Long
              
'If cboFilterby(0).Text <> "" And CboCompany.Text <> "" Then
'whereCon = " where clientname= '" & CboCompany.Text & "'"
'Else
'whereCon = ""
'End If
whereCon = WhereConn


'MsgBox WhereConn
  'WHERE   HREmployees.EmploymentStatus IN  ('C')

sqlString = "SELECT ClientTable.ClientCode, ClientTable.ClientName, HREmployees.EmploymentStatus, HREmploymentStatus.EmploymentStatusDesc, Count(HREmployees.EmploymentStatus) AS CountOfEmploymentStatus" & _
" FROM (ClientTable INNER JOIN HREmployees ON ClientTable.ClientCode = HREmployees.ClientCode) INNER JOIN HREmploymentStatus ON HREmployees.EmploymentStatus = HREmploymentStatus.EmploymentStatus " & whereCon & " " & _
" GROUP BY ClientTable.ClientCode, ClientTable.ClientName, HREmployees.EmploymentStatus , HREmploymentStatus.EmploymentStatusDesc, HREmployees.EmploymentStatus " & _
" ORDER BY ClientTable.ClientName, HREmploymentStatus.EmploymentStatusDesc;"
               
Set rsTemp = ReturnRst(sqlString, cnDCPI)
If rsTemp.EOF = False Then
    Call ShowEmployeesCount(rsTemp, 4)
  Else
    MsgBox "No Records!!!", vbCritical, ""
End If
End Function


Private Sub ShowEmployeesCount(ByRef rs As ADODB.Recordset, ByVal excelReportType As Integer)
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim mTotalPay As Double
Dim mTotalDeduction As Double
Dim mRange As String
Dim sumString As String
Dim ctr As Long
Dim i As Long, j As Long
Dim cPos As Integer

Set xlApp1 = New Excel.Application
Set xlBook1 = xlApp1.Workbooks.Add()
xlApp1.Visible = True
xlApp1.DisplayAlerts = False

Set xlSheet = xlBook1.Worksheets(1)
cPos = 6
ctr = 7
i = 1

'On Error GoTo Errh
' Print Header
xlSheet.Range("A2") = Me.txtTitle.Text
xlSheet.Range("A3") = "AGENCY :"
xlSheet.Range("A4") = "AS OF DATE : "

xlSheet.Range("D3") = CompanyName 'Format(Date, "mmmm dd, yyyy ") 'rs!SalaryRelease  'Format(Date, "mmm dd yyyy")
xlSheet.Range("D4") = "" & Format(Date, "mmmm dd, yyyy") 'rs!ClientName


xlSheet.Cells(cPos, 1).Value = "No."
xlSheet.Cells(cPos, 2).Value = "CLIENTNAME"
xlSheet.Cells(cPos, 3).Value = "CODE"
xlSheet.Cells(cPos, 4).Value = "EMPLOYMENT STATUS"
xlSheet.Cells(cPos, 5).Value = "COUNTS"
cPos = cPos + 1

Do While Not rs.EOF
    'DoEvents
    xlSheet.Cells(ctr, 1).Value = i
    xlSheet.Cells(ctr, 2).Value = rs!ClientName
    xlSheet.Cells(ctr, 3).Value = rs!employmentstatus
    xlSheet.Cells(ctr, 4).Value = rs!EmploymentStatusdesc
    xlSheet.Cells(ctr, 5).Value = rs!COUNTOFEMPLOYMENTSTATUS
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Cells.EntireRow.AutoFit
    i = i + 1
    ctr = ctr + 1
    rs.MoveNext
Loop
   
    Dim R
    R = "A6:E" & ctr
    xlSheet.Range("" & R & "").Select
     xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
     xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With xlApp1.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
        xlSheet.Range("A6:E6").Select
     xlApp1.Selection.Font.Bold = True
    With xlApp1.Selection.Font
        .Name = "Tahoma"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
      
    xlApp1.Range("A6:E6").Select
    With xlApp1.Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
    End With
    
    
        xlApp1.Columns("A:A").ColumnWidth = 3.57
    xlApp1.Range("A2:D4").Select
    xlApp1.Selection.Font.Bold = True
    With xlApp1.Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlApp1.ActiveWindow.ScrollRow = 28
    xlApp1.ActiveWindow.ScrollRow = 1
    xlApp1.Range("A6").Select
    xlApp1.Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(5), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
        
        
        
    Set xlApp1 = Nothing
    Exit Sub
Errh:

    MsgBox "Error: " & Err.Description, vbInformation, PayrollAlert
    Err.Clear
     Set xlApp1 = Nothing
End Sub

Public Sub ShowReport_benefits(ByVal valRpt As Integer)
If valRpt = 1 Then
Set rsItems = CreateRecords1("hrrptbenefits")
If rsItems.EOF = True Then MsgBox "Record not found!", vbInformation: Exit Sub
   With rptBenefits
     Set .DataSource = rsItems
     .DataMember = rsItems.DataMember
     .Sections("Section2").Controls("lblAddress").Caption = "Suite 1404, Summit One Tower" & vbCrLf & "530 Shaw Blvd., Mandaluyong City" & vbCrLf & "1552 Phillippines" & vbCrLf & "Tel. no. 02-5339984" & vbCrLf & "Fax. 02-5340082"
       If CompanyCode = 0 Then
            .Sections("Section2").Controls("image1").Visible = True
            .Sections("Section2").Controls("image2").Visible = False
       ElseIf CompanyCode = 1 Then
            .Sections("Section2").Controls("image2").Visible = True
            .Sections("Section2").Controls("image1").Visible = False
        End If
    .Sections("Section1").Controls("teXT1").DataField = SELECT1
    .Sections("Section1").Controls("Text2").DataField = SELECT2
    .Sections("Section1").Controls("Text3").DataField = SELECT3
    .Sections("Section1").Controls("Text4").DataField = SELECT4
  '  .Sections("Section1").Controls("Text5").DataField = SELECT5
    .Sections("Section2").Controls("lblname").Caption = CboName.Text
    .Sections("Section2").Controls("lblidno").Caption = GetDataName("hrrptbenefits where employeename = '" & CboName.Text & "'", "employeenum")
    .Sections("Section2").Controls("lblTitle").Caption = txtTitle
  DoEvents
 .Show
   End With
 
ElseIf valRpt = 2 Then
        Set rsItems = CreateRecords1("hremployeesassignment")
        If rsItems.EOF = True Then MsgBox "Record not found!", vbInformation: Exit Sub
           With rptAssignment
             Set .DataSource = rsItems
             .DataMember = rsItems.DataMember
             .Sections("Section2").Controls("lblAddress").Caption = "Suite 1404, Summit One Tower" & vbCrLf & "530 Shaw Blvd., Mandaluyong City" & vbCrLf & "1552 Phillippines" & vbCrLf & "Tel. no. 02-5339984" & vbCrLf & "Fax. 02-5340082"
               If CompanyCode = 0 Then
                    .Sections("Section2").Controls("image1").Visible = True
                    .Sections("Section2").Controls("image2").Visible = False
               ElseIf CompanyCode = 1 Then
                    .Sections("Section2").Controls("image2").Visible = True
                    .Sections("Section2").Controls("image1").Visible = False
                End If
            .Sections("Section1").Controls("teXT1").DataField = SELECT1
            .Sections("Section1").Controls("Text2").DataField = SELECT2
            .Sections("Section1").Controls("Text3").DataField = SELECT3
            .Sections("Section1").Controls("Text4").DataField = SELECT4
           .Sections("Section1").Controls("Text5").DataField = SELECT5
            .Sections("Section2").Controls("lblname").Caption = CboName.Text
            .Sections("Section2").Controls("lblidno").Caption = GetDataName("hrrptbenefits where employeename = '" & CboName.Text & "'", "employeenum")
            .Sections("Section2").Controls("lblTitle").Caption = txtTitle
          DoEvents
         .Show
           End With
 
End If

 
  Exit Sub
End Sub
Public Sub ShowReport_jci()
   ' Get a recordset to use with the report
   Set rsItems = CreateRecords
   If rsItems.EOF = True Then MsgBox "Record not found!", vbInformation: Exit Sub
   With rptCustomizedReport_JCI
    ' Bind the Recordset to the Report
     Set .DataSource = rsItems
    .DataMember = rsItems.DataMember
     .Sections("Section2").Controls("lblAddress").Caption = "Suite 1404, Summit One Tower" & vbCrLf & "530 Shaw Blvd., Mandaluyong City" & vbCrLf & "1552 Phillippines" & vbCrLf & "Tel. no. 02-5339984" & vbCrLf & "Fax. 02-5340082"
       
    .Sections("Section2").Controls("lblTitle").Caption = txtTitle
    .Sections("Section2").Controls("L1").Caption = SELECT1TITLE
    .Sections("Section2").Controls("L2").Caption = SELECT2TITLE
    .Sections("Section2").Controls("L3").Caption = SELECT3TITLE
    .Sections("Section2").Controls("L4").Caption = SELECT4TITLE
    .Sections("Section2").Controls("L5").Caption = SELECT5TITLE
    .Sections("Section1").Controls("Text1").DataField = SELECT1
    .Sections("Section1").Controls("Text2").DataField = SELECT2
    .Sections("Section1").Controls("Text3").DataField = SELECT3
    .Sections("Section1").Controls("Text4").DataField = SELECT4
    .Sections("Section1").Controls("Text5").DataField = SELECT5
    '.Sections("Section5").Controls("lblRecCount").Caption = CreateRecords.RecordCount & " items listed"
    DoEvents
    .Show
  End With
End Sub
Public Sub ShowReport_sapi()
   ' Get a recordset to use with the report
   Set rsItems = CreateRecords
   If rsItems.EOF = True Then MsgBox "Record not found!", vbInformation: Exit Sub
   With rptCustomizedReport_SAPI
    ' Bind the Recordset to the Report
     Set .DataSource = rsItems
    .DataMember = rsItems.DataMember
     .Sections("Section2").Controls("lblAddress").Caption = "Suite 1404, Summit One Tower" & vbCrLf & "530 Shaw Blvd., Mandaluyong City" & vbCrLf & "1552 Phillippines" & vbCrLf & "Tel. no. 02-5339984" & vbCrLf & "Fax. 02-5340082"
       
    .Sections("Section2").Controls("lblTitle").Caption = txtTitle
    .Sections("Section2").Controls("L1").Caption = SELECT1TITLE
    .Sections("Section2").Controls("L2").Caption = SELECT2TITLE
    .Sections("Section2").Controls("L3").Caption = SELECT3TITLE
    .Sections("Section2").Controls("L4").Caption = SELECT4TITLE
    .Sections("Section2").Controls("L5").Caption = SELECT5TITLE
    .Sections("Section1").Controls("Text1").DataField = SELECT1
    .Sections("Section1").Controls("Text2").DataField = SELECT2
    .Sections("Section1").Controls("Text3").DataField = SELECT3
    .Sections("Section1").Controls("Text4").DataField = SELECT4
    .Sections("Section1").Controls("Text5").DataField = SELECT5
    '.Sections("Section5").Controls("lblRecCount").Caption = CreateRecords.RecordCount & " items listed"
    DoEvents
    .Show
  End With
End Sub
Public Function CreateRecords1(ByVal tblName As String) As ADODB.Recordset
  Dim rsTemp As New ADODB.Recordset, i As Integer
  Dim Sortby, ByDirection As String
  ByDirection = IIf(chkSort(0).Value = 1, "Desc", "Asc")
  If cboSort(0).Text <> "(None)" And Trim(cboSort(0).Text) <> "" Then Sortby = " order by " & cboSort(0).Text & " " & ByDirection
  sqlString = "SELECT " & _
              "" & SELECTED & "" & _
               " From " & tblName & "  where employeename = '" & CboName.Text & "'" & whereCon & Sortby
 Set CreateRecords1 = ReturnRst(sqlString, Cnnstr, mread)
End Function

Public Function CreateRecords() As ADODB.Recordset
  Dim rsTemp As New ADODB.Recordset, i As Integer
  Dim Sortby, ByDirection As String
  ByDirection = IIf(chkSort(0).Value = 1, "Desc", "Asc")
  
  If cboSort(0).Text <> "(None)" And Trim(cboSort(0).Text) <> "" Then Sortby = " order by " & cboSort(0).Text & " " & ByDirection
   
  sqlString = "SELECT " & _
              "" & SELECTED & "" & _
               " From qryCustomized " & whereCon & Sortby
 Set CreateRecords = ReturnRst(sqlString, Cnnstr, mread)
End Function



Private Sub Form_Unload(Cancel As Integer)
    Set rsItems = Nothing
End Sub

Private Sub lstVw_Click()
    txtTitle = UCase(lstVw.Text)
'    If lstVw.ListIndex = 12 Then
'    Frame2.Visible = False
'    Frame1.Visible = False
'    Frame3.Visible = True
'    Call ListBenefitName
'    CboName.SetFocus
'      ElseIf lstVw.ListIndex = 13 Then
'    Frame2.Visible = False
'    Frame1.Visible = False
'    Frame3.Visible = True
'    Call ListBenefitName
'    CboName.SetFocus
'    Else
'    Frame2.Visible = True
'    Frame1.Visible = True
'    Frame3.Visible = False
'    End If

Dim X As Integer
 For X = 0 To 3
    Check1(X).Enabled = True
Next X
  
If cboFilterby(0).ListCount <= 1 Then
 For i = 0 To cboFilterby.Count - 1
           cboFilterby(0).Clear
           cboFilterby(0).AddItem "Client name"
           cboFilterby(0).AddItem "Start of contract"
           cboFilterby(0).AddItem "End of contract"
           cboFilterby(0).AddItem "Civil status"
 Next
 End If
 
  If lstVw.ListIndex = 7 Then
   Call ChangeChkStatus(False)
   chkStatus(5).Value = 1
   
   ElseIf lstVw.ListIndex = 15 Then
   Call ChangeChkStatus(True)
   chkStatus(1).Value = 1
If cboFilterby(0).ListCount > 1 Then
   '  Check1(0).Enabled = False 'remove
   Check1(1).Enabled = False
   Check1(2).Enabled = False
   Check1(3).Enabled = False
    cboFilterby(0).RemoveItem (2)
   cboFilterby(0).RemoveItem (2)
   cboFilterby(0).RemoveItem (1)
 Check2.Enabled = False
 
 End If
 
   
   
   Exit Sub
   Else
    Call ChangeChkStatus(True)
    chkStatus(1).Value = 1
   End If
   
  'chkStatus(5).Enabled = True
  'chkStatus(7).Enabled = True
  Check2.Enabled = True
  
    For X = 0 To 3
   If Me.Check1(X).Value = 1 Then
   cboFilterby(X).Enabled = True
   Else
    
    cboFilterby(X).Enabled = False
   End If
  
   Next X

End Sub
Private Sub ListBenefitName()
Dim rsemp As New ADODB.Recordset
Set rsemp = cnDCPI.Execute("Select distinct employeename From hrrptbenefits")
CboName.Clear
Do While Not rsemp.EOF
        CboName.AddItem rsemp(0)
        rsemp.MoveNext
    Loop
    rsemp.Close
    Set rsemp = Nothing
End Sub

Private Function GetDataName(ByVal tblWhere As String, ByVal returnFld As String) As String
Dim rsemp As New ADODB.Recordset
Set rsemp = cnDCPI.Execute("Select distinct " & returnFld & " From " & tblWhere)
CboName.Clear
If rsemp.EOF = False Then
        GetDataName = rsemp(0)
        rsemp.MoveNext
  End If
    rsemp.Close
    Set rsemp = Nothing
End Function

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Index = 1 Or Item.Index = 2 Then
        Lv.ListItems(Item.Index).Checked = True
        MsgBox "Field cannot be remove from the report", vbInformation
        Exit Sub
    Else
    Dim X As Integer
    If isCheckExceed(Item.Index) = True Then
    For X = 0 To 3
        
         Me.Check1(X).Enabled = True
        cboFilterby(X).Enabled = False
      
    Next X
  '   chkStatus(5).Enabled = True
  'chkStatus(7).Enabled = True
    Else
    
   '  chkStatus(5).Enabled = False
 ' chkStatus(7).Enabled = False
     For X = 0 To 3
    cboFilterby(X).Enabled = False
    Me.Check1(X).Enabled = False
    Me.Check1(X).Value = 0
          cboFilterby(X).BackColor = &H8000000A
          cboFilterby(X).ListIndex = -1
          lblCon(X).Visible = False
   'If Me.Check1(X).Value = 1 Then cboFilterby(X).Enabled = True
   'If Me.Check1(X).Value = 0 Then cboFilterby(X).Enabled = False
    Next X
    End If
    End If
End Sub

Private Sub mskDate1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then mskDate1(1).SetFocus
End Sub

Private Sub mskDate1_Validate(Index As Integer, Cancel As Boolean)
    If mskDate1(Index) <> "__/__/____" And IsDate(mskDate1(Index)) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskDate1(Index).SetFocus
End Sub

Private Sub mskDate2_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then mskDate2(1).SetFocus
End Sub

Private Sub mskDate2_Validate(Index As Integer, Cancel As Boolean)
    If mskDate2(Index) <> "__/__/____" And IsDate(mskDate2(Index)) = False Then Cancel = True: MsgBox "Invalid date", vbInformation: mskDate2(Index).SetFocus
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
 If PreviousTab = 2 Then
    Frame2.Visible = False
    Frame1.Visible = False

    Dim X As Integer
    If isCheckExceed(1) = True Then
    For X = 0 To 3
         Me.Check1(X).Enabled = False
        cboFilterby(X).Enabled = False
    Next X
  '   chkStatus(5).Enabled = False
 ' chkStatus(7).Enabled = False
    Else
    Frame2.Visible = True
    Frame1.Visible = True
    Frame3.Visible = False
    lstVw.ListIndex = 0
    End If
    End If
    
End Sub

Private Sub txtTitle_GotFocus()
    Set oCommonFunction.ValidAlpha = txtTitle
     Call txtHLigths(txtTitle)
End Sub
Public Sub CheckField()
    SELECT1 = ""
    SELECT2 = ""
    SELECT3 = ""
    SELECT4 = ""
    SELECT5 = ""
    SELECTED = ""
    i = 1
    Do While Not i = 38
     If Lv.ListItems.Item(i).Checked = True Then
     If SELECTED <> "" Then SELECTED = SELECTED & ","
        SELECTED = SELECTED & Lv.ListItems.Item(i).SubItems(1)
       If SELECT1 = "" Then
        SELECT1 = Lv.ListItems.Item(i).SubItems(1)
        SELECT1TITLE = Lv.ListItems.Item(i).Text
       ElseIf SELECT2 = "" Then
        SELECT2 = Lv.ListItems.Item(i).SubItems(1)
        SELECT2TITLE = Lv.ListItems.Item(i).Text
       ElseIf SELECT3 = "" Then
        SELECT3 = Lv.ListItems.Item(i).SubItems(1)
        SELECT3TITLE = Lv.ListItems.Item(i).Text
       ElseIf SELECT4 = "" Then
        SELECT4 = Lv.ListItems.Item(i).SubItems(1)
        SELECT4TITLE = Lv.ListItems.Item(i).Text
       ElseIf SELECT5 = "" Then
        SELECT5 = Lv.ListItems.Item(i).SubItems(1)
        SELECT5TITLE = Lv.ListItems.Item(i).Text
       End If
    End If
    i = i + 1
    Loop
End Sub
Public Function isCheckExceed(Optional ByVal ndx As Integer) As Boolean
i = 1
Dim iss As Integer
  'If Lv.ListItems.Item(i).Checked = True Then cboFilterby(0).AddItem Lv.ListItems.Item(ndx).Text
  'If Lv.ListItems.Item(i).Checked = false Then cboFilterby(0) AddItem Lv.ListItems.Item(ndx).Text
  
    Do While Not i = 38
     If Lv.ListItems.Item(i).Checked = True Then
        iss = iss + 1
     End If
     i = i + 1
    Loop
   If iss = 5 Then
   isCheckExceed = True
   Else
   isCheckExceed = False
   End If
End Function

Private Sub CreateMasterList()
Dim str As String
Dim rsTemp As ADODB.Recordset
Dim xlApp1 As Excel.Application
Dim xlBook1 As Excel.Workbook
Dim xlSheet1 As Excel.Worksheet
Dim ctr As Long
Dim i As Long

Dim Sortby, ByDirection As String
  ByDirection = IIf(chkSort(0).Value = 1, "Desc", "Asc")
  
  If cboSort(0).Text <> "(None)" And Trim(cboSort(0).Text) <> "" Then Sortby = " order by " & cboSort(0).Text & " " & ByDirection
   
   If Sortby = "" Then Sortby = " order by Empname"
SELECTED = " Empname,Controlno,Brname,Deptname, Empposition,Telnum1,Telnum2,startcontract,endcontract,sss,philhealth,taxid,birthdate,civilstatus,Dailyrate,ecolaamount,reasonres,employmentstatusdesc "
  sqlString = "SELECT " & _
              "" & SELECTED & "" & _
               " From qryCustomized " & whereCon & Sortby
Sortby = ""
               
Set rsTemp = ReturnRst(sqlString, cnDCPI)
If rsTemp.EOF = False Then
    Set xlApp1 = New Excel.Application
    Set xlBook1 = xlApp1.Workbooks.Add()
    xlApp1.Visible = True
    CreateExcel rsTemp, xlBook1, xlApp1
    xlApp1.DisplayAlerts = False
    
 Else
    MsgBox "No Records!!!", vbCritical, "HR Leave OT System"
End If
End Sub

Function CreateExcel(rs As ADODB.Recordset, book1 As Excel.Workbook, ByVal xlApp1 As Excel.Application)
    Dim xlSheet As Excel.Worksheet
    Dim ctr As Long
    Dim i As Long
    Set xlSheet = book1.Worksheets.Add
    ctr = 7
    i = 1
    xlSheet.Cells(1, 1).Value = "MASTERLIST OF STAFFERS " & IIf(CboCompany.ListIndex >= 0, " ASSIGNED TO " & CboCompany.Text, "")
    xlSheet.Cells(2, 1).Value = "As of:"
    xlSheet.Cells(2, 3).Value = Format(Date, "mmmm dd, yyyy")
    xlSheet.Cells(3, 1).Value = "Agency:"
    xlSheet.Cells(3, 3).Value = CompanyName
    xlSheet.Cells(4, 1).Value = "Prepared by:"
    xlSheet.Cells(4, 3).Value = UCase(GETuSERNAME(operatorID))
    xlSheet.Cells(6, 1).Value = "No."
    xlSheet.Cells(6, 2).Value = "NAME"
    xlSheet.Cells(6, 3).Value = "ID No."
    xlSheet.Cells(6, 4).Value = "Branch"
    xlSheet.Cells(6, 5).Value = "Department"
    xlSheet.Cells(6, 6).Value = "Position"
    xlSheet.Cells(6, 7).Value = "Tel. no."
    xlSheet.Cells(6, 8).Value = "Mobile no."
    xlSheet.Cells(6, 9).Value = "FROM"
    'xlSheet.Cells.Range("F6:G6").Merge
    xlSheet.Cells(6, 10).Value = "TO"
    
    xlSheet.Cells(6, 11).Value = "SSS"
    
    xlSheet.Cells(6, 12).Value = "PH#"
    xlSheet.Cells(6, 13).Value = "TIN#"
    xlSheet.Cells(6, 14).Value = "DOB"
    xlSheet.Cells(6, 15).Value = "CS"
    xlSheet.Cells(6, 16).Value = "BASIC PAY"
    xlSheet.Cells(6, 17).Value = "ECOLA"
    xlSheet.Cells(6, 18).Value = "STATUS"
    xlSheet.Cells(6, 19).Value = "REMARKS"
    
    xlSheet.Cells.EntireColumn.AutoFit
    xlSheet.Cells.EntireRow.AutoFit
    xlSheet.Cells.Range("A1:S6").Font.Bold = True
    xlSheet.Cells.Range("A1:a5").HorizontalAlignment = xlLeft
    xlSheet.Cells.Range("c1:c5").HorizontalAlignment = xlLeft
    xlSheet.Cells.Range("A6:S6").Interior.ColorIndex = 15
       xlSheet.Columns("A:A").ColumnWidth = 5.71
    Do While Not rs.EOF
        DoEvents
        xlSheet.Cells(ctr, 1).Value = i
        xlSheet.Cells(ctr, 2).Value = rs!empname
        xlSheet.Cells(ctr, 3).Value = rs!ControlNo
        xlSheet.Cells(ctr, 4).Value = rs!brname
        xlSheet.Cells(ctr, 5).Value = rs!deptname
        xlSheet.Cells(ctr, 6).Value = rs!Empposition
        xlSheet.Cells(ctr, 7).Value = rs!Telnum1
        xlSheet.Cells(ctr, 8).Value = rs!Telnum2
        
         xlSheet.Cells(ctr, 9).Value = rs!StartContract
        xlSheet.Cells(ctr, 9).NumberFormat = "mm/dd/yyyy"
        xlSheet.Cells(ctr, 10).Value = rs!endcontract
        xlSheet.Cells(ctr, 10).NumberFormat = "mm/dd/yyyy"
    
        xlSheet.Cells(ctr, 11).Value = IIf(rs!sss = "__-_______-_", "", rs!sss)
        xlSheet.Cells(ctr, 12).Value = IIf(rs!philhealth = "__-_________-_", "", rs!philhealth)
        xlSheet.Cells(ctr, 13).Value = IIf(rs!taxid = "___-___-___", "", rs!taxid)
        xlSheet.Cells(ctr, 14).Value = rs!BIRTHDATE
        xlSheet.Cells(ctr, 14).NumberFormat = "mm/dd/yyyy"
        xlSheet.Cells(ctr, 15).Value = rs!civilstatus
        
        If rs!DailyRate <> "" Or IsNull(rs!DailyRate) = False Then
        xlSheet.Cells(ctr, 16).Value = "Php " & Format(rs!DailyRate, "#,##0.00")
        End If
        
        If rs!EcolaAmount <> "" Or IsNull(rs!EcolaAmount) = False Then
        xlSheet.Cells(ctr, 17).Value = " Php " & Format(rs!EcolaAmount, "#,##0.00")
        End If
        xlSheet.Cells(ctr, 18).Value = rs!EmploymentStatusdesc
        xlSheet.Cells(ctr, 19).Value = rs!reasonres
        
        i = i + 1
        ctr = ctr + 1
        rs.MoveNext
    Loop
        xlSheet.Cells.EntireColumn.AutoFit
        xlSheet.Cells.EntireRow.AutoFit
       xlSheet.Columns("A:A").ColumnWidth = 5.71
    
       xlSheet.Columns("C:C").ColumnWidth = 10
    
   
    
    'center
    
        xlSheet.Range("A6:S6").Select
    With xlApp1.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    
    
    'center id
    
      xlSheet.Range("A7:A" & ctr).Select
    With xlApp1.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
     
       'line
    'Range("A6:Q14").Select
    xlSheet.Range("A6:S" & ctr).Select
    xlApp1.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    xlApp1.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With xlApp1.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlApp1.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    
    
    
    
    
    '    xlSheet.Rows("1:1").Select
    '    With Selection.Font
    '        .Name = "Arial"
    '        .Size = 16
    '    End With
    '    xlSheet.Range("C2:C4").Select
    '    Selection.Font.Bold = False
    xlSheet.Name = "Masterlist"
End Function
