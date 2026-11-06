VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSelectCutoff 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Payroll Cut-Off"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "frmSelectCutoff.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEditCutoff 
      Caption         =   "&Edit Cut-Off"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   17
      Top             =   3240
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3150
      TabIndex        =   8
      Top             =   3240
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1890
      TabIndex        =   7
      Top             =   3240
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3165
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CheckBox chkForYr 
         Caption         =   "For the Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   23
         Top             =   240
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.TextBox txtCutoffYr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5160
         TabIndex        =   22
         Text            =   "2019"
         ToolTipText     =   "Cut-off Year"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtSeriesNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4560
         TabIndex        =   20
         Text            =   "0"
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtCutoffBase 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3420
         TabIndex        =   4
         Text            =   "Semi-Monthly (15)"
         Top             =   1710
         Width           =   1905
      End
      Begin VB.Frame Frame2 
         Caption         =   "Payroll Period"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   630
         TabIndex        =   13
         Top             =   2070
         Width           =   4785
         Begin MSMask.MaskEdBox txtCutOffDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   990
            TabIndex        =   5
            Top             =   270
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "mm/dd/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtCutOffDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "MM/dd/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   3060
            TabIndex        =   6
            Top             =   270
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   0
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "mm/dd/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblEntries 
            BackStyle       =   0  'Transparent
            Caption         =   "To:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   2610
            TabIndex        =   15
            Top             =   330
            Width           =   465
         End
         Begin VB.Label lblEntries 
            BackStyle       =   0  'Transparent
            Caption         =   "From:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   14
            Top             =   330
            Width           =   825
         End
      End
      Begin VB.CheckBox chkNewCutoff 
         Caption         =   "New Cut-Off"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   540
         TabIndex        =   3
         Top             =   1710
         Width           =   1365
      End
      Begin VB.TextBox txtDateToday 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1290
         TabIndex        =   9
         Text            =   "txtDateToday"
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbCutOff 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSelectCutoff.frx":08CA
         Left            =   1800
         List            =   "frmSelectCutoff.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1170
         Width           =   3375
      End
      Begin VB.ComboBox cmbClient 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSelectCutoff.frx":08CE
         Left            =   1800
         List            =   "frmSelectCutoff.frx":08DB
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   4440
         TabIndex        =   21
         Top             =   270
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   5220
         Picture         =   "frmSelectCutoff.frx":08F8
         Top             =   750
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   5220
         Picture         =   "frmSelectCutoff.frx":0E82
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Cut-Off Basis:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   16
         Top             =   1740
         Width           =   1365
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Cut-Off:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   540
         TabIndex        =   12
         Top             =   1260
         Width           =   1365
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Client:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   540
         TabIndex        =   11
         Top             =   810
         Width           =   1095
      End
      Begin VB.Label lblEntries 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Today:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   10
         Top             =   270
         Width           =   1095
      End
   End
   Begin MSMask.MaskEdBox txtCutOffDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   0
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   0
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "mm/dd/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtCutOffDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   270
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   0
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "mm/dd/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "frmSelectCutoff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCutOffTable As ADODB.Recordset
Dim OnProcessFlag As Boolean
Dim NoDaysperCutoff As Integer
Dim StrQuery As String
Dim EditCutoffSw As Boolean
Dim PerYearFlag As Boolean

Private Sub chkNewCutoff_Click()
Dim mEndDate As Date
    If chkNewCutoff = 1 Then ' get new cut-off w/ accordance to the last cut-off
        If rsCutOffTable.RecordCount = 0 Then ' enable the start and end cut-off date for manual entry
            txtCutOffDate(0).Text = Format(Now - NoDaysperCutoff, "mm/dd/yyyy")
            txtCutOffDate(1).Text = Format(Now, "mm/dd/yyyy")
        Else
            rsCutOffTable.MoveLast  ' go to last record and get the last cut-off
            mEndDate = Format(rsCutOffTable!EndCutoffDate + 1, "mm/dd/yyyy")
            txtCutOffDate(0).Text = Format(mEndDate, "mm/dd/yyyy")
            txtCutOffDate(1).Text = Format((mEndDate + NoDaysperCutoff - 1), "mm/dd/yyyy")
        End If
        txtCutOffDate(0).Text = "__/__/____"
        txtCutOffDate(1).Text = "__/__/____"
        txtCutOffDate(0).Enabled = True
        txtCutOffDate(1).Enabled = True
        
        txtSeriesNo.Text = CompanyShort & Format(Now, "yy") & "-" & Right("0000" & ClientSeriesNo, 4)
        txtSeriesNo.Refresh
        txtCutOffDate(0).SetFocus
    Else
        txtCutOffDate(0).Text = txtCutOffDate(2).Text
        txtCutOffDate(1).Text = txtCutOffDate(3).Text
        txtCutOffDate(0).Enabled = False
        txtCutOffDate(1).Enabled = False
    End If
End Sub

Private Sub cmbClient_Click()
    ' Get Cutoff table based from selected client.
    mClientCode = cmbClient.ItemData(cmbClient.ListIndex)
    rsClientTable.MoveFirst
    rsClientTable.Find "clientcode = " & mClientCode, 0, adSearchForward ' goto selected client record
    If Not rsClientTable.EOF Then
        NoDaysperCutoff = IIf(IsNull(rsClientTable!ClientCutoffDates), 0, Val(rsClientTable!ClientCutoffDates))
        txtCutoffBase.Text = rsClientTable!ClientCutoff & " (" & Trim(NoDaysperCutoff) & ")"
        ClientSeriesNo = IIf(IsNull(rsClientTable!ClientSeriesNo), 0, Val(rsClientTable!ClientSeriesNo))
        ClientSeriesNo = ClientSeriesNo + 1 ' series number + 1
        If rsClientTable!clientflag1 = True Then
            Image3.Visible = True ' show the on process flag
        Else
            Image3.Visible = False
        End If
        Image2.Visible = False
       
        If PerYearFlag = False Then
            ' Get only those cutoff period with viewflag = true
            Set rsCutOffTable = ReturnRst("SELECT * FROM cutofftable where clientcode = " & mClientCode & " and viewflag = 1 ORDER BY RecNo DESC", cnnstr, mstatic)
        Else
            ' Get only those cutoff period for selected year.
            Set rsCutOffTable = ReturnRst("SELECT * FROM cutofftable WHERE clientcode = " & mClientCode & " AND YEAR(StartCutoffDate) = " & Val(txtCutoffYr.Text) & " ORDER BY RecNo DESC", cnnstr, mstatic)
        End If
        ' Populate cutoff list.
        chkNewCutoff.Visible = True
        chkNewCutoff.Enabled = True
        With rsCutOffTable
            cmbCutoff.Clear
            While Not (.EOF Or .BOF)
                'cmbCutoff.AddItem Format(.Fields(5), "mm/dd/yyyy") & "-" & Format(.Fields(6), "mm/dd/yyyy") & "  " & IIf(!billedflag = 1, "Close", "Open")
                cmbCutoff.AddItem Format(!StartCutoffDate, "mm/dd/yyyy") & "-" & Format(!EndCutoffDate, "mm/dd/yyyy") & "  " & IIf(!BilledFlag = True, "Closed", "Open")
                cmbCutoff.ItemData(cmbCutoff.NewIndex) = .Fields(0)
                If !BilledFlag = False Then
                    chkNewCutoff.Visible = False
                End If
                .MoveNext
            Wend
            cmbCutoff.Refresh
        End With
        txtSeriesNo.Text = CompanyShort & Format(Now, "yy") & "-" & Right("0000" & ClientSeriesNo, 4)
        txtSeriesNo.Refresh
    Else
        MsgBox "Invalid client code selected!", vbInformation, PayrollAlert
    End If
End Sub

Private Sub cmbClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub cmbClient_LostFocus()
    ' Get Cutoff table based from selected client.
    mClientCode = cmbClient.ItemData(cmbClient.ListIndex)
    mCompanyName = Trim(cmbClient.Text)
    
    If PerYearFlag = False Then
        'Get Current payroll cut-off.
        Set rsCutOffTable = ReturnRst("SELECT * FROM cutofftable where clientcode = " & mClientCode & " and viewflag = 1 ORDER BY RecNo", cnnstr, mstatic)
    Else
        ' Get only those cutoff period for selected year.
        Set rsCutOffTable = ReturnRst("SELECT * FROM cutofftable WHERE clientcode = " & mClientCode & " AND YEAR(StartCutoffDate) = " & Val(txtCutoffYr.Text) & " ORDER BY RecNo", cnnstr, mstatic)
    End If
    
    If rsCutOffTable.RecordCount = 0 Then ' no payroll cut-off selected, enable new cutoff
        chkNewCutoff.Enabled = True
    Else
        chkNewCutoff.Enabled = False
        chkNewCutoff.VALUE = 0
    End If
    ' Populate cutoff list.
    chkNewCutoff.Visible = True
    chkNewCutoff.Enabled = True
    With rsCutOffTable
        cmbCutoff.Clear
        While Not (.EOF Or .BOF)
            cmbCutoff.AddItem Format(!StartCutoffDate, "mm/dd/yyyy") & "-" & Format(!EndCutoffDate, "mm/dd/yyyy") & "  " & IIf(!BilledFlag = True, "Close", "Open")
            cmbCutoff.ItemData(cmbCutoff.NewIndex) = .Fields(0)
            If !BilledFlag = False Then
                chkNewCutoff.Visible = False
            End If
            .MoveNext
        Wend
        cmbCutoff.Refresh
    End With
End Sub

Private Sub cmbCutoff_Click()
Dim CutOffIndex As Integer
Dim mEndDate As Date

    'On Error GoTo ExitSub
    If cmbCutoff.ListIndex = -1 Then
        Exit Sub
    Else
        CutOffIndex = IIf(cmbCutoff.ListIndex = -1, 0, cmbCutoff.ItemData(cmbCutoff.ListIndex))
    End If
    rsCutOffTable.MoveFirst
    rsCutOffTable.Find "recno = " & CutOffIndex, 0, adSearchForward ' goto selected client record
    txtCutOffDate(0).Text = Format(rsCutOffTable!StartCutoffDate, "mm/dd/yyyy")
    txtCutOffDate(1).Text = Format(rsCutOffTable!EndCutoffDate, "mm/dd/yyyy")
    txtCutOffDate(2).Text = txtCutOffDate(0).Text
    txtCutOffDate(3).Text = txtCutOffDate(1).Text
    ' Get the old client series no.
    txtSeriesNo.Text = IIf(IsNull(rsCutOffTable!ClientSeriesNo), txtSeriesNo.Text, rsCutOffTable!ClientSeriesNo)
    txtSeriesNo.Refresh
    
    If rsCutOffTable!BilledFlag = False Then
        Image2.Visible = True
        chkNewCutoff.Enabled = False
        chkNewCutoff.VALUE = 0
    Else
        Image2.Visible = False
        chkNewCutoff.Enabled = True
        'chkNewCutoff.Visible = True
    End If
    If chkNewCutoff = 1 Then ' get new cut-off w/ accordance to the last cut-off
        rsCutOffTable.MoveLast  ' go to last record and get the last cut-off
        mEndDate = Format(rsCutOffTable!EndCutoffDate + 1, "mm/dd/yyyy")
        txtCutOffDate(0).Text = Format(mEndDate, "mm/dd/yyyy")
        txtCutOffDate(1).Text = Format((mEndDate + NoDaysperCutoff), "mm/dd/yyyy")
        txtCutOffDate(0).Enabled = True
        txtCutOffDate(1).Enabled = True
    Else
        txtCutOffDate(0).Enabled = False
        txtCutOffDate(1).Enabled = False
    End If
    Exit Sub

ExitSub:

    'MsgBox "Error: In Cut-off Selection!", vbInformation, PayrollAlert
    Err.Clear

End Sub

Private Sub cmbCutoff_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub cmdCancel_Click()

If EditCutoffSw = False Then
    Unload Me
Else    ' Cancel Edited Cut-off period changes
    'Enable buttons and change caption to default
    cmbClient.Enabled = True
    cmbCutoff.Enabled = True
    chkNewCutoff.Enabled = True
    Frame2.Caption = "Payroll Period"
    cmdOK.Caption = "&OK" '
    txtCutOffDate(0).Enabled = False
    txtCutOffDate(1).Enabled = False
    cmdEditCutoff.Enabled = True
    EditCutoffSw = False
End If

End Sub

Private Sub chkForYr_Click()
If chkForYr.VALUE = 1 Then
    PerYearFlag = True
Else
    PerYearFlag = False
End If

If PerYearFlag = False Then
    ' Get only those cutoff period with viewflag = true
    Set rsCutOffTable = ReturnRst("SELECT * FROM cutofftable where clientcode = " & mClientCode & " and viewflag = 1 ORDER BY RecNo", cnnstr, mstatic)
Else
    ' Get only those cutoff period for selected year.
    Set rsCutOffTable = ReturnRst("SELECT * FROM cutofftable WHERE clientcode = " & mClientCode & " AND YEAR(StartCutoffDate) = " & Val(txtCutoffYr.Text) & " ORDER BY RecNo", cnnstr, mstatic)
End If
' Populate cutoff list.
chkNewCutoff.Visible = True
chkNewCutoff.Enabled = True
With rsCutOffTable
    cmbCutoff.Clear
    While Not (.EOF Or .BOF)
        cmbCutoff.AddItem Format(!StartCutoffDate, "mm/dd/yyyy") & "-" & Format(!EndCutoffDate, "mm/dd/yyyy") & "  " & IIf(!BilledFlag = True, "Closed", "Open")
        cmbCutoff.ItemData(cmbCutoff.NewIndex) = .Fields(0)
        If !BilledFlag = False Then
            chkNewCutoff.Visible = False
        End If
        .MoveNext
    Wend
    cmbCutoff.Refresh
End With
txtSeriesNo.Text = CompanyShort & Format(Now, "yy") & "-" & Right("0000" & ClientSeriesNo, 4)
txtSeriesNo.Refresh

End Sub

Private Sub cmdEditCutoff_Click()
'Save the selected payroll cut-off for edit
txtCutOffDate(2).Text = txtCutOffDate(0).Text
txtCutOffDate(3).Text = txtCutOffDate(1).Text
cmbClient.Enabled = False
cmbCutoff.Enabled = False
chkNewCutoff.Enabled = False
Frame2.Caption = "Edit Cut-Off"
cmdOK.Caption = "Save"
txtCutOffDate(0).Enabled = True
txtCutOffDate(1).Enabled = True
txtCutOffDate(0).SetFocus
cmdEditCutoff.Enabled = False
EditCutoffSw = True
End Sub

Private Sub cmdOK_Click()
Dim mAns

ClientSeriesNo = Val(mID(txtSeriesNo.Text, InStr(txtSeriesNo.Text, "-") + 1))
txtSeriesNo.Text = CompanyShort & Format(Now, "yy") & "-" & Right("0000" & ClientSeriesNo, 4)
txtSeriesNo.Refresh

If EditCutoffSw = False Then
    If chkNewCutoff.VALUE = 1 Then
        'added by Bernard 01042006
        If txtCutOffDate(0).Text = "__/__/____" Or txtCutOffDate(1).Text = "__/__/____" Then
            MsgBox "Please complete the cut-off date.", vbInformation
            Exit Sub
        ElseIf IsDate(txtCutOffDate(0).Text) = False Or IsDate(txtCutOffDate(1).Text) = False Then
        MsgBox "Invalid cut off date", vbInformation
        Exit Sub
        End If
       
        mAns = MsgBox("Create New Payroll Cut-off: " & txtCutOffDate(0).Text & " - " & txtCutOffDate(1).Text & "?", vbQuestion + vbYesNo, PayrollAlert)
        If mAns = vbYes Then
            Call ProtectDatabaseRecords(CompanyCode, mClientCode, True)
            frmPayrollEntry.Show
            frmSelectCutoff.Hide
        End If
    
    Else
        'retrieve previous payroll cut-off
        mAns = MsgBox("Retrieve previous Payroll Cut-off: " & txtCutOffDate(0).Text & " - " & txtCutOffDate(1).Text & "?", vbQuestion + vbYesNo, PayrollAlert)
        If mAns = vbYes And cmbCutoff.Text <> "" Then
            Call ProtectDatabaseRecords(CompanyCode, mClientCode, True)
            frmPayrollEntry.Show
            frmSelectCutoff.Hide
        Else
            If mAns = vbYes And cmbCutoff.Text = "" Then
                MsgBox "Select Payroll Cut-off from the list!", vbInformation, PayrollAlert
            End If
        End If
    End If
Else    ' Save Edited Cut-off period and reflect it to all tables
    'Check if there is a possible gap that will occur.
    EditCutoffSw = False
End If

End Sub

Private Sub Form_Activate()
'chkNewCutoff.Value = 0
Call Form_Load
End Sub

Private Sub Form_Load()
Dim StrQuery As String
    
    ' Center screen
 '   Move (mdiMain.Width - Width) \ 2, (mdiMain.Height - Height) \ 2
    Me.Top = (mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
    txtDateToday = Format(Now, "mm/dd/yyyy")
    
    txtCutoffYr.Text = Format(Now, "yyyy")
    PerYearFlag = True

    ' Get Company Info
    Call CompanyInfo
    ' Get client list
    Set rsClientTable = New ADODB.Recordset
'    strQuery = "SELECT clientcode, clientname, clientcutoff, clientcutoffdates, clientflag1 FROM clienttable WHERE clientcode <> 0"
    
'    If AccountType = "FMS" Then ' Updated by Bernard 6/30/2011.
'        StrQuery = "SELECT * FROM clienttable WHERE CLIENTCODE2 = 'FMS' AND clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' ORDER BY clientname"
'    Else
'        StrQuery = "SELECT * FROM clienttable WHERE clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' ORDER BY clientname"
'    End If
    StrQuery = GetClientRecord(0, False, " ORDER BY ClientName")
    'StrQuery = "SELECT * FROM clienttable WHERE clientcode <> 0 AND RTRIM(CLIENTNAME)<>'' ORDER BY clientname"
    
    Set rsClientTable = ReturnRst(StrQuery, cnnstr, mstatic)
    Call oCommonFunction.FillCombo(cmbClient, rsClientTable)
    If rsClientTable.RecordCount <> 0 Then cmbClient.ListIndex = 0
    'If rsClientTable.State = adStateOpen Then rsClientTable.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsCompanyTable = Nothing
Set rsClientTable = Nothing
Unload Me
End Sub

Private Sub txtCutoffDate_GotFocus(Index As Integer)
Sendkeys "{HOME}+{END}"
End Sub

Private Sub txtCutoffDate_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii = 13 Then
        Sendkeys "{tab}"
    End If
End Sub

Public Sub CompanyInfo()
    
    StrQuery = "SELECT * FROM companytable WHERE agencycode = 1"
    Set rsCompanyTable = ReturnRst(StrQuery, CnnstrHR, mstatic)

End Sub

Private Sub txtCutOffDate_Validate(Index As Integer, Cancel As Boolean)
'updated by Bernard 01042007
     If IsDate(txtCutOffDate(Index).Text) = False And txtCutOffDate(Index).Text <> "__/__/____" Then
        MsgBox "Invalid Date Entered!"
        txtCutOffDate(Index).SetFocus
        Exit Sub
     End If
    If (IsDate(txtCutOffDate(0).Text) = False Or IsDate(txtCutOffDate(1).Text) = False) And txtCutOffDate(0).Text <> "__/__/____" And txtCutOffDate(1).Text <> "__/__/____" Then
        MsgBox "Invalid Date Entered!"
        Exit Sub
    Else
        If txtCutOffDate(0).Text <> "__/__/____" And txtCutOffDate(1).Text <> "__/__/____" Then
            If CDate(txtCutOffDate(0).Text) > CDate(txtCutOffDate(1).Text) Then
            MsgBox "End Cut-off should be greater than start Cut-off!", vbInformation, PayrollAlert
            Cancel = True
            txtCutOffDate(0).SetFocus
            Exit Sub
        End If
        End If
        
    End If
End Sub

