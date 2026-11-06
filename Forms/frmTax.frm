VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTaxTable 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Table Maintenance"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   Icon            =   "frmTax.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   8115
   Begin TabDlg.SSTab SSTab1 
      Height          =   1035
      Index           =   2
      Left            =   1200
      TabIndex        =   12
      Top             =   4800
      Visible         =   0   'False
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   1826
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
      TabCaption(0)   =   "      "
      TabPicture(0)   =   "frmTax.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblExit"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   270
         TabIndex        =   16
         Top             =   540
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   495
         TabIndex        =   15
         Top             =   540
         Width           =   1320
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2205
         TabIndex        =   14
         Top             =   540
         Width           =   1320
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3915
         TabIndex        =   13
         Top             =   540
         Width           =   1320
      End
      Begin VB.Label Label2 
         Caption         =   "Rate from"
         Height          =   240
         Left            =   270
         TabIndex        =   21
         Top             =   270
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Rate to"
         Height          =   240
         Left            =   765
         TabIndex        =   20
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label Label4 
         Caption         =   "Base Tax"
         Height          =   240
         Left            =   2565
         TabIndex        =   19
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label Label5 
         Caption         =   "Per over"
         Height          =   240
         Left            =   4185
         TabIndex        =   18
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label lblExit 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5310
         MouseIcon       =   "frmTax.frx":08E6
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   90
         Width           =   195
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3120
      Left            =   720
      TabIndex        =   11
      Top             =   1560
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   5503
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
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
      NumItems        =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   1080
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   12
      TabsPerRow      =   12
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
      TabCaption(0)   =   "HF"
      TabPicture(0)   =   "frmTax.frx":0BF0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "HF1"
      TabPicture(1)   =   "frmTax.frx":0C0C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "HF2"
      TabPicture(2)   =   "frmTax.frx":0C28
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "HF3"
      TabPicture(3)   =   "frmTax.frx":0C44
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "HF4"
      TabPicture(4)   =   "frmTax.frx":0C60
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "ME"
      TabPicture(5)   =   "frmTax.frx":0C7C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "ME1"
      TabPicture(6)   =   "frmTax.frx":0C98
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "ME2"
      TabPicture(7)   =   "frmTax.frx":0CB4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "ME3"
      TabPicture(8)   =   "frmTax.frx":0CD0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabCaption(9)   =   "ME4"
      TabPicture(9)   =   "frmTax.frx":0CEC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      TabCaption(10)  =   "S"
      TabPicture(10)  =   "frmTax.frx":0D08
      Tab(10).ControlEnabled=   0   'False
      Tab(10).ControlCount=   0
      TabCaption(11)  =   "Z"
      TabPicture(11)  =   "frmTax.frx":0D24
      Tab(11).ControlEnabled=   0   'False
      Tab(11).ControlCount=   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7425
      Top             =   225
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
            Picture         =   "frmTax.frx":0D40
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   873
      ButtonWidth     =   609
      ButtonHeight    =   714
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.CommandButton cmdedit 
         Height          =   405
         Left            =   600
         Picture         =   "frmTax.frx":0D9E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   1800
         Picture         =   "frmTax.frx":0EE8
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   1215
         Picture         =   "frmTax.frx":141A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   0
         Picture         =   "frmTax.frx":194C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "New"
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton cmdrefresh 
         Height          =   405
         Left            =   2970
         Picture         =   "frmTax.frx":1E7E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmddelete 
         Height          =   405
         Left            =   2385
         Picture         =   "frmTax.frx":1FC8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   465
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5475
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   9657
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
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
      TabCaption(0)   =   "DAILY"
      TabPicture(0)   =   "frmTax.frx":20CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "WEEKLY"
      TabPicture(1)   =   "frmTax.frx":20E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "SEMIMONTHLY"
      TabPicture(2)   =   "frmTax.frx":2102
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "MONTHLY"
      TabPicture(3)   =   "frmTax.frx":211E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "YEARLY"
      TabPicture(4)   =   "frmTax.frx":213A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Table 2018"
      TabPicture(5)   =   "frmTax.frx":2156
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PURPOSE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   765
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(1)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   210
   End
End
Attribute VB_Name = "FrmTaxTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public i As Integer

Dim blSave As Boolean

Private Sub cmdRefresh_Click()
    Call loadlv(SSTab1(1).Caption & "taxtable", SSTab1(0).Caption)
    Call ValidateButton(Me, True, True, False, False, False, True)
    ListView1.Enabled = True
    ListView1.SelectedItem.SELECTED = True
End Sub

Private Sub Form_Load()
    
    Me.Picture = LoadPicture(App.Path & "\pics\bg.gif")
    Me.Top = (mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
    Call SSTab1_Click(1, 1)
    Call ValidateButton(Me, True, True, False, False, False, False)

End Sub

Private Sub cmdAdd_Click()
    Call ValidateButton(Me, False, False, True, True, False, False)
    ListView1.Enabled = False
        
    blSave = True
    SSTab1(2).Visible = True
    Text(0).Enabled = True
    Text(1).Enabled = True
    
    Call clearText
    Text(1).Text = "99999.99"
End Sub

Private Sub save()
    If blSave = True Then
        If CheckText = True Then
            If SSTab1(1).Tab <> 4 Then
                'sqlString = "Insert into " & SSTab1(1).Caption & "taxtable (TAXCODE,salrate2,basetax,perover,USERID) values ('" & SSTab1(0).Caption & "','" & Text(1).Text & "','" & Text(2).Text & "','" & Text(3).Text & "'," & operatorID & ")"
                sqlString = "Insert INTO " & GetTaxTable & " (TAXCODE,salrate2,basetax,perover,USERID) values ('" & SSTab1(0).Caption & "','" & Text(1).Text & "','" & Text(2).Text & "','" & Text(3).Text & "'," & operatorID & ")"
            Else
                '    sqlString = "Insert into " & SSTab1(1).Caption & "taxtable (salrate1,salrate2,basetax,perover,USERID) values ('" & Text(0).Text & "','" & Text(1).Text & "','" & Text(2).Text & "','" & Text(3).Text & "'," & operatorID & ")"
                'sqlString = "Insert into " & SSTab1(1).Caption & "taxtable (salrate2,basetax,perover,USERID) values ('" & Text(1).Text & "','" & Text(2).Text & "','" & Text(3).Text & "'," & operatorID & ")"
                sqlString = "Insert INTO " & GetTaxTable & " (salrate2,basetax,perover,USERID) values ('" & Text(1).Text & "','" & Text(2).Text & "','" & Text(3).Text & "'," & operatorID & ")"
            End If
        Else
            Exit Sub
        End If
    Else
        'sqlString = "Update " & SSTab1(1).Caption & "taxtable  set salrate2='" & Text(1).Text & "',basetax='" & Text(2).Text & "',perover='" & Text(3).Text & "',USERID=" & operatorID & " where id =" & ListView1.SelectedItem.Text & ""
        sqlString = "Update " & GetTaxTable & " set salrate2='" & Text(1).Text & "',basetax='" & Text(2).Text & "',perover='" & Text(3).Text & "',USERID=" & operatorID & " where id =" & ListView1.SelectedItem.Text & ""
    End If
    Call Action_Query(sqlString, cnnstr)
    SSTab1(2).Visible = False
    'Call loadlv(SSTab1(1).Caption & "taxtable", SSTab1(0).Caption)
    Call loadlv(GetTaxTable, SSTab1(0).Caption)
End Sub

Function GetTaxTable() As String
Dim mTable As String

If SSTab1(1).Caption = "Table 2018" Then
    mTable = "semimonthlytax2018"
Else
    mTable = SSTab1(1).Caption & "taxtable"
End If
GetTaxTable = mTable

End Function

Function CheckText() As Boolean
If Val(Text(0)) > Val(Text(1)) Then
MsgBox "Salary rate 1 must be less than salary rate 2"
CheckText = False
Else
CheckText = True

End If
End Function

Private Sub cmdCancel_Click()
    SSTab1(2).Visible = False
    Call ValidateButton(Me, True, True, False, False, False, True)
    ListView1.Enabled = True
    ListView1.SelectedItem.SELECTED = True
End Sub

Private Sub cmdDelete_Click()
Dim strID As Integer

' If SSTab1(1).Tab = 4 Then
If ListView1.SelectedItem.Index < ListView1.ListItems.Count Then
    MsgBox "Cannot remove selected item. Only the last row is allow to delete.", vbExclamation
    Exit Sub
End If
   ' End If
    
If MsgBox("Are you sure you want to remove this item?", vbQuestion + vbYesNo) = vbYes Then
    strID = ListView1.ListItems(ListView1.SelectedItem.Index).Text
    'Call Action_Query("Delete from " & SSTab1(1).Caption & "taxtable where id=" & strID & "", cnnstr)
    Call Action_Query("Delete from " & GetTaxTable & " where id=" & strID & "", cnnstr)
    Call loadlv(SSTab1(1).Caption & "taxtable", SSTab1(0).Caption)
End If

End Sub

Private Sub cmdEdit_Click()
Call ValidateButton(Me, False, False, True, True, False, False)
ListView1.Enabled = False
    
    blSave = False
    'Text(0).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(1).Text
    Text(1).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(1).Text
    Text(2).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(2).Text
    Text(3).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(3).Text
    SSTab1(2).Visible = True
    Text(0).Enabled = False
    Text(1).Enabled = True ' January 2012 updates

End Sub

Private Sub cmdSave_Click()
Dim X As Integer
For X = 1 To 3
    If Text(X).Text = "" Then MsgBox "Fill up the blank box.", vbInformation: Exit Sub
Next

Call ValidateButton(Me, True, True, False, False, False, True)
Call save

End Sub
Private Sub lblExit_Click()
    SSTab1(2).Visible = False
    Call cmdCancel_Click
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmddelete.Enabled = True
    cmdedit.Enabled = True
End Sub

Private Sub SSTab1_Click(Index As Integer, PreviousTab As Integer)
    Dim StrTab As String
    StrTab = SSTab1(0).Caption
    If SSTab1(1).Tab = 0 Then
       Call loadlv("dailytaxtable", StrTab)
    ElseIf SSTab1(1).Tab = 1 Then
       Call loadlv("weeklytaxtable", StrTab)
    ElseIf SSTab1(1).Tab = 2 Then
       Call loadlv("semimonthlytaxtable", StrTab)
    ElseIf SSTab1(1).Tab = 3 Then
       Call loadlv("monthlytaxtable", StrTab)
    ElseIf SSTab1(1).Tab = 4 Then
       Call loadlv("yearlytaxtable")
    ElseIf SSTab1(1).Tab = 5 Then   ' updated 7/14/2018
       Call loadlv("semimonthlytax2018")
    End If
End Sub

Private Sub loadlv(ByVal TableName As String, Optional ByVal StrTab As String)
    Dim rsList As New ADODB.Recordset
    ListView1.ListItems.Clear
        
     TableName = GetTaxTable
     'If SSTab1(1).Tab = 4 Then
     If SSTab1(1).Tab >= 4 Then
         'Set rsList = ReturnRst("Select salrate1 as [RATE FROM],salrate2 AS [RATE TO],basetax AS [BASE TAX], perover AS [PER OVER]  from " & tableName & " order by perover ", Cnnstr, mstatic)
         Set rsList = ReturnRst("Select ID, salrate2 AS [RATE],basetax AS [BASE TAX], perover AS [PER OVER] from " & TableName & " order by ID ", cnnstr, mstatic)
         SSTab1(0).Visible = False
     Else
         'Set rsList = ReturnRst("Select id,basesalary,percentagetax,exemptionamt,bracket from " & tableName & "   where taxstatus='" & StrTab & "' order by id", Cnnstr, mstatic)
         Set rsList = ReturnRst("Select ID,salrate2 AS [RATE],basetax AS [BASE TAX], perover AS [PER OVER] from " & TableName & "   where taxcode='" & StrTab & "' order by ID ", cnnstr, mstatic)
         SSTab1(0).Visible = True
     End If
     
     If rsList.EOF = False Then
         PopulateList ListView1, rsList
     End If
     ListView1.ColumnHeaders(1).Width = 0
     ListView1.ColumnHeaders(2).Width = 2000
     ListView1.ColumnHeaders(3).Width = 2000
    ' ListView1.ColumnHeaders(4).Width = 1600
     If SSTab1(1).Tab <> 3 Then
         ListView1.ColumnHeaders(4).Width = 2000
         ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
     End If
     ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
     ListView1.ColumnHeaders(3).Alignment = lvwColumnRight
    ' ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
     ListView1.Enabled = True

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub clearText()
    Dim X As Integer
    For X = 0 To 3
    Text(X).Text = ""
    Next X
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strValid As String
strValid = "0123456789." & Chr(13) & Chr(8)
If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0

If Index <= 3 Then
    If KeyAscii = 13 Then
        If Index = 3 Then
            Text(1).SetFocus
        Else
            Text(Index + 1).SetFocus
        End If
    End If
End If
End Sub
