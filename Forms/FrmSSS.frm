VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSSS 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SSS Table Maintenance"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   Icon            =   "FrmSSS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   8445
   Begin TabDlg.SSTab SSTab1 
      Height          =   1275
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   1935
      Visible         =   0   'False
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   2249
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
      TabPicture(0)   =   "FrmSSS.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblExit"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label8"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text(4)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text(5)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text(6)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text(7)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
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
         Index           =   7
         Left            =   6750
         TabIndex        =   25
         Top             =   540
         Width           =   870
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
         Index           =   6
         Left            =   5760
         TabIndex        =   24
         Top             =   540
         Width           =   870
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
         Index           =   5
         Left            =   4815
         TabIndex        =   23
         Top             =   540
         Width           =   870
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
         Index           =   4
         Left            =   3870
         TabIndex        =   21
         Top             =   540
         Width           =   870
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
         Left            =   2925
         TabIndex        =   8
         Top             =   540
         Width           =   870
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
         Left            =   1980
         TabIndex        =   7
         Top             =   540
         Width           =   870
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
         Left            =   1035
         TabIndex        =   6
         Top             =   540
         Width           =   870
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
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   540
         Width           =   870
      End
      Begin VB.Label Label8 
         Caption         =   "CREDIT"
         Height          =   240
         Left            =   1980
         TabIndex        =   28
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "EC SHARE"
         Height          =   240
         Left            =   6705
         TabIndex        =   27
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label6 
         Caption         =   "MED EE"
         Height          =   240
         Left            =   4770
         TabIndex        =   26
         Top             =   270
         Width           =   690
      End
      Begin VB.Label Label5 
         Caption         =   "MED ER"
         Height          =   240
         Index           =   1
         Left            =   5715
         TabIndex        =   22
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label lblExit 
         BackColor       =   &H8000000B&
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
         Left            =   7470
         MouseIcon       =   "FrmSSS.frx":05A6
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   45
         Width           =   195
      End
      Begin VB.Label Label5 
         Caption         =   "EE SHARE"
         Height          =   240
         Index           =   0
         Left            =   3825
         TabIndex        =   12
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label Label4 
         Caption         =   "ER SHARE"
         Height          =   240
         Left            =   2880
         TabIndex        =   11
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "MAX WAGE"
         Height          =   240
         Left            =   1035
         TabIndex        =   10
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label Label2 
         Caption         =   "MIN WAGE"
         Height          =   240
         Left            =   45
         TabIndex        =   9
         Top             =   270
         Width           =   1320
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3390
      Left            =   225
      TabIndex        =   3
      Top             =   1170
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   5980
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
      Height          =   3960
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   765
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   6985
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
      TabCaption(0)   =   "SSS"
      TabPicture(0)   =   "FrmSSS.frx":08B0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
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
            Picture         =   "FrmSSS.frx":08CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8445
      _ExtentX        =   14896
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
         Picture         =   "FrmSSS.frx":092A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Edit"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdcancel 
         Height          =   405
         Left            =   1800
         Picture         =   "FrmSSS.frx":0A74
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdsave 
         Height          =   405
         Left            =   1215
         Picture         =   "FrmSSS.frx":0FA6
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmdadd 
         Height          =   405
         Left            =   0
         Picture         =   "FrmSSS.frx":14D8
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "New"
         Top             =   0
         Width           =   450
      End
      Begin VB.CommandButton cmdrefresh 
         Height          =   405
         Left            =   2970
         Picture         =   "FrmSSS.frx":1A0A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   465
      End
      Begin VB.CommandButton cmddelete 
         Height          =   405
         Left            =   2385
         Picture         =   "FrmSSS.frx":1B54
         Style           =   1  'Graphical
         TabIndex        =   14
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   0
      Width           =   210
   End
End
Attribute VB_Name = "frmSSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public i As Integer
'
Dim blSave As Boolean

Private Sub cmdRefresh_Click()
Call ValidateButton(Me, True, True, False, False, False, True)
    ListView1.Enabled = True
    ListView1.SelectedItem.SELECTED = True
    
   Call loadlv
End Sub

Private Sub Form_Load()
Call ValidateButton(Me, True, False, False, False, False, False)

    Me.Picture = LoadPicture(App.Path & "\pics\bg.gif")
    Me.Top = (mdiMain.Height - Me.Height) / 8
    Me.Left = (mdiMain.Width - Me.Width) / 2
   Call loadlv
End Sub

Private Sub cmdAdd_Click()
Call ValidateButton(Me, False, False, True, True, False, False)
    ListView1.Enabled = False
    
    blSave = True
    SSTab1(2).Visible = True
    Text(0).Enabled = True
    Text(1).Enabled = True
    
    Call clearText
    Text(1).Text = "999999.99"
    Text(1).Enabled = False
End Sub

Private Sub save()
    If blSave = True Then
          If CheckText = True Then
           sqlString = "Update SSStable  set maxwage=" & Text(0).Text - 0.01 & " WHERE ID= " & getMaxWAGEid("SSStable") & ""
           Call Action_Query(sqlString, cnnstr)
           sqlString = "Insert into SSStable (MINWAGE,MAXWAGE ,CREDIT,ER_SHARE,EE_SHARE,MED_EE,MED_ER,EC_SHARE,USERID ) values (" & Text(0).Text & "," & Text(1).Text & "," & Text(2).Text & "," & Text(3).Text & "," & Text(4).Text & "," & Text(5).Text & "," & Text(6).Text & "," & Text(7).Text & "," & operatorID & ")"
      Else
        Exit Sub
      End If
    Else
        sqlString = "Update SSStable  set minwage=" & Text(0).Text & ",maxwage=" & Text(1).Text & ",CREDIT=" & Text(2).Text & ",ER_SHARE=" & Text(3).Text & ",EE_SHARE=" & Text(4).Text & ",MED_EE=" & Text(5).Text & ",MED_ER=" & Text(6).Text & ",EC_SHARE=" & Text(7).Text & ",USERID=" & operatorID & " where id =" & ListView1.SelectedItem.Text & ""
    End If
    Call Action_Query(sqlString, cnnstr)
    
    SSTab1(2).Visible = False
    Call loadlv
End Sub

'Function getMaxWAGEid() As Integer
'    Dim rs As New ADODB.Recordset
'    Set rs = ReturnRst("Select MAX(ID) from SSStable", Cnnstr, mstatic)
'    If rs.EOF = False Then
''      Dim rs2 As New ADODB.Recordset
''      Set rs2 = ReturnRst("Select MAXWAGE from SSStable WHERE ID=" & rs(0).Value & "", Cnnstr, mstatic)
'    getMaxWAGEid = rs(0)
'    End If
'End Function
Function CheckText() As Boolean
If Val(Text(0)) > Val(Text(1)) Then
MsgBox "Minimum wage must be less than Maximum wage"
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
    If ListView1.SelectedItem.Index < ListView1.ListItems.Count Then
    MsgBox "Cannot remove selected item. Only the last row is allow to delete.", vbExclamation
    Exit Sub
    End If
    
    If MsgBox("Are you sure you want to remove this item?", vbQuestion + vbYesNo) = vbYes Then
        strID = ListView1.ListItems(ListView1.SelectedItem.Index).Text
        Call Action_Query("Delete from SSSTABLE where id=" & strID & "", cnnstr)
        sqlString = "Update SSStable  set maxwage=999999 WHERE ID= " & getMaxWAGEid("SSStable") & ""
         Call Action_Query(sqlString, cnnstr)
        Call loadlv
    End If
End Sub

Private Sub cmdEdit_Click()
Call ValidateButton(Me, False, False, True, True, False, False)
    ListView1.Enabled = False
    
    blSave = False
    Text(0).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(1).Text
    Text(1).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(2).Text
    Text(2).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(3).Text
    Text(3).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(4).Text
    Text(4).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(5).Text
    Text(5).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(6).Text
    Text(6).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(7).Text
    Text(7).Text = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(8).Text
    
    SSTab1(2).Visible = True
'    Text(0).Enabled = False
'    Text(1).Enabled = False
End Sub

Private Sub cmdSave_Click()
Dim X As Integer
For X = 0 To 7
If Text(X).Text = "" Then MsgBox "Fill up the blank box.", vbInformation: Exit Sub
Next

Call ValidateButton(Me, True, True, False, False, False, True)

    ListView1.Enabled = True
    ListView1.SelectedItem.SELECTED = True
    
    Call save
End Sub
Private Sub lblExit_Click()
    SSTab1(2).Visible = False
    Call cmdCancel_Click
End Sub


Private Sub loadlv()
    Dim rsList As New ADODB.Recordset
    ListView1.ListItems.Clear
        
            Set rsList = ReturnRst("Select ID, MINWAGE,MAXWAGE ,CREDIT,ER_SHARE,EE_SHARE,MED_EE,MED_ER,EC_SHARE from SSStable order by ID ", cnnstr, mstatic)
    
        If rsList.EOF = False Then
            PopulateList ListView1, rsList
        End If
        ListView1.ColumnHeaders(1).Width = 0
        ListView1.ColumnHeaders(2).Width = 1000
        ListView1.ColumnHeaders(3).Width = 1000
        ListView1.ColumnHeaders(4).Width = 1000
        ListView1.ColumnHeaders(5).Width = 1000
        ListView1.ColumnHeaders(6).Width = 1000
        ListView1.ColumnHeaders(7).Width = 1000
        ListView1.ColumnHeaders(8).Width = 1000
        ListView1.ColumnHeaders(9).Width = 500
        
        ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
        ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
        ListView1.ColumnHeaders(3).Alignment = lvwColumnRight
        ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
        ListView1.ColumnHeaders(6).Alignment = lvwColumnRight
        ListView1.ColumnHeaders(7).Alignment = lvwColumnRight
        ListView1.ColumnHeaders(8).Alignment = lvwColumnRight
        ListView1.ColumnHeaders(9).Alignment = lvwColumnRight
              ListView1.Enabled = True

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub clearText()
    Dim X As Integer
    For X = 0 To 7
    Text(X).Text = ""
    Next X
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
cmddelete.Enabled = True
cmdedit.Enabled = True
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strValid As String
strValid = "0123456789." & Chr(13) & Chr(8)
If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0

If Index <= 7 Then
    If KeyAscii = 13 Then
        If Index = 7 Then
           Text(0).SetFocus
        Else
           Text(Index + 1).SetFocus
        End If
    End If
End If
End Sub
