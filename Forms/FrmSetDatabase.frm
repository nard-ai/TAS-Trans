VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmSetDatabase 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Database Connection"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "FrmSetDatabase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   45
      Width           =   4515
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      Height          =   525
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   325
      Left            =   1920
      TabIndex        =   4
      Top             =   1635
      Width           =   1095
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Apply"
      Height          =   325
      Left            =   3120
      TabIndex        =   3
      Top             =   1650
      Width           =   1095
   End
   Begin VB.CommandButton CmdBrowse 
      Caption         =   "Browse"
      Height          =   285
      Left            =   4305
      TabIndex        =   2
      Top             =   990
      Width           =   915
   End
   Begin MSComDlg.CommonDialog cdgDialog 
      Left            =   4005
      Top             =   -45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtbrowse 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   345
      TabIndex        =   0
      Top             =   990
      Width           =   3885
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Database here"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   4680
      Picture         =   "FrmSetDatabase.frx":08CA
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   165
      TabIndex        =   1
      Top             =   30
      Width           =   45
   End
End
Attribute VB_Name = "FrmSetDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oldtxtbrowse As String
Dim oldtxtbrowseHR As String

Dim intretval As String

Private Sub cmdbrowse_Click()
    cdgDialog.Flags = cdlOFNFileMustExist
    cdgDialog.Filter = "Database (*.mdb) |*.mdb"
    cdgDialog.FilterIndex = 2
    cdgDialog.Filename = txtbrowse
    cdgDialog.DialogTitle = "Select Access Database"
    cdgDialog.ShowOpen
    txtbrowse.Text = cdgDialog.Filename
    If txtbrowse.Text = "" Then
        txtbrowse.Text = oldtxtbrowse
    End If
End Sub




Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Set cnn = New ADODB.Connection
    StrDbpath = Me.txtbrowse.Text
    If IsFileExists(Me.txtbrowse.Text) = False Then MsgBox "Filename not exist.", vbInformation: txtbrowse.SetFocus: Exit Sub
    If Len(txtbrowse.Text) < 1 Then
            MsgBox "Please specify the path(.mdb)"
            Exit Sub
    End If
  
  cnnstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrDbpath & ""
'   Cnnstr = "Driver={Microsoft Access Driver (*.mdb)};" & _
'                        "Dbq=" & StrDbpath & ";" & _
'                        "Uid=;Pwd=alpha;"

If MsgBox("Are you sure you want to change the database settings? " & vbCrLf & " " & vbCrLf & "The system will logout if you click YES.", vbYesNo + vbQuestion) = vbYes Then
    'Establish the Connection
   'intretval = OpenDatabaseAdo(cnn, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtbrowse.Text)
   intretval = OpenDatabaseAdo(cnn, cnnstr)
   If intretval <> 0 Then
        MsgBox "Unable To Establish the Connection. ", vbCritical
        FrmSetDatabase.Show
        Exit Sub
    Else
        If IsTableEXIST(cnn, "HREMPLOYEES") = Exists Then
            cmdOK.Enabled = True
        Else
            MsgBox "Invalid Database, Select again.", vbCritical
            cmdOK.Enabled = False
            CmdBrowse.SetFocus
            Exit Sub
        End If
            
'        If CompanyCode = 1 Then
'            SaveSetting "DataSearch", "SAPI", "Path2Version2", txtbrowse.Text
'            StrDbpath = GetSetting("DataSearch", "SAPI", "Path2Version2")
'
'
'        ElseIf CompanyCode = 2 Then
'            SaveSetting "DataSearch", "JCI", "Path2Version2", txtbrowse.Text
'            StrDbpath = GetSetting("DataSearch", "JCI", "Path2Version2")
'
'
'
'        ElseIf CompanyCode = 3 Then
'            SaveSetting "DataSearch", "CSPI", "Path2Version2", txtbrowse.Text
'            StrDbpath = GetSetting("DataSearch", "CSPI", "Path2Version2")
'
'
'
'        End If
              
        SaveSetting "DataSearch", CompanyShort, "Path2Version2", txtbrowse.Text
        StrDbpath = GetSetting("DataSearch", CompanyShort, "Path2Version2")
        MsgBox "New Database setting saved.", vbInformation
        Unload mdiMain
        Unload frmScreen
        frmLogin.Show
        Unload Me
        End If
Else
Unload Me
End If
   
End Sub

Private Sub Form_Load()

    Me.Picture = LoadPicture(App.Path & "\pics\bg.gif")
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    txtbrowse.Text = GetSetting("DataSearch", CompanyShort, "Path2Version2")
'    If CompanyCode = 1 Then 'SE
'    txtbrowse.Text = GetSetting("DataSearch", "SAPI", "Path2Version2")
'
'
'    ElseIf CompanyCode = 2 Then 'JCI
'    txtbrowse.Text = GetSetting("DataSearch", "JCI", "Path2Version2")
'
'
'    ElseIf CompanyCode = 3 Then 'CSPI
'    txtbrowse.Text = GetSetting("DataSearch", "CSPI", "Path2Version2")
'
'
'    End If
    
    If txtbrowse.Text <> "" Then
        oldtxtbrowse = txtbrowse.Text
    End If
   

End Sub

Private Sub txtbrowse_Change()
        cmdOK.Enabled = True
End Sub

Private Sub SetToNothing()
    cnn.Close
    Set cnn = Nothing
    Set rsClientTable = Nothing
    cnn.Close
    Set cnn = Nothing
    
    Set rsCompanyTable = Nothing
    Set rsPayrollMast = Nothing
    Set rsPayrollTrans = Nothing
    Set cnPayroll = Nothing
    With frmHRFinalEmployees
        Set .rsHREmp = Nothing
        Set .rs = Nothing
    End With
    
    With frmHRMainEmployee
         Set .rs = Nothing
    End With

End Sub
