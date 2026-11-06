VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmReportViewer 
   Caption         =   "Payroll Summary Report"
   ClientHeight    =   4875
   ClientLeft      =   3780
   ClientTop       =   2910
   ClientWidth     =   7560
   Icon            =   "frmReportViewer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   7560
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      lastProp        =   500
      _cx             =   8070
      _cy             =   5530
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   0   'False
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Call MDIJobRequest.AddToWin(Me.Caption, Me.Name)
'Call mdiMain.AddToWin(Me.Caption, Me.Name)
End Sub

Private Sub Form_Resize()
'comment by bernard 6/8/2012
CRViewer91.Height = Me.Height - 500
CRViewer91.Width = Me.Width - 150
End Sub

Public Sub CommandPass(ByVal strCommand As String)
Select Case strCommand
    Case "Refresh"
        Call Form_Load
    Case "Close"
        Unload Me
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Call MDIJobRequest.RemToWin(Me.Name)
'Call mdiMain.RemToWin(Me.Name)
End Sub
