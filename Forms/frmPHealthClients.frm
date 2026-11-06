VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPHealthClients 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Clients"
   ClientHeight    =   7380
   ClientLeft      =   4410
   ClientTop       =   3525
   ClientWidth     =   10230
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "frmPHealthClients.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   10230
   Begin TabDlg.SSTab SSTab1 
      Height          =   6885
      Index           =   1
      Left            =   90
      TabIndex        =   0
      Top             =   540
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   12144
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
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
      TabCaption(0)   =   "Employee Basic Salary"
      TabPicture(0)   =   "frmPHealthClients.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   6660
         Index           =   4
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   9855
         Begin MSComctlLib.ListView ListView1 
            Height          =   1695
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   2990
            View            =   3
            Sorted          =   -1  'True
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
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   13
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Client Name"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Remarks"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Processed By"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Processed Date"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Group Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "SBR Number"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2535
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   3240
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   4471
            SortKey         =   1
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   18
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Client Name"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Surname"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Given Name"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Middle Name"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "PhilHealth No."
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Basic Salary"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Bracket"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Personal Share"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Employer Share"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Personal PH"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Employer PH"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Difference"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Client Name"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   16
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   17
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.CommandButton cmdSelectEmployee 
            Caption         =   "View Employees"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5400
            TabIndex        =   14
            Top             =   2760
            Width           =   2055
         End
         Begin VB.Frame Frame1 
            Height          =   735
            Left            =   5040
            TabIndex        =   25
            Top             =   120
            Width           =   4695
            Begin VB.TextBox txtCustomized 
               Height          =   285
               Left            =   4080
               TabIndex        =   29
               Text            =   "10"
               Top             =   240
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.OptionButton optEmployerType 
               Caption         =   "Table Based"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   240
               TabIndex        =   28
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optEmployerType 
               Caption         =   "With 1 Adjustment"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   1560
               TabIndex        =   27
               Top             =   240
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton optEmployerType 
               Caption         =   "Customized Adjustment"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   2880
               TabIndex        =   26
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.CommandButton cmdComputeLess1 
            Caption         =   "Recompute with 1.00 Adjustment"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4680
            TabIndex        =   23
            Top             =   5880
            Width           =   2775
         End
         Begin VB.CommandButton cmdNoPHNo 
            Caption         =   "Check No PHealth No."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5400
            TabIndex        =   20
            Top             =   2760
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CommandButton cmdCheckZero 
            Caption         =   "Check Zero Contributions"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3240
            TabIndex        =   18
            Top             =   2760
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txtRemarks 
            Height          =   285
            Left            =   3120
            TabIndex        =   16
            Top             =   6240
            Width           =   4365
         End
         Begin VB.CommandButton cmdDeleteSelected 
            Caption         =   "Delete Client from List"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7560
            TabIndex        =   13
            Top             =   2760
            Width           =   2055
         End
         Begin VB.CommandButton cmdAddlist 
            Caption         =   "Add to PhilHealth List"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   10
            Top             =   6240
            Width           =   2055
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7560
            TabIndex        =   9
            Top             =   6240
            Width           =   2055
         End
         Begin VB.TextBox txtAppMonth 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2640
            TabIndex        =   6
            Text            =   "AppMonth"
            Top             =   240
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.TextBox txtAppYear 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3960
            TabIndex        =   5
            Text            =   "AppYear"
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2535
            Index           =   2
            Left            =   120
            TabIndex        =   22
            Top             =   3240
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   4471
            SortKey         =   1
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   17
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Client Name"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Surname"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Given Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Middle Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "SSS No."
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Basic Salary"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Bracket"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Employee Share"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Employer Share"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "ECC"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Employee SSS"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Employer SSS"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Difference"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "Client Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   16
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "(Group Code:"
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
            Index           =   4
            Left            =   1200
            TabIndex        =   30
            Top             =   720
            Width           =   3165
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Difference:"
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
            Index           =   6
            Left            =   7560
            TabIndex        =   24
            Top             =   5880
            Width           =   2085
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "No PHealth No."
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
            Index           =   7
            Left            =   2400
            TabIndex        =   21
            Top             =   3000
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Zero Contributions:"
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
            Index           =   2
            Left            =   5160
            TabIndex        =   19
            Top             =   3000
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks:"
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
            Index           =   5
            Left            =   2400
            TabIndex        =   17
            Top             =   6240
            Width           =   765
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Employees:"
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
            Index           =   3
            Left            =   240
            TabIndex        =   15
            Top             =   5880
            Width           =   1965
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "List of Clients"
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
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   1125
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "List of Employees:"
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
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   3000
            Width           =   2445
         End
         Begin VB.Label lblclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Applicable Period:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   38
            Left            =   180
            TabIndex        =   4
            Top             =   240
            Width           =   4485
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9600
      Top             =   600
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
            Picture         =   "frmPHealthClients.frx":05A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   767
      ButtonWidth     =   609
      ButtonHeight    =   609
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.Line Line1 
         BorderColor     =   &H80000001&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   800
         Y1              =   1500
         Y2              =   1200
      End
   End
   Begin VB.Label lblclient 
      BackStyle       =   0  'Transparent
      Caption         =   "Agency Rate:"
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
      Height          =   255
      Index           =   23
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "frmPHealthClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public i As Integer
Dim rsEmployeeBasicData As ADODB.Recordset
Dim rsClientData As ADODB.Recordset
Dim rsCurrentContract As ADODB.Recordset
Dim rsContractHistory As ADODB.Recordset
Dim rsPhilHealth As ADODB.Recordset

Dim StrQuery As String
Dim mRecNo As Integer
Dim mClientCode As Integer
Dim mCompany As String
Dim mEmpName As String
Dim mEmpNo As Integer
Dim ActiveCtr As Integer
Dim InActiveCtr As Integer
Dim RefreshSw As Boolean
Dim paymentbase As String
Dim CurrEmployee As String
Dim CurrAmount As String    ' holds the current value of salaries
Dim AmountSw As Boolean     ' variable for basic salary entries
Dim ActiveSw As Boolean
Dim FormLoadSw As Boolean
Dim ZeroDed As Integer
Dim NoPHNo As Integer
Dim mClientName As String
Dim WithAdjustmentFlag As Integer
Dim mAdjustment As Integer
Dim AddToListFlag As Boolean
Dim WithSBRNoFlag As Boolean

Private Sub cmdCheckZero_Click()
Dim mZeroDed As Double
Dim mListView As ListView

If PHealthFlag = True Then
    Set mListView = ListView1(1)
Else
    Set mListView = ListView1(2)
End If
ZeroDed = 0
For i = 1 To mListView.ListItems.Count
    mZeroDed = Val(mListView.ListItems.Item(i).SubItems(7)) ' get employee share
    If mZeroDed = 0# Then
        mListView.ListItems.Item(i).Checked = True
        ZeroDed = ZeroDed + 1
    End If
Next i
lblclient(2).Caption = "Zero Contributions: " & ZeroDed

End Sub

Private Sub cmdComputeLess1_Click()
Call RecomputeContributions
cmdAddlist.Enabled = True
End Sub

Private Sub cmdDeleteSelected_Click()
Dim res As Variant
Dim itmFound As ListItem   ' FoundItem variable.

If ListView1(0).ListItems.Count = 0 Then
    GoTo ExitRtn1
End If
'mClientName = ListView1(0).SelectedItem.Text
res = MsgBox("Delete Client " & mClientName & "?", vbQuestion + vbYesNo, PayrollAlert)
If res = vbNo Then      ' Do not want to delete
    GoTo ExitRtn1
End If

If PHealthFlag = True Then
    StrQuery = "DELETE FROM PHContributions WHERE clientname = '" & mClientName & _
        "' AND ApplicableMonth = '" & txtAppMonth.Text & "' and ApplicableYear = " & txtAppYear.Text & " AND ProcessedBy = '" & mUserID & "' AND GroupCode = '" & Replace(frmPHealthReport.cmbGroupCode.Text, "New", "") & "'"
Else
    StrQuery = "DELETE FROM SSSContributions WHERE clientname = '" & mClientName & _
        "' AND ApplicableMonth = '" & txtAppMonth.Text & "' and ApplicableYear = " & txtAppYear.Text & " AND ProcessedBy = '" & mUserID & "' AND GroupCode = '" & Replace(frmSSSReport.cmbGroupCode.Text, "New", "") & "'"
End If
'strQuery = "UPDATE PHContributions SET Remarks = 'Deleted.' WHERE CLientName = '" & mClientName & "'"
If Action_Query(StrQuery, cnnstr) = True Then
    If PHealthFlag = True Then
        StrQuery = "SELECT a.Clientcode, a.ClientName, b.Remarks, b.ApplicableMonth, b.ApplicableYear, b.ProcessedDate, b.ProcessedBy, a.ClientFlag3 AS EcolaFlag, b.GroupCode, b.SBRNo " & _
        "FROM clienttable a LEFT JOIN (SELECT DISTINCT AgencyCode, ClientCode, ClientName, ContriCode, Remarks, ApplicableMonth, ApplicableYear, ProcessedDate, ProcessedBy, GroupCode, SBRNo, AmountPaid, DatePaid, DatePosted " & _
        "FROM PHContributions WHERE ApplicableMonth = '" & txtAppMonth.Text & "' AND ApplicableYear = " & txtAppYear.Text & ") AS b " & _
        "ON a.ClientName = b.ClientName " & _
        "WHERE a.clientstatus = 1 and a.clientname is not null ORDER BY a.clientname"
    Else                        ' For SSS
        StrQuery = "SELECT a.Clientcode, a.ClientName, b.Remarks, b.ApplicableMonth, b.ApplicableYear, b.ProcessedDate, b.ProcessedBy, a.ClientFlag3 AS EcolaFlag, b.GroupCode, b.SBRNo " & _
        "FROM clienttable a LEFT JOIN (SELECT DISTINCT AgencyCode, ClientCode, ClientName, ContriCode, Remarks, ApplicableMonth, ApplicableYear, ProcessedDate, ProcessedBy, GroupCode, SBRNo, AmountPaid, DatePaid, DatePosted " & _
        "FROM SSSContributions WHERE ApplicableMonth = '" & txtAppMonth.Text & "' AND ApplicableYear = " & txtAppYear.Text & ") AS b " & _
        "ON a.ClientName = b.ClientName " & _
        "WHERE a.clientstatus = 1 and a.clientname is not null ORDER BY a.clientname"
    End If

    Set rsClientData = ReturnRst(StrQuery, cnnstr, mstatic)
    If rsClientData.RecordCount <> 0 Then
        Call FillClientList2(ListView1(0), rsClientData)
        rsClientData.MoveFirst
    End If
    Call DeleteClientFromList
    MsgBox "Client " & mClientName & " was successfully deleted.", vbInformation, PayrollAlert
    cmdDeleteSelected.Enabled = False
    AddToListFlag = True
Else
    MsgBox "Deletion of client failed.", vbInformation, PayrollAlert
End If

ExitRtn1:

Set itmFound = ListView1(0). _
    FindItem(mClientName, lvwText, , lvwPartial)
   
   ' If no ListItem is found, then inform user and exit. If a
   ' ListItem is found, scroll the control using the EnsureVisible
   ' method, and select the ListItem.
   If itmFound Is Nothing Then  ' If no match, inform user and exit.
   Else
       itmFound.EnsureVisible ' Scroll ListView to show found ListItem.
       itmFound.SELECTED = True   ' Select the ListItem.
      ' Return focus to the control to see selection.
       ListView1(0).SetFocus
   End If
mClientName = ListView1(0).SelectedItem.Text

End Sub

Private Sub DeleteClientFromList()

Dim litem1 As ListItem
Dim Listctr As Integer
Dim rsPHList As ADODB.Recordset
Dim mListView As ListView

If PHealthFlag = True Then
    Set mListView = ListView1(1)
Else
    Set mListView = ListView1(2)
End If

Listctr = 0
If PHealthFlag = True Then
    StrQuery = "SELECT * FROM PHContributions " & _
        " WHERE ApplicableMonth = '" & txtAppMonth.Text & "' and ApplicableYear = " & txtAppYear.Text & " AND ProcessedBy = '" & mUserID & "' AND GroupCode = '" & Replace(frmPHealthReport.cmbGroupCode.Text, "New", "") & "' ORDER BY Surname, Firstname"
Else
    StrQuery = "SELECT * FROM SSSContributions " & _
        " WHERE ApplicableMonth = '" & txtAppMonth.Text & "' and ApplicableYear = " & txtAppYear.Text & " AND ProcessedBy = '" & mUserID & "' AND GroupCode = '" & Replace(frmSSSReport.cmbGroupCode.Text, "New", "") & "' ORDER BY Surname, Firstname"
End If
Set rsPHList = ReturnRst(StrQuery, cnnstr, mstatic)
If rsPHList.RecordCount = 0 Then
    If PHealthFlag = True Then
        frmPHealthReport.ListView1(1).ListItems.Clear
        ListView1(1).ListItems.Clear
    Else
        frmSSSReport.ListView1(0).ListItems.Clear
        ListView1(2).ListItems.Clear
    End If
    GoTo Exit22
End If

With rsPHList
If PHealthFlag = True Then
    frmPHealthReport.ListView1(1).ListItems.Clear
    ListView1(1).ListItems.Clear
Else
    frmSSSReport.ListView1(0).ListItems.Clear
    ListView1(2).ListItems.Clear
End If
While Not (.EOF Or .BOF)
    If PHealthFlag = True Then
        Set litem1 = frmPHealthReport.ListView1(1).ListItems.Add(, , IIf(IsNull(!ClientName), "Blank", !ClientName))
    Else
        Set litem1 = frmSSSReport.ListView1(0).ListItems.Add(, , IIf(IsNull(!ClientName), "Blank", !ClientName))
    End If
        litem1.SubItems(1) = IIf(IsNull(!Surname), "", !Surname)
        litem1.SubItems(2) = IIf(IsNull(!FirstName), "", !FirstName)
        litem1.SubItems(3) = IIf(IsNull(!MiddleName), "", !MiddleName)
        litem1.SubItems(4) = IIf(IsNull(!PHealthSSSNo), "", !PHealthSSSNo)
        litem1.SubItems(5) = Format(!BasicPay, "0.##")
        litem1.SubItems(6) = !SalaryBracket
        litem1.SubItems(7) = Format(!ActualEmployeeDed, "0.##")
        litem1.SubItems(8) = Format(!ActualEmployerDed, "0.##")
        If PHealthFlag = True Then
            litem1.SubItems(9) = Format(!AdjustedEmployeeDed, "0.##")
            litem1.SubItems(10) = Format(!AdjustedEmployerDed, "0.##")
            litem1.SubItems(11) = Format(!AmtDifference, "0.##")
        Else
            litem1.SubItems(9) = Format(!AdjustedECC, "0.##")
            litem1.SubItems(10) = Format(!AdjustedEmployeeDed, "0.##")
            litem1.SubItems(11) = Format(!AdjustedEmployerDed, "0.##")
            litem1.SubItems(12) = Format(!AmtDifference, "0.##")
            litem1.SubItems(14) = Format(!TableBasedEE, "0.##")
            litem1.SubItems(15) = Format(!TableBasedER, "0.##")
        End If
'        If !ExcludedFlag = 1 Then
'            'frmPHealthReport.ListView1(1).ListItems.Item(i).Checked = True
'            frmPHealthReport.ListView1(1).ListItems.Checked = True
'        End If
    .MoveNext
Wend
End With
Set rsPHList = Nothing
Call GetTotals

Exit22:

End Sub

Private Sub cmdNoPHNo_Click()
Dim mListView As ListView

If PHealthFlag = True Then
    Set mListView = ListView1(1)
Else
    Set mListView = ListView1(2)
End If

NoPHNo = 0
For i = 1 To mListView.ListItems.Count
    If mListView.ListItems.Item(i).SubItems(4) = "" Or mListView.ListItems.Item(i).SubItems(4) = "__-_______-_" Then ' get philhealth number
        mListView.ListItems.Item(i).Checked = True
        NoPHNo = NoPHNo + 1
    End If
Next i
If PHealthFlag = True Then
    lblclient(7).Caption = "No PHealth No.: " & NoPHNo
    txtRemarks.Text = "Total Employees: " & mListView.ListItems.Count
Else
    lblclient(7).Caption = "No SSS No.: " & NoPHNo
    txtRemarks.Text = "Total Employees: " & mListView.ListItems.Count
End If
End Sub

Private Sub RecomputeContributions(Optional TableFlag As Boolean = False)
Dim mBasicPay As Double
Dim mZeroDed As Integer
Dim mActualEE As Double
Dim mNewEE As Double
Dim mListView As ListView

If PHealthFlag = True Then
    Set mListView = ListView1(1)
Else
    Set mListView = ListView1(2)
End If

Call cmdCheckZero_Click
mZeroDed = 0
mActualEE = 0
mNewEE = 0
For i = 1 To mListView.ListItems.Count
    If mListView.ListItems.Item(i).Checked = False Then ' recompute only those uncheck
        mBasicPay = Val(mListView.ListItems.Item(i).SubItems(5)) ' get basic pay
        mZeroDed = Val(mListView.ListItems.Item(i).SubItems(7)) ' get employee share
'        If mZeroDed = 0 Then
'            ZeroDed = ZeroDed + 1
'        End If
ReadPHTable:
        If PHealthFlag = True Then
            StrQuery = "SELECT * FROM PhilHealthTable WHERE " & mBasicPay & " between minwage and maxwage"
        Else
            StrQuery = "SELECT bracket as code, minwage, maxwage, er_share as ph_er, ee_share as ph_ee FROM SSSTable WHERE " & mBasicPay & " between minwage and maxwage"
        End If
        Set rsPhilHealth = ReturnRst(StrQuery, cnnstr, mstatic)
        If rsPhilHealth.EOF = False Then
            If WithAdjustmentFlag = 2 And (rsPhilHealth!ph_ee - 1) > Val(mListView.ListItems.Item(i).SubItems(7)) Then
                mBasicPay = rsPhilHealth!minwage - 100
                Set rsPhilHealth = Nothing
                GoTo ReadPHTable
            ElseIf WithAdjustmentFlag = 3 And (rsPhilHealth!ph_ee) > (Val(mListView.ListItems.Item(i).SubItems(7)) - mAdjustment) Then
'                If PHealthFlag = True Then
'                    mActualEE = Val(mListView.ListItems.Item(i).SubItems(7)) - Val(mListView.ListItems.Item(i).SubItems(9))
'                Else
'                    mActualEE = Val(mListView.ListItems.Item(i).SubItems(7)) - Val(mListView.ListItems.Item(i).SubItems(10))
'                End If
'                If mActualEE < mAdjustment Then
                    mBasicPay = rsPhilHealth!minwage - 100
                    Set rsPhilHealth = Nothing
                    GoTo ReadPHTable
                'End If
            End If
            mListView.ListItems.Item(i).SubItems(6) = IIf(IsNull(rsPhilHealth!code), 0, rsPhilHealth!code)
            If PHealthFlag = True Then
                mListView.ListItems.Item(i).SubItems(9) = rsPhilHealth!ph_ee
                mListView.ListItems.Item(i).SubItems(10) = rsPhilHealth!ph_er
                mListView.ListItems.Item(i).SubItems(11) = Format(Val(mListView.ListItems.Item(i).SubItems(7)) - Val(mListView.ListItems.Item(i).SubItems(9)), "0.##")
                mActualEE = mActualEE + Val(mListView.ListItems.Item(i).SubItems(7))
                mNewEE = mNewEE + Val(mListView.ListItems.Item(i).SubItems(9))
            Else
                If TableFlag = False Then
                    mListView.ListItems.Item(i).SubItems(10) = rsPhilHealth!ph_ee
                    mListView.ListItems.Item(i).SubItems(11) = rsPhilHealth!ph_er
                    mListView.ListItems.Item(i).SubItems(12) = Format(Val(mListView.ListItems.Item(i).SubItems(7)) - Val(mListView.ListItems.Item(i).SubItems(10)), "0.##")
                    mActualEE = mActualEE + Val(mListView.ListItems.Item(i).SubItems(7))
                    mNewEE = mNewEE + Val(mListView.ListItems.Item(i).SubItems(10))
                Else
                    mListView.ListItems.Item(i).SubItems(14) = rsPhilHealth!ph_ee
                    mListView.ListItems.Item(i).SubItems(15) = rsPhilHealth!ph_er
                End If
            End If
        End If
    End If
Next i
If TableFlag = False Then
    lblclient(6).Caption = "Difference:   PhP " & Format(mActualEE - mNewEE, "0.##")
    lblclient(2).Caption = "Zero Contributions: " & ZeroDed
    txtRemarks.Text = "Total Employees: " & mListView.ListItems.Count
    MsgBox "Recomputation done.", vbInformation, PayrollAlert
End If
End Sub

Private Sub cmdSelectEmployee_Click()
    mClientName = ListView1(0).SelectedItem.Text
    
    If ListView1(0).ListItems.Count >= 1 Then
        Call DisplayClientData(ListView1(0).SelectedItem.Text, False, 0)
    End If
    If PHealthFlag = True Then
        lblclient(3).Caption = "Total Employees: " & ListView1(1).ListItems.Count
        txtRemarks.Text = "Total Employees: " & ListView1(1).ListItems.Count
    Else
        lblclient(3).Caption = "Total Employees: " & ListView1(2).ListItems.Count
        txtRemarks.Text = "Total Employees: " & ListView1(2).ListItems.Count
    End If
    cmdDeleteSelected.Enabled = True
    cmdComputeLess1.Enabled = True
    cmdAddlist.Enabled = False
    ListView1(0).FindItem mClientName
End Sub

Private Sub Form_Activate()
Dim mComp As String
Dim mEmpN As String

    If ActiveSw = False Then
        Me.Top = (mdiMain.Height - Me.Height) / 8
        Me.Left = (mdiMain.Width - Me.Width) / 2
    End If
    FormLoadSw = False
    mAdjustment = 1
    WithAdjustmentFlag = 2
    cmdAddlist.Enabled = False

    If PHealthFlag = True Then
        frmPHealthReport.Visible = True
        'frmSSSReport.Visible = False
        lblclient(7).Caption = "No PHealth No."
        cmdNoPHNo.Caption = "Check No PHealth No."
        cmdAddlist.Caption = "Add to PhilHealth List"
        Me.Caption = "Client List for PhilHealth Report for the month of " & frmPHealthReport.cmdAppMonth.Text & " " & frmPHealthReport.cmdAppYear.Text
        lblclient(38).Caption = "Applicable Month: " & frmPHealthReport.cmdAppMonth.Text & " " & frmPHealthReport.cmdAppYear.Text
    Else
        'frmPHealthReport.Visible = False
        frmSSSReport.Visible = True
        lblclient(7).Caption = "No SSS No."
        cmdNoPHNo.Caption = "Check No SSS No."
        cmdAddlist.Caption = "Add to SSS List"
        Me.Caption = "Client List for SSS Report for the month of " & frmSSSReport.cmdAppMonth.Text & " " & frmSSSReport.cmdAppYear.Text
        lblclient(38).Caption = "Applicable Month: " & frmSSSReport.cmdAppMonth.Text & " " & frmSSSReport.cmdAppYear.Text
    End If
    ActiveSw = True
    RefreshSw = False
    If PHealthFlag = True Then
        'Me.Caption = "Select Clients for PhilHealth Report"
        txtAppMonth.Text = frmPHealthReport.cmdAppMonth.Text
        txtAppYear.Text = frmPHealthReport.cmdAppYear.Text
        lblclient(4).Caption = "(Group Code: " & frmPHealthReport.cmbGroupCode.Text & ")"
    Else
        'Me.Caption = "Select Clients for SSS Report"
        txtAppMonth.Text = frmSSSReport.cmdAppMonth.Text
        txtAppYear.Text = frmSSSReport.cmdAppYear.Text
        lblclient(4).Caption = "(Group Code: " & frmSSSReport.cmbGroupCode.Text & ")"
    End If
    Call GetClientData
    FormLoadSw = True
    txtRemarks.Text = "Total Employees: 0"
End Sub

Private Sub ListView1_Click(Index As Integer)
    'Call DisplayClientData(ListView1(0).SelectedItem.Text, False, Index)
'    If Index = 0 Then
'        lblclient(2).Caption = "Name: " & ListView1(1).SelectedItem.Text
'    End If
End Sub

Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    'Call DisplayClientData(Item.Text, False, Index)
    mClientName = ListView1(0).SelectedItem.Text
    WithSBRNoFlag = False
    cmdComputeLess1.Enabled = False
    If Index = 0 And Item.ListSubItems(1).Text <> "" And Item.ListSubItems(2).Text <> mUserID Then
        cmdDeleteSelected.Enabled = False
        cmdSelectEmployee.Enabled = False
    Else
        cmdDeleteSelected.Enabled = True
        cmdSelectEmployee.Enabled = True
        If Index = 0 And Item.ListSubItems(1).Text <> "" And Item.ListSubItems(1).Text <> "Deleted." Then
            If Index = 0 And Item.ListSubItems(5).Text <> "" Then ' with SBR# already, do not allow changes
                cmdDeleteSelected.Enabled = False
                AddToListFlag = False
                WithSBRNoFlag = True
            Else
                cmdDeleteSelected.Enabled = True
                AddToListFlag = False
            End If
            'cmdAddlist.Enabled = False
        Else
            cmdDeleteSelected.Enabled = False
            AddToListFlag = True
            'cmdAddlist.Enabled = True
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Set rsClientData = Nothing
    Set rsEmployeeBasicData = Nothing
    Set rsCurrentContract = Nothing
    Set rsContractHistory = Nothing
    Unload Me
End Sub

Private Sub FillEmployeeList(ByVal ClientName As String, ByRef mLsEmp As ListView, _
                        ByRef mRs As Recordset, Optional AppendSw As Boolean = True)
Dim litem1 As ListItem
Dim Listctr As Integer

    Listctr = 0
    mLsEmp.View = lvwReport
    'ListView1(0).View = lvwReport
    'ListView1(1).ListItems.Clear
    mLsEmp.ListItems.Clear
    mRs.MoveFirst
    If PHealthFlag = True Then ' for philhealth
        While Not (mRs.EOF Or mRs.BOF)
            'Set litem1 = ListView1(1).ListItems.Add(, , mRs!ClientCode)
            Set litem1 = ListView1(1).ListItems.Add(, , IIf(IsNull(ClientName), "Blank", ClientName))
            'litem1.SubItems(1) = IIf(IsNull(ClientName), "Blank", ClientName)
            litem1.SubItems(1) = IIf(IsNull(mRs!Surname), "", mRs!Surname)
            litem1.SubItems(2) = IIf(IsNull(mRs!FirstName), "", mRs!FirstName)
            litem1.SubItems(3) = IIf(IsNull(mRs!MiddleName), "", mRs!MiddleName)
            litem1.SubItems(4) = IIf(IsNull(mRs!PhilHealth), "", mRs!PhilHealth)
            litem1.SubItems(5) = Format(mRs!BasicPay, "0.##")
            litem1.SubItems(6) = ""
            litem1.SubItems(7) = Format(mRs!PHealthEmployee, "0.##")
            litem1.SubItems(8) = Format(mRs!PHealthEmployer, "0.##")
            litem1.SubItems(9) = 0#   'Format(mRs!PHealthEmployee, "0.##")
            litem1.SubItems(10) = 0#   'Format(mRs!PHealthEmployer, "0.##")
            litem1.SubItems(11) = 0#   'Format(mRs!PHealthEmployer, "0.##")
            litem1.SubItems(12) = IIf(IsNull(ClientName), "Blank", ClientName)
            litem1.SubItems(13) = Replace(frmPHealthReport.cmbGroupCode.Text, "New", "")
            ' Save employee number.
            litem1.SubItems(14) = IIf(IsNull(mRs!StartContract) Or mRs!StartContract = #1/1/1900#, "__/__/____", Format(mRs!StartContract, "mm/dd/yyyy"))
            litem1.SubItems(15) = IIf(IsNull(mRs!empPosition), "", mRs!empPosition)
            litem1.SubItems(16) = mRs!EmployeeID
            Listctr = Listctr + 1
            mRs.MoveNext
        Wend
    Else            ' for SSS
        While Not (mRs.EOF Or mRs.BOF)
            Set litem1 = ListView1(2).ListItems.Add(, , IIf(IsNull(ClientName), "Blank", ClientName))
            litem1.SubItems(1) = IIf(IsNull(mRs!Surname), "", mRs!Surname)
            litem1.SubItems(2) = IIf(IsNull(mRs!FirstName), "", mRs!FirstName)
            litem1.SubItems(3) = IIf(IsNull(mRs!MiddleName), "", mRs!MiddleName)
            litem1.SubItems(4) = IIf(IsNull(mRs!sss), "", mRs!sss)
            litem1.SubItems(5) = Format(mRs!BasicPay, "0.##")
            litem1.SubItems(6) = ""
            litem1.SubItems(7) = Format(mRs!SSSEmployee, "0.##")
            litem1.SubItems(8) = IIf(mRs!SSSEmployer <> 0, Format(mRs!SSSEmployer - IIf(mRs!ECCEmployer = 0, 10, mRs!ECCEmployer), "0.##"), 0#)
            litem1.SubItems(9) = IIf(mRs!SSSEmployer <> 0, IIf(mRs!ECCEmployer = 0, 10, mRs!ECCEmployer), 0#) ' ECC
            litem1.SubItems(10) = 0#   'Format(mRs!PHealthEmployee, "0.##")
            litem1.SubItems(11) = 0#   'Format(mRs!PHealthEmployer, "0.##")
            litem1.SubItems(12) = 0#   'Difference
            litem1.SubItems(13) = IIf(IsNull(ClientName), "Blank", ClientName)
            'litem1.SubItems(14) = 'table based ee
            'litem1.SubItems(15) = 'table based er
            litem1.SubItems(16) = Replace(frmSSSReport.cmbGroupCode.Text, "New", "")
            Listctr = Listctr + 1
            mRs.MoveNext
        Wend
        ' Recompute based from table.
        WithAdjustmentFlag = 1
        mAdjustment = 0
        Call RecomputeContributions(True)
        ' Reset the setting to default.
        WithAdjustmentFlag = 2
        mAdjustment = 1
    End If
    MsgBox "There were " & Listctr & " records selected.", vbInformation, PayrollAlert
End Sub

Private Sub FillContractList(ByRef mLsEmp As ListView, _
                        ByRef mRs As Recordset)
Dim litem1 As ListItem
Dim litem2 As ListItem

    mLsEmp.View = lvwReport
    'mLsEmp.ListItems.Clear
    While Not (mRs.EOF Or mRs.BOF)
        Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!FullName), "Blank Name", mRs!FullName))
            litem1.SubItems(1) = mRs(0)
        mRs.MoveNext
    Wend
End Sub

Private Sub cmdAddList_Click()  ' add to philhealth list
Dim litem1 As ListItem
Dim Listctr As Integer
Dim PHNo As String
Dim ActualEE As Double
Dim ActualER As Double
Dim AdjustedEE As Double, TableBasedEE As Double
Dim AdjustedER As Double, TableBasedER As Double
Dim AdjustedECC As Double
Dim mSurname As String
Dim mFirstName As String
Dim mMiddlename As String
Dim mBasicPay As Double
Dim mBracket As String, mContractDte As String
Dim mDiff As Double
Dim ExcludedFlag As Integer
Dim mListView As ListView
Dim itmFound As ListItem   ' FoundItem variable.

If PHealthFlag = True Then
    Set mListView = ListView1(1)
Else
    Set mListView = ListView1(2)
End If

If mListView.ListItems.Count = 0 Then
    GoTo ExitRtn2
End If

If WithSBRNoFlag = True Then
    MsgBox "This client was already paid. Cannot add to list!", vbInformation, PayrollAlert
    GoTo ExitRtn2
End If
If AddToListFlag = False Then
    MsgBox "This client was already appended. Please delete it first and reprocess again!", vbInformation, PayrollAlert
    GoTo ExitRtn2
End If

Listctr = 0
'mClientName = ListView1(0).SelectedItem.Text
For i = 1 To mListView.ListItems.Count
        If PHealthFlag = True Then
            Set litem1 = frmPHealthReport.ListView1(1).ListItems.Add(, , mListView.ListItems.Item(i).SubItems(12))
        Else
            Set litem1 = frmSSSReport.ListView1(0).ListItems.Add(, , mListView.ListItems.Item(i).SubItems(13))
        End If
            
        If mListView.ListItems.Item(i).Checked = True Then  ' get only those without check
            'frmPHealthReport.mlistview.ListItems.Item(i).Checked = True
            ExcludedFlag = 1
        Else
            ExcludedFlag = 0
        End If
        mSurname = mListView.ListItems.Item(i).SubItems(1)
        mFirstName = mListView.ListItems.Item(i).SubItems(2)
        mMiddlename = mListView.ListItems.Item(i).SubItems(3)
        
        litem1.SubItems(1) = mSurname
        litem1.SubItems(2) = mFirstName
        litem1.SubItems(3) = mMiddlename
        PHNo = mListView.ListItems.Item(i).SubItems(4) ' PhilHealth no
        litem1.SubItems(4) = PHNo
        mBasicPay = mListView.ListItems.Item(i).SubItems(5)
        mBracket = mListView.ListItems.Item(i).SubItems(6)
        litem1.SubItems(5) = mBasicPay
        litem1.SubItems(6) = mBracket
        
        ActualEE = mListView.ListItems.Item(i).SubItems(7)
        ActualER = mListView.ListItems.Item(i).SubItems(8)
        If PHealthFlag = True Then
            'AdjustedEE = mListView
            AdjustedEE = mListView.ListItems.Item(i).SubItems(9)
            AdjustedER = mListView.ListItems.Item(i).SubItems(10)
        Else
            AdjustedECC = mListView.ListItems.Item(i).SubItems(9)
            AdjustedEE = mListView.ListItems.Item(i).SubItems(10)
            AdjustedER = mListView.ListItems.Item(i).SubItems(11)
            TableBasedEE = IIf(mListView.ListItems.Item(i).SubItems(14) = "", 0, mListView.ListItems.Item(i).SubItems(14))
            TableBasedER = IIf(mListView.ListItems.Item(i).SubItems(15) = "", 0, mListView.ListItems.Item(i).SubItems(15))
        End If
        
        litem1.SubItems(7) = ActualEE
        litem1.SubItems(8) = ActualER
        If PHealthFlag = True Then
            litem1.SubItems(9) = AdjustedEE
            litem1.SubItems(10) = AdjustedER
            mDiff = mListView.ListItems.Item(i).SubItems(11)
            litem1.SubItems(11) = mDiff
            litem1.SubItems(12) = mListView.ListItems.Item(i).SubItems(12) ' include the client name, please add
            litem1.SubItems(13) = mListView.ListItems.Item(i).SubItems(13) ' include Group Code
            
            litem1.SubItems(16) = mListView.ListItems.Item(i).SubItems(14) ' Start Contract
            litem1.SubItems(17) = mListView.ListItems.Item(i).SubItems(15) ' Employee Position
            litem1.SubItems(18) = mListView.ListItems.Item(i).SubItems(16) ' Employee Number
            
            mContractDte = IIf(litem1.SubItems(16) = "__/__/____", "01/01/1900", Format(litem1.SubItems(16), "mm/dd/yyyy"))
            StrQuery = "INSERT INTO PHContributions (EmployeeID, StartContract, EmpPosition, AgencyCode, ClientCode, ClientName, Surname, Firstname, Middlename, ContriCode, PhealthSSSNo, BasicPay, SalaryBracket, ActualEmployeeDed, ActualEmployerDed, AdjustedEmployeeDed, AdjustedEmployerDed, AmtDifference, GroupCode, Remarks, ExcludedFlag, ApplicableMonth, ApplicableYear, ProcessedDate, ProcessedBy) " & _
                " VALUES (" & litem1.SubItems(18) & ", '" & mContractDte & "', '" & litem1.SubItems(17) & "', " & CompanyCode & ", 1, '" & mClientName & "', '" & mSurname & "','" & mFirstName & "','" & mMiddlename & "', 'PH', '" & PHNo & "', " & mBasicPay & ", '" & mBracket & "', " & ActualEE & ", " & ActualER & ", " & AdjustedEE & ", " & AdjustedER & ", " & mDiff & ", '" & Replace(frmPHealthReport.cmbGroupCode.Text, "New", "") & "', '" & txtRemarks.Text & "', " & ExcludedFlag & ", '" & txtAppMonth.Text & "', " & txtAppYear.Text & ", '" & Format(Now, "mm/dd/yyyy") & "', '" & mUserID & "')"
        Else
            litem1.SubItems(9) = AdjustedECC
            litem1.SubItems(10) = AdjustedEE
            litem1.SubItems(11) = AdjustedER
            mDiff = mListView.ListItems.Item(i).SubItems(12)
            litem1.SubItems(12) = mDiff
            litem1.SubItems(13) = mListView.ListItems.Item(i).SubItems(13) ' include the client name
            
            litem1.SubItems(14) = TableBasedEE  ' table based EE
            litem1.SubItems(15) = TableBasedER  ' table based ER
            litem1.SubItems(16) = mListView.ListItems.Item(i).SubItems(16) ' include Group Code
            
            StrQuery = "INSERT INTO SSSContributions (AgencyCode, ClientCode, ClientName, Surname, Firstname, Middlename, ContriCode, PhealthSSSNo, BasicPay, SalaryBracket, ActualEmployeeDed, ActualEmployerDed, ActualECC, AdjustedEmployeeDed, AdjustedEmployerDed, AdjustedECC, AmtDifference, TableBasedEE, TableBasedER, GroupCode, Remarks, ExcludedFlag, ApplicableMonth, ApplicableYear, ProcessedDate, ProcessedBy) " & _
                " VALUES (" & CompanyCode & ", 1, '" & mClientName & "', '" & mSurname & "', '" & mFirstName & "', '" & mMiddlename & "', 'SSS', '" & PHNo & "', " & mBasicPay & ", '" & mBracket & "', " & ActualEE & ", " & ActualER & ", " & AdjustedECC & ", " & AdjustedEE & ", " & AdjustedER & ", " & AdjustedECC & ", " & mDiff & ", " & TableBasedEE & ", " & TableBasedER & ", '" & Replace(frmSSSReport.cmbGroupCode.Text, "New", "") & "', '" & txtRemarks.Text & "', " & ExcludedFlag & ", '" & txtAppMonth.Text & "', " & txtAppYear.Text & ", '" & Format(Now, "mm/dd/yyyy") & "', '" & mUserID & "')"
        End If
        Call Action_Query(StrQuery, cnnstr)
    
    Listctr = Listctr + 1
Next i
If Listctr <> 0 Then
    If PHealthFlag = True Then
        StrQuery = "SELECT a.Clientcode, a.ClientName, b.Remarks, b.ApplicableMonth, b.ApplicableYear, b.ProcessedDate, b.ProcessedBy, a.ClientFlag3 AS EcolaFlag, b.GroupCode, b.SBRNo " & _
        "FROM clienttable a LEFT JOIN (SELECT DISTINCT AgencyCode, ClientCode, ClientName, ContriCode, Remarks, ApplicableMonth, ApplicableYear, ProcessedDate, ProcessedBy, GroupCode, SBRNo, AmountPaid, DatePaid, DatePosted " & _
        "FROM PHContributions WHERE ApplicableMonth = '" & txtAppMonth.Text & "' AND ApplicableYear = " & txtAppYear.Text & ") AS b " & _
        "ON a.ClientName = b.ClientName " & _
        "WHERE a.clientstatus = 1 and a.clientname is not null ORDER BY a.clientname"
    Else                        ' For SSS
        StrQuery = "SELECT a.Clientcode, a.ClientName, b.Remarks, b.ApplicableMonth, b.ApplicableYear, b.ProcessedDate, b.ProcessedBy, a.ClientFlag3 AS EcolaFlag, b.GroupCode, b.SBRNo " & _
        "FROM clienttable a LEFT JOIN (SELECT DISTINCT AgencyCode, ClientCode, ClientName, ContriCode, Remarks, ApplicableMonth, ApplicableYear, ProcessedDate, ProcessedBy, GroupCode, SBRNo, AmountPaid, DatePaid, DatePosted " & _
        "FROM SSSContributions WHERE ApplicableMonth = '" & txtAppMonth.Text & "' AND ApplicableYear = " & txtAppYear.Text & ") AS b " & _
        "ON a.ClientName = b.ClientName " & _
        "WHERE a.clientstatus = 1 and a.clientname is not null ORDER BY a.clientname"
    End If

    Set rsClientData = ReturnRst(StrQuery, cnnstr, mstatic)
    If rsClientData.RecordCount <> 0 Then
        Call FillClientList2(ListView1(0), rsClientData)
        rsClientData.MoveFirst
    End If
    AddToListFlag = False
End If
Call GetTotals
If PHealthFlag = True Then
    MsgBox "There were " & Listctr & " records added to PhilHealth.", vbInformation, PayrollAlert
Else
    MsgBox "There were " & Listctr & " records added to SSS.", vbInformation, PayrollAlert
End If
cmdAddlist.Enabled = False  ' To prevent double adding of the same company

ExitRtn2:
  
Set itmFound = ListView1(0). _
    FindItem(mClientName, lvwText, , lvwPartial)
   
   ' If no ListItem is found, then inform user and exit. If a
   ' ListItem is found, scroll the control using the EnsureVisible
   ' method, and select the ListItem.
   If itmFound Is Nothing Then  ' If no match, inform user and exit.
   Else
       itmFound.EnsureVisible ' Scroll ListView to show found ListItem.
       itmFound.SELECTED = True   ' Select the ListItem.
      ' Return focus to the control to see selection.
       ListView1(0).SetFocus
   End If
mClientName = ListView1(0).SelectedItem.Text

End Sub

Private Sub GetTotals()
Dim rsTotal As ADODB.Recordset

If PHealthFlag = True Then
    StrQuery = "SELECT SUM(AdjustedEmployeeDed) AS AdjustedEE, SUM(ADjustedEmployerDed) AS AdjustedER FROM PHContributions WHERE ApplicableMonth = '" & txtAppMonth.Text & "' and ApplicableYear = " & txtAppYear.Text & " AND ProcessedBy = '" & mUserID & "'"
    Set rsTotal = ReturnRst(StrQuery, cnnstr, mstatic)
    If rsTotal.RecordCount <> 0 Then
        frmPHealthReport.txtSubTotal.Text = Format(rsTotal!AdjustedEE, "0.##")
        frmPHealthReport.txtGrandTotal.Text = Format(rsTotal!AdjustedER, "0.##")
    Else
        frmPHealthReport.txtSubTotal.Text = 0#
        frmPHealthReport.txtGrandTotal.Text = 0#
    End If
Else
    StrQuery = "SELECT SUM(AdjustedEmployeeDed) AS AdjustedEE, SUM(ADjustedEmployerDed) AS AdjustedER FROM SSSContributions WHERE ApplicableMonth = '" & txtAppMonth.Text & "' and ApplicableYear = " & txtAppYear.Text & " AND ProcessedBy = '" & mUserID & "'"
    Set rsTotal = ReturnRst(StrQuery, cnnstr, mstatic)
    If rsTotal.RecordCount <> 0 Then
        frmSSSReport.txtSubTotal.Text = Format(rsTotal!AdjustedEE, "0.##")
        frmSSSReport.txtGrandTotal.Text = Format(rsTotal!AdjustedER, "0.##")
    Else
        frmSSSReport.txtSubTotal.Text = 0#
        frmSSSReport.txtGrandTotal.Text = 0#
    End If
End If
Set rsTotal = Nothing

End Sub

Private Sub cmdPhilHealth_Click()
'    If ReportType = 4 Then
'        Call GeneratePhilHealthReports(rsDeductions)
'    Else
'        MsgBox "Select PhilHealth Button first, before you can generate report.", vbInformation, PayrollAlert
'    End If
End Sub

Private Sub cmdClearAll_Click()
'    If ReportType = 4 Then
'        ListView1(0).ListItems.Clear
'    Else
'        MsgBox "Select PhilHealth Button first, before you can clear the list.", vbInformation, PayrollAlert
'    End If
End Sub

Private Sub GetContributionsFromPayroll()
Dim whereCon As String
Dim pYear As Integer

whereCon = " DATENAME(month, PayrollMaster.StartPaydate) = '" & txtAppMonth.Text & "' AND YEAR(PayrollMaster.StartPaydate) = " & txtAppYear.Text
'whereCon = " format(PayrollMaster.EndPaydate, ""mmmm"") = '" & txtAppMonth.Text & "' AND format(PayrollMaster.EndPaydate, ""yyyy"") = " & txtAppYear.Text
'whereCon = " format(PayrollMaster.EntryDate, ""mmmm"") = '" & txtAppMonth.Text & "' AND format(PayrollMaster.EntryDate, ""yyyy"") = " & txtAppYear.Text

StrQuery = "DELETE FROM tmpPayReport WHERE LTRIM(RTRIM(useriD)) = '" & mUserID & "'"
Call Action_Query(StrQuery, cnnstr)
        ' Updated ECC Employer 8/6/2009
        StrQuery = "INSERT INTO tmpPayReport " & _
                    "(EmployeeID, EmpNo, FirstName, Middlename, Surname, SSS, Philhealth, " & _
                    "Pagibig, TIN, BasicPay, ClientCode, " & _
                    "CashAdvance, WTaxAmount, MiscFee, CalamityLoan, SSSEmployee, SSSEmployer, ECCEmployer, PHealthEmployee, PHealthEmployer, " & _
                    "PagibigEmployee, PagibigEmployer, SSSLoan, PagibigLoan, OtherLoan, StartPayDate, EndPayDate, MonthNo, CutOffNo, ClientName, CutOffDate, EmpPosition, StartContract, MonthOf, UserID) " & _
                    "SELECT PayrollMaster.EmployeeID, PayrollMaster.EmpNo, HREmployees.FirstName, HREmployees.Middlename, HREmployees.Surname, HREmployees.SSS, HREmployees.Philhealth, " & _
                    "HREmployees.Pagibig, HREmployees.TIN, PayrollMaster.BasicPay, PayrollMaster.ClientCode, " & _
                    "PayrollMaster.CashAdvance, PayrollMaster.WTaxAmount, PayrollMaster.MiscFee, PayrollMaster.CalamityLoan, PayrollMaster.SSSEmployee, PayrollMaster.SSSEmployer, PayrollMaster.ECCEmployer, PayrollMaster.PHealthEmployee, PayrollMaster.PHealthEmployer, " & _
                    "PayrollMaster.PagibigEmployee, PayrollMaster.PagibigEmployer, PayrollMaster.SSSLoan, PayrollMaster.PagibigLoan, PayrollMaster.OtherLoan, PayrollMaster.StartPayDate, PayrollMaster.EndPayDate, PayrollMaster.MonthNo, PayrollMaster.CutOffNo, ClientTable.ClientName, ClientTable.CutOffPeriod, HREmployees.EmpPosition, HREmployees.StartContract, DATENAME(month, PayrollMaster.StartPayDate), '" & mUserID & "'" & _
                    " FROM (PayrollMaster INNER JOIN HREmployees ON (PayrollMaster.EmployeeID = HREmployees.EmployeeID)) INNER JOIN ClientTable ON PayrollMaster.ClientCode = ClientTable.ClientCode " & _
                    " WHERE " & whereCon
        Call Action_Query(StrQuery, cnnstr)
        
        ' Get also the data from current payroll transactions.
        whereCon = Replace(whereCon, "payrollmaster", "PayrollTrans")
        whereCon = Replace(whereCon, "PayrollMaster", "PayrollTrans")
        StrQuery = "INSERT INTO tmpPayReport " & _
                    "(EmployeeID, EmpNo, FirstName, Middlename, Surname, SSS, Philhealth, " & _
                    "Pagibig, TIN, BasicPay, ClientCode, " & _
                    "CashAdvance, WTaxAmount, MiscFee, CalamityLoan, SSSEmployee, SSSEmployer, ECCEmployer, PHealthEmployee, PHealthEmployer, " & _
                    "PagibigEmployee, PagibigEmployer, SSSLoan, PagibigLoan, OtherLoan, StartPayDate, EndPayDate, MonthNo, CutOffNo, ClientName, MonthOf, UserID, CutOffDate, EmpPosition, StartContract) " & _
                    "SELECT PayrollTrans.EmployeeID, PayrollTrans.EmpNo, HREmployees.FirstName, HREmployees.Middlename, HREmployees.Surname, HREmployees.SSS, HREmployees.Philhealth, " & _
                    "HREmployees.Pagibig, HREmployees.TIN, PayrollTrans.BasicPay, PayrollTrans.ClientCode, " & _
                    "PayrollTrans.CashAdvance, PayrollTrans.WTaxAmount, PayrollTrans.MiscFee, PayrollTrans.CalamityLoan, PayrollTrans.SSSEmployee, PayrollTrans.SSSEmployer, PayrollTrans.ECCEmployer, PayrollTrans.PHealthEmployee, PayrollTrans.PHealthEmployer, " & _
                    "PayrollTrans.PagibigEmployee, PayrollTrans.PagibigEmployer, PayrollTrans.SSSLoan, PayrollTrans.PagibigLoan, PayrollTrans.OtherLoan, PayrollTrans.StartPayDate, PayrollTrans.EndPayDate, PayrollTrans.MonthNo, PayrollTrans.CutOffNo, ClientTable.ClientName, DATENAME(Month, PayrollTrans.StartPayDate), ' " & mUserID & "', ClientTable.CutOffPeriod, HREmployees.EmpPosition, HREmployees.StartContract " & _
                    " FROM (PayrollTrans INNER JOIN HREmployees ON (PayrollTrans.EmployeeID = HREmployees.EmployeeID)) INNER JOIN ClientTable ON PayrollTrans.ClientCode = ClientTable.ClientCode " & _
                    " WHERE " & whereCon
        Call Action_Query(StrQuery, cnnstr)

        ' Delete all endcutoff date >= 16 with keyword (Backlog).
        StrQuery = "DELETE FROM tmpPayReport WHERE DAY(StartPayDate) >= 16 AND DATENAME(month, StartPayDate) = '" & txtAppMonth.Text & "' and YEAR(StartPayDate) = " & txtAppYear.Text & " AND CutOffDate like '%Backlog%' AND LTRIM(RTRIM(useriD)) = '" & mUserID & "'"
        Call Action_Query(StrQuery, cnnstr)
'        strQuery = "DELETE FROM tmpPayReport WHERE Format(EndPayDate, ""dd"") <= 16 AND Format(EndPayDate, ""mm"") = " & getMonthNum(txtAppMonth.Text, txtAppYear.Text, False, True) & " and Format(EndPayDate, ""yyyy"") = " & getMonthNum(txtAppMonth.Text, txtAppYear.Text, True, True) & " AND CutOffDate like '%Backlog%' AND useriD like '%" & mUserID & "%'"
'        Call Action_Query(strQuery, cnnStr)

'        strQuery = "DELETE FROM tmpPayReport WHERE Format(EndPayDate, ""dd"") >= 16 AND Format(EndPayDate, ""mmmm"") = '" & txtAppMonth.Text & "' and Format(EndPayDate, ""yyyy"") = " & txtAppYear.Text & " AND CutOffDate like '%Backlog%' AND useriD like '%" & mUserID & "%'"
'        Call Action_Query(strQuery, cnnStr)
        ' Get all the corresponding payroll transactions from previous month.
        StrQuery = "INSERT INTO tmpPayReport " & _
                    "(EmployeeID, EmpNo, FirstName, Middlename, Surname, SSS, Philhealth, " & _
                    "Pagibig, TIN, BasicPay, ClientCode, " & _
                    "CashAdvance, WTaxAmount, MiscFee, CalamityLoan, SSSEmployee, SSSEmployer, ECCEmployer, PHealthEmployee, PHealthEmployer, " & _
                    "PagibigEmployee, PagibigEmployer, SSSLoan, PagibigLoan, OtherLoan, StartPayDate, EndPayDate, MonthNo, CutOffNo, ClientName, CutOffDate, EmpPosition, StartContract, MonthOf, UserID) " & _
                    "SELECT PayrollMaster.EmployeeID, PayrollMaster.EmpNo, HREmployees.FirstName, HREmployees.Middlename, HREmployees.Surname, HREmployees.SSS, HREmployees.Philhealth, " & _
                    "HREmployees.Pagibig, HREmployees.TIN, PayrollMaster.BasicPay, PayrollMaster.ClientCode, " & _
                    "PayrollMaster.CashAdvance, PayrollMaster.WTaxAmount, PayrollMaster.MiscFee, PayrollMaster.CalamityLoan, PayrollMaster.SSSEmployee, PayrollMaster.SSSEmployer, PayrollMaster.ECCEmployer, PayrollMaster.PHealthEmployee, PayrollMaster.PHealthEmployer, " & _
                    "PayrollMaster.PagibigEmployee, PayrollMaster.PagibigEmployer, PayrollMaster.SSSLoan, PayrollMaster.PagibigLoan, PayrollMaster.OtherLoan, PayrollMaster.StartPayDate, PayrollMaster.EndPayDate, PayrollMaster.MonthNo, PayrollMaster.CutOffNo, ClientTable.ClientName, ClientTable.CutOffPeriod, HREmployees.EmpPosition, HREmployees.StartContract, '" & txtAppMonth.Text & "', '" & mUserID & "'" & _
                    " FROM (PayrollMaster INNER JOIN HREmployees ON (PayrollMaster.EmployeeID = HREmployees.EmployeeID)) INNER JOIN ClientTable ON PayrollMaster.ClientCode = ClientTable.ClientCode " & _
                    " WHERE DAY(PayrollMaster.StartPayDate) >= 16 AND MONTH(PayrollMaster.StartPayDate) = " & getMonthNum(txtAppMonth.Text, txtAppYear.Text) & " AND YEAR(PayrollMaster.StartPayDate) = " & getMonthNum(txtAppMonth.Text, txtAppYear.Text, False) & _
                    " AND ClientTable.CutOffPeriod like '%Backlog%'"
        Call Action_Query(StrQuery, cnnstr)
        
        StrQuery = "SELECT tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, min(tmpPayReport.SSS) as SSS, " & _
                "MIN( tmpPayReport.BirthDate) as BirthDate, MIN(tmpPayReport.Philhealth) as PhilHealth, min(tmpPayReport.Pagibig) as pagibig, min(tmpPayReport.TIN) as TIN, tmpPayReport.ClientCode, tmppayreport.monthof, " & _
                "SUM(tmpPayReport.BasicPay) as BasicPay, SUM(tmpPayReport.Month13Payment) as month13payment, sum(tmpPayReport.TASAesla) as TASAesla, sum(tmpPayReport.LMCDed) as LMCDed, " & _
                "SUM(tmpPayReport.ValewithClient) as valewithclient, sum(tmpPayReport.CashAdvance) as CashAdvance, sum(tmpPayReport.WTaxAmount) as WTaxAmount, " & _
                "SUM(tmpPayReport.Medical) as Medical, sum(tmpPayReport.SalaryLoan) as SalaryLoan, sum(tmpPayReport.Uniform) as Uniform, sum(tmpPayReport.Variances) as Variances, sum(tmpPayReport.Vale_IOU) as Vale_IOU, " & _
                "SUM(tmpPayReport.MiscFee) as miscfee, sum(tmpPayReport.CalamityLoan) as CalamityLoan, sum(tmpPayReport.SSSEmployee) as SSSEmployee, " & _
                "SUM(tmpPayReport.SSSEmployer) as sssemployer, MAX(tmpPayReport.ECCEmployer) as ECCemployer, SUM(tmppayreport.sssloan) as sssloan, SUM(tmppayreport.pagibigloan) as pagibigloan, sum(tmpPayReport.PHealthEmployee) as phealthemployee, sum(tmpPayReport.PHealthEmployer) as PHealthEmployer, " & _
                "SUM(tmpPayReport.PagibigEmployee) as pagibigemployee, sum(tmpPayReport.PagibigEmployer) as pagibigemployer, " & _
                "MIN(tmpPayReport.CutOffNo) as cutoffno, min(tmpPayReport.ClientName) as clientname, min(tmpPayReport.CutOffDate) as CutOffDate, min(tmpPayreport.EmpPosition) AS EmpPosition, min(tmpPayReport.StartContract) AS StartContract FROM tmpPayReport " & _
                "WHERE LTRIM(RTRIM(userID)) = '" & mUserID & "' " & _
                "GROUP BY tmpPayReport.EmployeeID, tmpPayReport.EmpNo, tmpPayReport.FirstName, tmpPayReport.Middlename, tmpPayReport.Surname, tmpPayReport.ClientCode, tmppayreport.monthof " & _
                "ORDER BY tmppayreport.ClientCode, tmppayreport.Surname, tmppayreport.FirstName"
'    Set rsDeductions = ReturnRst(strQuery, cnnStr, mstatic)
Set rsEmployeeBasicData = ReturnRst(StrQuery, cnnstr, mstatic)
                
End Sub

Private Sub Form_Load()
Dim mComp As String
Dim mEmpN As String

    If ActiveSw = False Then
        Me.Top = (mdiMain.Height - Me.Height) / 8
        Me.Left = (mdiMain.Width - Me.Width) / 2
    End If
    FormLoadSw = False
    mAdjustment = 1
    WithAdjustmentFlag = 2
    cmdAddlist.Enabled = False
    AddToListFlag = True
    
    If PHealthFlag = True Then
        'frmSSSReport.Visible = False
        frmPHealthReport.Visible = True
        lblclient(7).Caption = "No PHealth No."
        cmdNoPHNo.Caption = "Check No PHealth No."
        cmdAddlist.Caption = "Add to PhilHealth List"
        Me.Caption = "Client List for PhilHealth Report for the month of " & frmPHealthReport.cmdAppMonth.Text & " " & frmPHealthReport.cmdAppYear.Text
        lblclient(38).Caption = "Applicable Month: " & frmPHealthReport.cmdAppMonth.Text & " " & frmPHealthReport.cmdAppYear.Text
    Else
        'frmPHealthReport.Visible = False
        frmSSSReport.Visible = True
        lblclient(7).Caption = "No SSS No."
        cmdNoPHNo.Caption = "Check No SSS No."
        cmdAddlist.Caption = "Add to SSS List"
        Me.Caption = "Client List for SSS Report for the month of " & frmSSSReport.cmdAppMonth.Text & " " & frmSSSReport.cmdAppYear.Text
        lblclient(38).Caption = "Applicable Month: " & frmSSSReport.cmdAppMonth.Text & " " & frmSSSReport.cmdAppYear.Text
    End If
    ActiveSw = True
    RefreshSw = False
    If PHealthFlag = True Then
        Me.Caption = "Select Clients for PhilHealth Report"
        txtAppMonth.Text = frmPHealthReport.cmdAppMonth.Text
        txtAppYear.Text = frmPHealthReport.cmdAppYear.Text
    Else
        Me.Caption = "Select Clients for SSS Report"
        txtAppMonth.Text = frmSSSReport.cmdAppMonth.Text
        txtAppYear.Text = frmSSSReport.cmdAppYear.Text
    End If
    'Call GetClientData ' call it in Form Activate
    txtRemarks.Text = "Total Employees is 0"
End Sub

Private Sub GetClientData()

If PHealthFlag = True Then  ' For PhilHealth
    StrQuery = "SELECT a.Clientcode, a.ClientName, b.Remarks, b.ApplicableMonth, b.ApplicableYear, b.ProcessedDate, b.ProcessedBy, a.ClientFlag3 AS EcolaFlag, b.GroupCode, b.SBRNo " & _
        "FROM clienttable a LEFT JOIN (SELECT DISTINCT AgencyCode, ClientCode, ClientName, ContriCode, Remarks, ApplicableMonth, ApplicableYear, ProcessedDate, ProcessedBy, GroupCode, SBRNo, AmountPaid, DatePaid, DatePosted " & _
        "FROM PHContributions WHERE ApplicableMonth = '" & txtAppMonth.Text & "' AND ApplicableYear = " & txtAppYear.Text & ") AS b " & _
        "ON a.ClientName = b.ClientName " & _
        "WHERE a.clientstatus = 1 and a.clientname is not null ORDER BY a.clientname"
Else                        ' For SSS
    StrQuery = "SELECT a.Clientcode, a.ClientName, b.Remarks, b.ApplicableMonth, b.ApplicableYear, b.ProcessedDate, b.ProcessedBy, a.ClientFlag3 AS EcolaFlag, b.GroupCode, b.SBRNo " & _
        "FROM clienttable a LEFT JOIN (SELECT DISTINCT AgencyCode, ClientCode, ClientName, ContriCode, Remarks, ApplicableMonth, ApplicableYear, ProcessedDate, ProcessedBy, GroupCode, SBRNo, AmountPaid, DatePaid, DatePosted " & _
        "FROM SSSContributions WHERE ApplicableMonth = '" & txtAppMonth.Text & "' AND ApplicableYear = " & txtAppYear.Text & ") AS b " & _
        "ON a.ClientName = b.ClientName " & _
        "WHERE a.clientstatus = 1 and a.clientname is not null ORDER BY a.clientname"
End If

Set rsClientData = ReturnRst(StrQuery, cnnstr, mstatic)
If rsClientData.RecordCount <> 0 Then
    Call FillClientList2(ListView1(0), rsClientData)
    rsClientData.MoveFirst
End If
Call GetContributionsFromPayroll

End Sub

Public Sub FillClientList2(ByRef mLsEmp As ListView, _
                        mRs As Recordset)
Dim litem1 As ListItem

    mLsEmp.View = lvwReport
    mLsEmp.ListItems.Clear
    While Not (mRs.EOF Or mRs.BOF)
'        Set litem1 = mLsEmp.ListItems.Add(, , mRs!ClientCode)
        Set litem1 = mLsEmp.ListItems.Add(, , IIf(IsNull(mRs!ClientName), "Blank", mRs!ClientName))
        'litem1.SubItems(1) = IIf(IsNull(mRs!ClientName), "Blank", mRs!ClientName)
        litem1.SubItems(1) = IIf(IsNull(mRs!Remarks), "", mRs!Remarks)
        litem1.SubItems(2) = IIf(IsNull(mRs!ProcessedBy), "", mRs!ProcessedBy)
        litem1.SubItems(3) = IIf(IsNull(mRs!ProcessedDate) Or mRs!ProcessedDate = #1/1/1900#, "__/__/____", Format(mRs!ProcessedDate, "mm/dd/yyyy"))
        litem1.SubItems(4) = IIf(IsNull(mRs!groupcode), "", mRs!groupcode)
        litem1.SubItems(5) = IIf(IsNull(mRs!SBRNo), "", mRs!SBRNo)
        mRs.MoveNext
    Wend
End Sub

Private Sub DisplayClientData(ByVal SelectedClient As String, Optional recSw As Boolean, Optional mIndex As Integer)

If mIndex = 0 Then
With rsEmployeeBasicData
    If recSw = False Then
        '.MoveFirst
        '.Find "clientcode = " & SelectedClient
        '.Find "clientname = '" & SelectedClient & "'"
        .Filter = "clientname = '" & Replace(SelectedClient, "'", "''") & "'"
    End If
    If PHealthFlag = True Then
        ListView1(1).Visible = True
        ListView1(2).Visible = False
        ListView1(1).ListItems.Clear
    Else
        ListView1(2).Visible = True
        ListView1(1).Visible = False
        ListView1(2).ListItems.Clear
    End If
    
    If Not .EOF And Not .BOF And mIndex = 0 Then
        If PHealthFlag = True Then
            Call FillEmployeeList(SelectedClient, ListView1(1), rsEmployeeBasicData, False)
        Else
            Call FillEmployeeList(SelectedClient, ListView1(2), rsEmployeeBasicData, False)
        End If
    Else
        MsgBox "No Employee records selected for " & SelectedClient & "!", vbInformation, PayrollAlert
        .Filter = ""
    End If
End With
End If

End Sub

'Private Sub FilterEmployee(ByVal mCompany As String, Optional NotYet As Boolean = False)
'    rsClientData.MoveFirst
'    rsClientData.Find "clientname = '" & Trim(mCompany) & "'"
'    If rsClientData.EOF = True Then
'        mClientCode = 0
'        Exit Sub
'    Else
'        mClientCode = rsClientData!ClientCode
'        'txtClientCode = rsClientData!ClientCode
'    End If
'    If FormLoadSw = True Then
'    With rsEmployeeBasicData
'        .Filter = "clientcode = " & mClientCode
'        If .RecordCount <> 0 Then
'            'If Not .EOF And Not .BOF Then
'            Call FillComboData(cmbEmployees, rsEmployeeBasicData)
'            cmbEmployees.ListIndex = 0
'            mEmpNo = cmbEmployees.ItemData(cmbEmployees.ListIndex)
'        Else
'            Call clearList
'        End If
'    End With
'    End If
'End Sub

Private Sub FillEmployeeList3(ByRef mLsEmp As ListView, _
                        ByRef mRs As Recordset, Optional AppendSw As Boolean = True)
Dim litem1 As ListItem

    InActiveCtr = 0
    ActiveCtr = 0
    'mLsEmp.View = lvwReport
    ListView1(1).View = lvwReport
    ListView1(1).ListItems.Clear
    'mRs.MoveFirst
    While Not (mRs.EOF Or mRs.BOF)
        Set litem1 = ListView1(1).ListItems.Add(, , mRs!ClientCode)
        litem1.SubItems(1) = IIf(IsNull(mRs!AgencyName), "Blank", mRs!AgencyName)
        litem1.SubItems(2) = IIf(IsNull(mRs!ClientName), "Blank", mRs!ClientName)
        litem1.SubItems(3) = IIf(IsNull(mRs!EmpNo), "Blank", mRs!EmpNo)
        litem1.SubItems(4) = IIf(IsNull(mRs!StartContract) Or mRs!StartContract = #1/1/1900#, "__/__/____", Format(mRs!StartContract, "mm/dd/yyyy"))
        litem1.SubItems(5) = IIf(IsNull(mRs!endcontract) Or mRs!endcontract = #1/1/1900#, "__/__/____", Format(mRs!endcontract, "mm/dd/yyyy"))
        litem1.SubItems(6) = Format(mRs!Cbamount, "0.##")
        litem1.SubItems(7) = Format(mRs!CBPayment, "0.##")
        litem1.SubItems(8) = IIf(IsNull(mRs!CBStartPayment) Or mRs!CBStartPayment = #1/1/1900#, "__/__/____", Format(mRs!CBStartPayment, "mm/dd/yyyy"))
        litem1.SubItems(9) = IIf(IsNull(mRs!CBEndPayment) Or mRs!CBEndPayment = #1/1/1900#, "__/__/____", Format(mRs!CBEndPayment, "mm/dd/yyyy"))
        mRs.MoveNext
    Wend
End Sub

Private Function FillComboData(ByRef mCmb As ComboBox, _
                        mRs As Recordset, Optional FilterFlag As Boolean = False, _
                        Optional ComboList As String = "")
    mCmb.Clear
    While Not (mRs.EOF Or mRs.BOF)
        If FilterFlag = False Then
            mCmb.AddItem Trim(Replace(IIf(IsNull(mRs.Fields(1)), "None", mRs.Fields(1)), "'", "''"))
            mCmb.ItemData(mCmb.NewIndex) = mRs.Fields(0)
        Else    ' need to filter the lists
            If InStr(1, ComboList, LTrim(RTrim(str(mRs.Fields(0)))), 1) = 0 Then
                mCmb.AddItem Trim(Replace(IIf(IsNull(mRs.Fields(1)), "None", mRs.Fields(1)), "'", "''"))
                mCmb.ItemData(mCmb.NewIndex) = mRs.Fields(0)
            End If
        End If
        mRs.MoveNext
    Wend
End Function

'Private Sub ListView1_Click(Index As Integer)
'Dim res
'
''    If ActiveCtr = 0 Then
''        Call DisplayBasicData("", False)
''    Else
''        If SSTab1(1).Tab = 1 Then ' in-active tab
''            If ListView1(1).ListItems.Count = 0 Then
''                Call DisplayBasicData("", False)
''            Else
''                Call DisplayBasicData(ListView1(1).SelectedItem.Text, False)
''                CurrEmployee = ListView1(1).SelectedItem.Text
''            End If
''        Else
''            If ListView1(0).ListItems.Count = 0 Then
''                Call DisplayBasicData("", False)
''            Else
''                Call DisplayBasicData(ListView1(0).SelectedItem.Text, False)
''                CurrEmployee = ListView1(0).SelectedItem.Text
''                If AddDeductions = True Or EditDeductions = True Then
''                    'res = MsgBox("Cancel changes to deductions?", vbQuestion + vbYesNo, PayrollAlert)
''                    'If res = vbYes Then ' cancel the changes
''                        Call EnableDisableDeductions(False)
''                        Call EnableDisableDeductionButtons("Add", "Edit", True)
''                        EditDeductions = False
''                        AddDeductions = False
''                        cmbTypes.Enabled = True
''                        MsgBox "Changes to deductions was cancelled!", vbInformation, PayrollAlert
''                    'End If
''                End If
''            End If
''        End If
''    End If
'End Sub

'Private Sub ListView1_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    ListView1(Index).SortKey = ColumnHeader.Index - 1
'    ListView1(Index).Sorted = True
'End Sub
'
'Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
'    CurrEmployee = Item.Text
'End Sub

Private Sub clearList()
    ListView1(1).ListItems.Clear
End Sub

'Private Sub GenerateReport()
'Dim rsDedRecords As ADODB.Recordset
'Dim xlApp1 As Excel.Application
'Dim xlBook1 As Excel.Workbook
'Dim xlSheet As Excel.Worksheet
'Dim mTotalPay As Double
'Dim mTotalDeduction As Double
'Dim mRange As String, mRange2 As String
'Dim sumString As String
'Dim ctr As Long
'Dim i As Integer
'Dim cPos As Integer
'Dim LineCtr As Integer
'Dim mHead As String
'Dim mDeduc As Double
'Dim StartDeduction As Variant
'Dim EndDeduction As Variant
'Dim DeductionYear As Integer
'
'' Get all deduction records for specific client.
''strQuery = "SELECT * FROM qryDeductionRecords WHERE ClientCode = " & mClientCode & " order by clientCode"
''Set rsDedRecords = ReturnRst(strQuery, cnnStr, mstatic)
''If rsDedRecords.EOF = True Then
''    MsgBox "No payroll deduction records selected! Cannot proceed.", vbInformation
''    Set rsDedRecords = Nothing
''    GoTo ExitRtn2
''End If
'
'Set xlApp1 = New Excel.Application
'Set xlBook1 = xlApp1.Workbooks.Add()
'
'xlApp1.WindowState = xlMaximized
'xlApp1.Visible = True
'xlApp1.DisplayAlerts = False
'
'Set xlSheet = xlBook1.Worksheets(1)
'
'ctr = 8
'i = 1
'
''Updated by Bernard May 13, 2008.
'mHead = CompanyName & " (" & CompanyShort & ")"
'Call PrintReportHeader(xlSheet, xlApp1, "A2", mHead, "", xlLeft, , False, False, False, , , 10, True)
'mHead = "Cash Bond Summary Report"
'Call PrintReportHeader(xlSheet, xlApp1, "A3", mHead, "", xlLeft, , False, False, False, , , 10, True)
'mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy")
'Call PrintReportHeader(xlSheet, xlApp1, "A4", mHead, "", xlLeft, , False, False, False, , , 10, True)
'mHead = "COMPANY NAME: " & cmbCompany.Text
'Call PrintReportHeader(xlSheet, xlApp1, "A5", mHead, "", xlLeft, , False, False, False, , , 10, True)
'
'Call PrintReportHeader(xlSheet, xlApp1, "A6:A7", "Agency Name", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 20)
'Call PrintReportHeader(xlSheet, xlApp1, "B6:B7", "Client Name", "", xlCenter, xlBottom, True, False, True, 15, , 9, False, 25)
'Call PrintReportHeader(xlSheet, xlApp1, "C6:C7", "Control No.", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 11)
'Call PrintReportHeader(xlSheet, xlApp1, "D6:D7", "Employee Name", "", xlCenter, xlBottom, True, True, True, 15, , 9, False, 23)
'Call PrintReportHeader(xlSheet, xlApp1, "E6:E7", "Start Contract", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 6.8)
'Call PrintReportHeader(xlSheet, xlApp1, "F6:F7", "End Contract", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 6.8)
'Call PrintReportHeader(xlSheet, xlApp1, "G6:G7", "Cash Bond Total", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.6)
'Call PrintReportHeader(xlSheet, xlApp1, "H6:H7", "Cash Bond Payment", "", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.6)
'Call PrintReportHeader(xlSheet, xlApp1, "I6:I7", "Start", "Cut-Off", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.5)
'Call PrintReportHeader(xlSheet, xlApp1, "J6:J7", "End", "Cut-Off", xlCenter, xlCenter, True, False, True, 15, , 8, False, 7.5)
'
'For i = 1 To ListView1(1).Index
'    If ListView1(1).ListItems.Item(i).Checked = True And Val(ListView1(1).ListItems.Item(i).SubItems(6)) Then
'        StartDeduction = Format(ListView1(1).ListItems.Item(i).SubItems(4), "mm/dd/yyyy") ' Start of Contract
'        EndDeduction = Format(ListView1(1).ListItems.Item(i).SubItems(5), "mm/dd/yyyy") ' End of Contract
'        'DeductionYear = Format(ListView1(1).ListItems.Item(i).SubItems(8), "yyyy")
'        strQuery = "SELECT EmployeeID, TASAesla, StartPaydate, EndPayDate FROM qryContractHistory WHERE EmployeeID = " & _
'            mEmpNo & " AND StartPaydate >= '" & StartDeduction & _
'            "' AND EndPayDate <= '" & EndDeduction & "' ORDER BY StartPayDate"
'        Set rsDedRecords = ReturnRst(strQuery, cnnStr, mstatic)
'        If rsDedRecords.EOF = False Then
'            xlSheet.Cells(ctr, 1).VALUE = ListView1(1).ListItems.Item(i).SubItems(1) ' Agency Name
'            xlSheet.Cells(ctr, 2).VALUE = ListView1(1).ListItems.Item(i).SubItems(2) ' Client Name
'            xlSheet.Cells(ctr, 3).VALUE = ListView1(1).ListItems.Item(i).SubItems(3) ' Control Number
'            xlSheet.Cells(ctr, 4).VALUE = cmbEmployees.Text  ' Employee Name
'            xlSheet.Cells(ctr, 5).VALUE = ListView1(1).ListItems.Item(i).SubItems(4) ' Start Contract
'            xlSheet.Cells(ctr, 6).VALUE = ListView1(1).ListItems.Item(i).SubItems(5) ' End Contract
'            xlSheet.Cells(ctr, 7).VALUE = Format(ListView1(1).ListItems.Item(i).SubItems(6), "#,##0.00") ' Total Cash Bond
'            Do While Not rsDedRecords.EOF
'                xlSheet.Cells(ctr, 8).VALUE = Format(rsDedRecords!TASAesla, "#,##0.00") ' cash bond Payment
'                xlSheet.Cells(ctr, 9).VALUE = Format(rsDedRecords!StartPayDate, "mm/dd/yyyy") ' start cutoff
'                xlSheet.Cells(ctr, 10).VALUE = Format(rsDedRecords!EndPayDate, "mm/dd/yyyy") ' end cutoff
'                ctr = ctr + 1
'                rsDedRecords.MoveNext
'            Loop
'            ' Print Sub-Totals.
'            xlSheet.Cells(ctr, 4).VALUE = "     Total Payments"
'            mRange = "H" & 8 & ":H" & ctr
'            xlSheet.Range(mRange).Select
'            mRange = "H" & ctr
'            xlSheet.Range(mRange).Activate
'            xlApp1.ActiveCell.FormulaR1C1 = "=SUM(R[" & (ctr - 7) * -1 & "]C:R[-2]C)"
'            xlApp1.Range(mRange).Font.Bold = True
'            mRange = "G8:H" & ctr
'            xlSheet.Range(mRange).Select
'            xlSheet.Range(mRange).NumberFormat = "#,##0.00_);(#,##0.00)"
'        Else
'            xlSheet.Cells(ctr, 1).VALUE = ListView1(1).ListItems.Item(i).SubItems(1) ' Agency Name
'            xlSheet.Cells(ctr, 2).VALUE = ListView1(1).ListItems.Item(i).SubItems(2) ' Client Name
'            xlSheet.Cells(ctr, 3).VALUE = ListView1(1).ListItems.Item(i).SubItems(3) ' Control Number
'            xlSheet.Cells(ctr, 4).VALUE = cmbEmployees.Text  ' Employee Name
'            xlSheet.Cells(ctr, 5).VALUE = ListView1(1).ListItems.Item(i).SubItems(4) ' Start Contract
'            xlSheet.Cells(ctr, 6).VALUE = ListView1(1).ListItems.Item(i).SubItems(5) ' End Contract
'            xlSheet.Cells(ctr, 7).VALUE = Format(ListView1(1).ListItems.Item(i).SubItems(6), "#,##0.00") ' Total Cash Bond
'            ctr = ctr + 1
'            xlSheet.Cells(ctr, 1).VALUE = "No payments found."
'        End If
'    Else
'        xlSheet.Cells(ctr, 1).VALUE = ListView1(1).ListItems.Item(i).SubItems(1) ' Agency Name
'        xlSheet.Cells(ctr, 2).VALUE = ListView1(1).ListItems.Item(i).SubItems(2) ' Client Name
'        xlSheet.Cells(ctr, 3).VALUE = ListView1(1).ListItems.Item(i).SubItems(3) ' Control Number
'        xlSheet.Cells(ctr, 4).VALUE = cmbEmployees.Text  ' Employee Name
'        xlSheet.Cells(ctr, 5).VALUE = ListView1(1).ListItems.Item(i).SubItems(4) ' Start Contract
'        xlSheet.Cells(ctr, 6).VALUE = ListView1(1).ListItems.Item(i).SubItems(5) ' End Contract
'        xlSheet.Cells(ctr, 7).VALUE = Format(ListView1(1).ListItems.Item(i).SubItems(6), "#,##0.00") ' Total Cash Bond
'        ctr = ctr + 1
'        xlSheet.Cells(ctr, 1).VALUE = "Zero (0) deductions."
'    End If
'Next i
'ctr = ctr + 2
'
''Box it.
'mRange = "A6:J" & ctr
'Call PlaceBox(xlSheet, xlApp1, mRange)
'
'ctr = ctr + 1
'mHead = "Printed Date: " & Format(Now, "mm/dd/yyyy") & "  Time: " & Format(Now, "HH:MM:SSAMPM")
'mRange = "A" & ctr
'Call CustomizedHeader(xlSheet, xlApp1, mRange, mHead, "", xlLeft, , False, False, False, , , 8, False, 25)
'
'ctr = ctr + 3
'xlSheet.Cells(ctr, 1).VALUE = "Prepared by: ______________"
'xlSheet.Cells(ctr, 3).VALUE = "Note by: ______________"
'xlSheet.Cells(ctr, 7).VALUE = "Received by: ______________"
'xlSheet.Range("A1").Select
'
'
'
'
'Do While Not rs.EOF
'    ' DoEvents
'    xlSheet.Cells(ctr, 1).VALUE = i
'    xlSheet.Cells(ctr, 2).VALUE = rs!EmpName
'    mDeduc = 0#
''    If rs!CbAmount <> 0 Then
''        mDeduc = RunningBalance(rsDedRecords, 1, rs!EmployeeID, rs!cbstartpayment, rs!cbendpayment)
''        xlSheet.Cells(ctr, 8).VALUE = mDeduc
''    End If
'    xlSheet.Cells(ctr, 3).VALUE = Format(rs!DailyRate, "0.##")
'    xlSheet.Cells(ctr, 20).VALUE = Format(rs!StartContract, "mm/dd/yyyy")
'    xlSheet.Cells(ctr, 21).VALUE = Format(rs!endcontract, "mm/dd/yyyy")
'    xlSheet.Cells(ctr, 22).VALUE = rs!AccountNo
'    xlSheet.Cells(ctr, 23).VALUE = rs!Remarks
'    ctr = ctr + 1
'    i = i + 1
'    rs.MoveNext
'Loop
'mRange = "B" & ctr + 1 & ":B" & ctr + 1
'Call PrintReportHeader(xlSheet, xlApp1, mRange, "sub-total", "", xlCenter, xlBottom, True, False, True, xlNone, , 9, False, 22)
'mRange = "B" & ctr + 2 & ":B" & ctr + 2
'Call PrintReportHeader(xlSheet, xlApp1, mRange, "Total of Staff", "", xlCenter, xlBottom, True, False, True, xlNone, , 9, False, 22)
'mRange = "D" & ctr + 2 & ":D" & ctr + 2
'Call PrintReportHeader(xlSheet, xlApp1, mRange, i - 1, "", xlCenter, xlBottom, True, False, True, xlNone, , 9, False, 5)
'mRange = "B" & ctr + 3 & ":B" & ctr + 3
'Call PrintReportHeader(xlSheet, xlApp1, mRange, "No. of lacking", "", xlCenter, xlBottom, True, False, True, xlNone, , 9, False, 22)

'ExitRtn2:
'
'End Sub

Private Sub optEmployerType_Click(Index As Integer)

txtCustomized.Refresh
Select Case Index
    Case 0
        cmdComputeLess1.Caption = "Recompute Based from Table"
        WithAdjustmentFlag = 1
        txtCustomized.Visible = False
        mAdjustment = 0
    Case 1
        cmdComputeLess1.Caption = "Recompute with 1 Adjustment"
        mAdjustment = 1
        WithAdjustmentFlag = 2
        txtCustomized.Visible = False
    Case 2
        cmdComputeLess1.Caption = "Recompute with " & txtCustomized.Text & " Adjustment"
        mAdjustment = Val(txtCustomized.Text)
        txtCustomized.Visible = True
        WithAdjustmentFlag = 3
End Select
cmdComputeLess1.Refresh

End Sub
