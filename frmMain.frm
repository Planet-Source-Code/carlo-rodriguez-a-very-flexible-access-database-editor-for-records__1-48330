VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3810
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   3810
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Select Operation"
      Height          =   1635
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   3795
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add Record"
         Enabled         =   0   'False
         Height          =   525
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   300
         Width           =   1245
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify Record"
         Enabled         =   0   'False
         Height          =   525
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   300
         Width           =   1245
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Record"
         Enabled         =   0   'False
         Height          =   525
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   930
         Width           =   1245
      End
      Begin VB.CommandButton cmdCompact 
         Caption         =   "&Compact/Repair Database"
         Enabled         =   0   'False
         Height          =   525
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   930
         Width           =   1245
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2130
      Top             =   -120
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   1950
      TabIndex        =   4
      Top             =   4140
      Width           =   1845
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "HH:MM:SS AMPM"
         Height          =   210
         Left            =   270
         TabIndex        =   5
         Top             =   270
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   0
      TabIndex        =   3
      Top             =   4140
      Width           =   1935
      Begin VB.Label lblDAte 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Month Day, Year"
         Height          =   210
         Left            =   360
         TabIndex        =   6
         Top             =   270
         Width           =   1245
      End
   End
   Begin MSComDlg.CommonDialog cdbox 
      Left            =   1620
      Top             =   -180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Open Database"
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   1650
      Width           =   3795
      Begin VB.CommandButton cmdOpen 
         Appearance      =   0  'Flat
         Caption         =   "&Open Database"
         Height          =   555
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   210
         Width           =   855
      End
      Begin VB.Label lblDBFilename 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAdd_Click()
frmMain.Hide
frmAdd.Show
End Sub

Private Sub cmdCompact_Click()
On Error GoTo errH
DBEngine.CompactDatabase (DbFileName), (App.Path & "\tempdb.mdb"), dbLangGeneral
Kill (DbFileName)
DBEngine.CompactDatabase (App.Path & "\tempdb.mdb"), (DbFileName), dbLangGeneral
Kill (App.Path & "\tempdb.mdb")
Exit Sub

errH:
MsgBox err.Description, vbCritical, "Compact Database."
err.Clear
End Sub

Private Sub cmdDelete_Click()
frmMain.Hide
frmDelete.Show
End Sub
Private Sub cmdModify_Click()
frmMain.Hide
frmModify.Show
End Sub
Private Sub cmdOpen_Click()
On Error GoTo errHandler:
cdbox.CancelError = True
cdbox.Filter = "Access Database (*.mdb)|*.mdb"
cdbox.ShowOpen
DbFileName = cdbox.FileName
lblDBFilename.Caption = DbFileName
DbOpen = True
frmPassword.Show
Enable_Buttons
errHandler:
Exit Sub
End Sub
Private Sub Form_Initialize()
DbOpen = False
lblDBFilename.Caption = DbFileName
End Sub
Private Sub Form_Load()
lblDAte.Caption = Format(Now, "Mmmm dd, yyyy")
lblTime.Caption = Format(Now, "hh:mm:ss AMPM")
End Sub
Private Sub Timer1_Timer()
lblTime.Caption = Format(Now, "hh:mm:ss AMPM")
End Sub
Private Sub Enable_Buttons()
If DbOpen = True Then
    cmdAdd.Enabled = True
    cmdModify.Enabled = True
    cmdDelete.Enabled = True
    cmdCompact.Enabled = True
End If
End Sub
