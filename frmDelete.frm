VERSION 5.00
Begin VB.Form frmDelete 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Record"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   150
      TabIndex        =   16
      Text            =   "Select Table Name"
      Top             =   270
      Width           =   2115
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3885
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   2415
      Begin VB.Frame Frame2 
         Caption         =   "Table Field(s)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3045
         Left            =   0
         TabIndex        =   14
         Top             =   810
         Width           =   2355
         Begin VB.ListBox lstField 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2580
            ItemData        =   "frmDelete.frx":0000
            Left            =   150
            List            =   "frmDelete.frx":0002
            TabIndex        =   15
            Top             =   330
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Table Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   0
         TabIndex        =   13
         Top             =   -30
         Width           =   2355
         Begin VB.TextBox txtTableName 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   180
            TabIndex        =   0
            ToolTipText     =   "Enter TableName and Press Enter"
            Top             =   300
            Width           =   2055
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Delete Records"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   2430
      TabIndex        =   2
      Top             =   0
      Width           =   3375
      Begin VB.TextBox txtData 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1140
         Width           =   3015
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1950
         TabIndex        =   7
         Top             =   2820
         Width           =   645
      End
      Begin VB.CommandButton cmdPrevRec 
         Height          =   465
         Left            =   1950
         Picture         =   "frmDelete.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Previous Record"
         Top             =   2340
         Width           =   645
      End
      Begin VB.CommandButton cmdPrev 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1260
         Picture         =   "frmDelete.frx":0446
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Previous Field"
         Top             =   2820
         Width           =   645
      End
      Begin VB.CommandButton cmdNext 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2640
         Picture         =   "frmDelete.frx":0888
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Next Field"
         Top             =   2820
         Width           =   645
      End
      Begin VB.CommandButton cmdNextRec 
         Height          =   465
         Left            =   1950
         Picture         =   "frmDelete.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Next Record"
         Top             =   3300
         Width           =   645
      End
      Begin VB.Label lblDataBase 
         Caption         =   "DatbaseFileName"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   240
         TabIndex        =   11
         Top             =   270
         Width           =   2595
      End
      Begin VB.Label lblFieldName 
         AutoSize        =   -1  'True
         Caption         =   "Current Data for Field:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   10
         Top             =   810
         Width           =   1575
      End
      Begin VB.Label lblDPosition 
         AutoSize        =   -1  'True
         Caption         =   "Data Postion"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   270
         TabIndex        =   9
         Top             =   2250
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4890
      TabIndex        =   1
      Top             =   4020
      Width           =   795
   End
End
Attribute VB_Name = "frmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAction_Click()
DelRec
ShowInfo
End Sub

Private Sub cmdBack_Click()
Unload Me
frmMain.Show
End Sub
Private Sub cmdNext_Click()
NextFld
ShowInfo
End Sub
Private Sub cmdNextRec_Click()
NextRecord
ShowInfo
End Sub
Private Sub cmdPrev_Click()
PrevFld
ShowInfo
End Sub
Private Sub cmdPrevRec_Click()
PrevRecord
ShowInfo
End Sub

Private Sub Combo1_Click()
OpenTable
End Sub

Private Sub Form_Load()
On Error GoTo ErrH
OpenDbase
Combo1.Enabled = True
lblDataBase.Caption = DbFileName
Exit Sub

ErrH:
If err.Number = 3031 Then
    MsgBox "The password you entered is incorrect.", vbCritical, "Open Database"
Else
    MsgBox err.Description, vbCritical, "Open Database."
End If
cmdPrevRec.Enabled = False
cmdNextRec.Enabled = False
cmdAction.Enabled = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
Combo1.Enabled = False
err.Clear
frmMain.Show

End Sub

Private Sub txtTableName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then OpenTable
End Sub
Private Sub OpenTable()
Dim i As Integer
rsName = Combo1.Text
RsOpen
lstField.Clear
For i = 0 To FldNum
    lstField.AddItem (rs.Fields(i).Name)
Next i
GetRecordData
GetFieldName
ShowInfo
End Sub
Private Sub ShowInfo()
txtData.Text = FldData
lblFieldName.Caption = CurrFldName
lblDPosition.Caption = CStr(RecrdPos) + " of " + CStr(RecNum)
End Sub

Private Sub Combo1_GotFocus()
Combo1.Clear
    For i = 0 To Db.TableDefs.Count - 1
        If UCase(Left(Db.TableDefs(i).Name, 4)) <> "MSYS" Then
            Combo1.AddItem (Db.TableDefs(i).Name)
        End If
    Next i
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    OpenTable
End If
End Sub

