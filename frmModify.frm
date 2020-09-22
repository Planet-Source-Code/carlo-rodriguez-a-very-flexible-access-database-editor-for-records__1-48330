VERSION 5.00
Begin VB.Form frmModify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Record"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5820
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   150
      TabIndex        =   16
      Text            =   "Select Table Name"
      Top             =   270
      Width           =   2115
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
      TabIndex        =   14
      Top             =   4020
      Width           =   795
   End
   Begin VB.Frame Frame4 
      Caption         =   "Edit Records"
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
      TabIndex        =   6
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton cmdNextRec 
         Height          =   465
         Left            =   1950
         Picture         =   "frmModify.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Next Record"
         Top             =   3300
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
         Picture         =   "frmModify.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Next Field"
         Top             =   2820
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
         Picture         =   "frmModify.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Previous Field"
         Top             =   2820
         Width           =   645
      End
      Begin VB.CommandButton cmdPrevRec 
         Height          =   465
         Left            =   1950
         Picture         =   "frmModify.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Previous Record"
         Top             =   2340
         Width           =   645
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Edit"
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
         TabIndex        =   2
         Top             =   1140
         Width           =   3015
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
         TabIndex        =   15
         Top             =   2250
         Width           =   1635
      End
      Begin VB.Label lblFieldName 
         AutoSize        =   -1  'True
         Caption         =   "Enter Data for Field:"
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
         TabIndex        =   11
         Top             =   810
         Width           =   1440
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
         TabIndex        =   10
         Top             =   270
         Width           =   2595
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3885
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2415
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
         TabIndex        =   5
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
         TabIndex        =   4
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
            ItemData        =   "frmModify.frx":1108
            Left            =   150
            List            =   "frmModify.frx":110A
            TabIndex        =   1
            Top             =   330
            Width           =   2055
         End
      End
   End
End
Attribute VB_Name = "frmModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim Numfields As Integer
Dim Currfield As Integer
Dim RecPos As Integer
Dim Mypassword As String
Private Sub OpenTable()
'On Error GoTo errHandler:
    rsName = Combo1.Text
    RsOpen
    RecOpen = True
    Numfields = rs.Fields.Count
    lstField.Clear
    For i = 0 To Numfields - 1
        lstField.AddItem (rs.Fields(i).Name)
    Next i
   Edit_Record
   
Exit Sub
'errHandler:
'MsgBox "Sorry that Table Name does not exist.", vbCritical, "Modify Records"
'Exit Sub
End Sub
Private Sub cmdAction_Click()
On Error GoTo errHandler:
If cmdAction.Caption = "Update" Then
    rs.Fields(Currfield) = txtData.Text
    rs.Update
    txtData.Enabled = False
    cmdAction.Caption = "Edit"
ElseIf cmdAction.Caption = "Edit" Then
    rs.Edit
    cmdAction.Caption = "Update"
    txtData.Enabled = True
End If
Exit Sub
errHandler:
MsgBox err.Description, vbCritical
Exit Sub
End Sub
Private Sub cmdBack_Click()
Unload Me
frmMain.Show
End Sub
Private Sub cmdNext_Click()
Next_Field
End Sub
Private Sub cmdNextRec_Click()
On Error GoTo errHandler:
If Not rs.EOF Then
    rs.MoveNext
    RecPos = RecPos + 1
    lblFieldName.Caption = "Enter Data for Field -> " & rs.Fields(Currfield).Name
    If Not IsNull(rs.Fields(Currfield)) Then
        txtData.Text = rs.Fields(Currfield)
    Else
        txtData.Text = ""
    End If
    lblDPosition.Caption = CStr(RecPos) + " of " + CStr(rs.RecordCount)
End If
Exit Sub
errHandler:
MsgBox err.Description, vbCritical
rs.MoveLast
RecPos = rs.RecordCount
Exit Sub
End Sub
Private Sub cmdPrev_Click()
Prev_field
End Sub
Private Sub cmdPrevRec_Click()
On Error GoTo errHandler:
If Not rs.EOF Then
    rs.MovePrevious
    RecPos = RecPos - 1
    lblFieldName.Caption = "Enter Data for Field -> " & rs.Fields(Currfield).Name
    If Not IsNull(rs.Fields(Currfield)) Then
        txtData.Text = rs.Fields(Currfield)
    Else
        txtData.Text = ""
    End If
    lblDPosition.Caption = CStr(RecPos) + " of " + CStr(rs.RecordCount)
End If
Exit Sub
errHandler:
MsgBox err.Description, vbCritical
rs.MoveFirst
RecPos = 1
Exit Sub
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
cmdPrev.Enabled = False
cmdPrevRec.Enabled = False
cmdNextRec.Enabled = False
cmdAction.Enabled = False
cmdNext.Enabled = False
Combo1.Enabled = False
err.Clear
frmMain.Show


End Sub

Private Sub txtTableName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then OpenTable
End Sub
Private Sub Prev_field()
On Error GoTo errHandler:
If Not (Currfield - 1) < 0 Then
    cmdNext.Enabled = True
    Currfield = Currfield - 1
    lblFieldName.Caption = "Enter Data for Field -> " & rs.Fields(Currfield).Name
    If Not IsNull(rs.Fields(Currfield)) Then
        txtData.Text = rs.Fields(Currfield)
    Else
        txtData.Text = ""
    End If
End If

If Currfield - 1 = Numfields Then
    cmdPrev.Enabled = False
End If
Exit Sub
errHandler:
MsgBox err.Description, vbCritical, "Add Record"
Exit Sub
End Sub
Private Sub Next_Field()
On Error GoTo errHandler:
If Not (Currfield + 1) > Numfields Then
    cmdPrev.Enabled = True
    Currfield = Currfield + 1
    lblFieldName.Caption = "Enter Data for Field -> " & rs.Fields(Currfield).Name
    If Not IsNull(rs.Fields(Currfield)) Then
        txtData.Text = rs.Fields(Currfield)
    Else
        txtData.Text = ""
    End If
End If
If Currfield + 1 = Numfields Then

    cmdNext.Enabled = False
End If
Exit Sub
errHandler:
MsgBox err.Description, vbCritical, "Add Record"
Exit Sub
End Sub
Private Sub Edit_Record()
If rs.RecordCount > 0 Then
Currfield = 0
RecPos = 1
lblDPosition.Caption = CStr(RecPos) + " of " + CStr(rs.RecordCount)
lblFieldName.Caption = "Enter Data for Field -> " & rs.Fields(Currfield).Name
If Not IsNull(rs.Fields(Currfield)) Then
        txtData.Text = rs.Fields(Currfield)
    Else
        txtData.Text = ""
End If
Else
    lblDPosition.Caption = "0 of 0"
End If
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

