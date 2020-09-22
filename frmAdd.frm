VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Records"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
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
   ScaleHeight     =   3900
   ScaleWidth      =   5880
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTableName 
      Height          =   315
      Left            =   150
      TabIndex        =   12
      ToolTipText     =   "Enter TableName and Press Enter"
      Top             =   3390
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main"
      Height          =   465
      Left            =   4890
      TabIndex        =   10
      Top             =   3330
      Width           =   795
   End
   Begin VB.Frame Frame4 
      Caption         =   "Add Records"
      Height          =   3225
      Left            =   2430
      TabIndex        =   4
      Top             =   -30
      Width           =   3435
      Begin VB.CommandButton cmdAction 
         Height          =   465
         Left            =   1320
         TabIndex        =   9
         Top             =   2550
         Width           =   795
      End
      Begin VB.CommandButton cmdPrev 
         Height          =   465
         Left            =   240
         Picture         =   "frmAdd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2550
         Width           =   795
      End
      Begin VB.CommandButton cmdNext 
         Height          =   465
         Left            =   2430
         Picture         =   "frmAdd.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2550
         Width           =   825
      End
      Begin VB.TextBox txtData 
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1170
         Width           =   3015
      End
      Begin VB.Label lblDataBase 
         Caption         =   "DatbaseFileName"
         Height          =   525
         Left            =   240
         TabIndex        =   11
         Top             =   270
         Width           =   2595
      End
      Begin VB.Label lblFieldName 
         AutoSize        =   -1  'True
         Caption         =   "Enter Data for Field:"
         Height          =   360
         Left            =   240
         TabIndex        =   5
         Top             =   810
         Width           =   2970
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.Frame Frame2 
         Caption         =   "Table Field(s)"
         Height          =   2385
         Left            =   0
         TabIndex        =   2
         Top             =   810
         Width           =   2355
         Begin VB.ListBox lstField 
            Height          =   1740
            ItemData        =   "frmAdd.frx":0884
            Left            =   150
            List            =   "frmAdd.frx":0886
            TabIndex        =   3
            Top             =   330
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Table Name"
         Height          =   825
         Left            =   0
         TabIndex        =   1
         Top             =   -30
         Width           =   2355
         Begin VB.ComboBox Combo1 
            Height          =   330
            Left            =   150
            TabIndex        =   13
            Text            =   "Select Table Name"
            Top             =   300
            Width           =   2115
         End
      End
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim Numfields As Integer
Dim Currfield As Integer
Private Sub cmdAction_Click()
On Error GoTo errHandler:
    If cmdAction.Caption = "&Cancel" Then
        rs.CancelUpdate
        cmdNext.Enabled = False
        cmdPrev.Enabled = False
        cmdBack.Enabled = True
        cmdAction.Caption = "&Add"
    ElseIf cmdAction.Caption = "&Save" Then
        rs.Fields(Currfield) = txtData.Text
        rs.Update
        cmdNext.Enabled = False
        cmdPrev.Enabled = False
        cmdBack.Enabled = True
        cmdAction.Caption = "&Add"
    ElseIf cmdAction.Caption = "&Add" Then
        Add_Record
    End If
Exit Sub
errHandler:
MsgBox Err.Description, vbCritical, "Add Records"
Exit Sub
End Sub
Private Sub cmdBack_Click()
Debug.Print DbOpen
Unload Me
frmMain.Show
End Sub
Private Sub cmdNext_Click()
Next_Field
End Sub
Private Sub cmdPrev_Click()
Prev_field
End Sub

Private Sub Combo1_Click()
OpenTable
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

Private Sub Form_Load()
On Error GoTo ErrH
Combo1.Enabled = True
OpenDbase
lblDataBase.Caption = DbFileName
RecOpen = False
Currfield = 0
Exit Sub

ErrH:
If Err.Number = 3031 Then
    MsgBox "The password you entered is incorrect.", vbCritical, "Open Database"
Else
    MsgBox Err.Description, vbCritical, "Open Database."
End If
cmdPrev.Enabled = False
cmdNext.Enabled = False
Combo1.Enabled = False
Err.Clear
frmMain.Show


End Sub
Private Sub Form_Unload(Cancel As Integer)
If RecOpen = True Then rs.Close
'Db.Close
End Sub
Private Sub txtTableName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    OpenTable
End If
End Sub
Private Sub OpenTable()
On Error GoTo errHandler:
    rsName = Combo1.Text
    RsOpen
    Numfields = rs.Fields.Count
    lstField.Clear
    For i = 0 To Numfields - 1
        lstField.AddItem (rs.Fields(i).Name)
    Next i
    Add_Record
Exit Sub
errHandler:
MsgBox "Sorry that Table Name does not exist.", vbCritical, "Add Records"
Exit Sub
End Sub
Private Sub Add_Record()
cmdBack.Enabled = False
cmdNext.Enabled = True
cmdPrev.Enabled = True
Currfield = 0
rs.AddNew
lblFieldName.Caption = "Enter Data for Field -> " & rs.Fields(Currfield).Name
txtData.SetFocus
cmdAction.Caption = "&Cancel"
End Sub
Private Sub Next_Field()
On Error GoTo errHandler:
If Not (Currfield + 1) > Numfields Then
    cmdPrev.Enabled = True
    rs.Fields(Currfield) = txtData.Text
    Currfield = Currfield + 1
    lblFieldName.Caption = "Enter Data for Field -> " & rs.Fields(Currfield).Name
    If Not IsNull(rs.Fields(Currfield)) Then
        txtData.Text = rs.Fields(Currfield)
    Else
        'rs.Fields(Currfield) = txtData.Text
    End If
End If
If Currfield + 1 = Numfields Then
    cmdNext.Enabled = False
    cmdAction.Caption = "&Save"
End If
Exit Sub
errHandler:
MsgBox Err.Description, vbCritical, "Add Record"
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
MsgBox Err.Description, vbCritical, "Add Record"
Exit Sub
End Sub

