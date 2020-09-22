VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3120
   Icon            =   "frmPasswordVEH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3120
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   3075
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         IMEMode         =   3  'DISABLE
         Left            =   210
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   510
         Width           =   2685
      End
      Begin VB.CheckBox Check1 
         Height          =   285
         Left            =   2700
         TabIndex        =   1
         Top             =   1140
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Password:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   3
         Top             =   210
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
DbPass = txtPassword.Text
Unload Me
frmMain.Show
End Sub

