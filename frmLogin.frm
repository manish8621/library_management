VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1605
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   948.287
   ScaleMode       =   0  'User
   ScaleWidth      =   4048
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Text            =   "admin"
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "admin"
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Username"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim usr As String
Dim pwd As String
Dim clean_year As String


Private Sub Form_Load()
Call init_db
exec_query ("SELECT username,password FROM Users")
usr = rs.Fields(0)
pwd = rs.Fields(1)
exec_query ("SELECT * from other_details")
fine_amount = rs.Fields(0)
day_limit = rs.Fields(1)
clean_year = rs.Fields(2)
Call DB_clean
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    If txtUserName.Text = usr And txtPassword.Text = pwd Then
        menu_frm.Show
        Unload Me
    Else
        MsgBox "Invalid Password, try again!"
    End If
End Sub
Private Sub DB_clean()
If (CStr(Format(Now, "yyyy")) <> clean_year) Then
'MsgBox "Cleaning Database"
exec_query ("DELETE FROM Attendance")
MsgBox "Cleaning Database Finished"
exec_query ("SELECT * from other_details")
rs.Fields(2) = CStr(Format(Now, "yyyy"))
rs.Update
End If
End Sub

Private Sub cmdCancel_Click()
End
End Sub


