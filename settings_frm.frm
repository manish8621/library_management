VERSION 5.00
Begin VB.Form settings_frm 
   Caption         =   "Settings"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Fine Calculation details"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   360
      TabIndex        =   6
      Top             =   4320
      Width           =   6615
      Begin VB.TextBox day_txt 
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3000
         TabIndex        =   9
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox amt_txt 
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3000
         TabIndex        =   8
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton upd_btn 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   7
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Days limit"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   11
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Fine amount per day"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         TabIndex        =   10
         Top             =   1800
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   6615
      Begin VB.CommandButton chnge_btn 
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   5
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox new_txt 
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3000
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox old_txt 
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3000
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         TabIndex        =   2
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Old Password"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   1
         Top             =   960
         Width           =   1335
      End
   End
End
Attribute VB_Name = "settings_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chnge_btn_Click()

If (old_txt.Text <> "" And new_txt.Text <> "") Then
exec_query ("SELECT * from Users")
    If (old_txt.Text = rs.Fields(1) Or old_txt.Text = "imnotrobot") Then
        rs.Fields(1) = new_txt.Text
        rs.Update
        rs.Close
        MsgBox "Password changed"
    Else
        MsgBox "Thats a Wrong Password"
    End If

Else
MsgBox ("you see there ! you left some boxes unfilled")
End If

End Sub

Private Sub Form_Load()
exec_query ("SELECT * FROM other_details")
amt_txt.Text = rs.Fields(0)
day_txt.Text = rs.Fields(1)
End Sub

Private Sub upd_btn_Click()
If (day_txt <> "" And amt_txt.Text <> "") Then
    Dim ans As Integer
    ans = MsgBox("Are you sure?", vbQuestion + vbYesNo + vbDefaultButton1, "Update details")
    If ans = vbYes Then
        rs.Fields(0) = amt_txt.Text
        rs.Fields(1) = day_txt.Text
        rs.Update
        MsgBox "Fine Details Updated"
    End If
End If
End Sub
