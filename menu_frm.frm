VERSION 5.00
Begin VB.Form menu_frm 
   Caption         =   "Form1"
   ClientHeight    =   9555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15855
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton no_d_btn 
      Caption         =   "No Due"
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
      Left            =   6120
      TabIndex        =   14
      Top             =   3960
      Width           =   3135
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Settings"
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
      Left            =   0
      TabIndex        =   13
      Top             =   9240
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Return/remove"
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
      Left            =   9960
      TabIndex        =   12
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add Entry"
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
      Left            =   9960
      TabIndex        =   11
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CommandButton srch_bk_btn 
      Caption         =   "Search"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CommandButton add_n_btn 
      Caption         =   "Add /  Remove"
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
      Left            =   13560
      TabIndex        =   8
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CommandButton attn_btn 
      Caption         =   "Attendance"
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
      Left            =   13560
      TabIndex        =   6
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
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
      Left            =   2520
      TabIndex        =   4
      Top             =   3960
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Entry"
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
      Left            =   6120
      TabIndex        =   3
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CommandButton return_btn 
      Caption         =   "Return/remove"
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
      Left            =   6120
      TabIndex        =   2
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CommandButton addbook_btn 
      Caption         =   "Add"
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
      Left            =   2520
      TabIndex        =   1
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Staff"
      Height          =   375
      Left            =   10080
      TabIndex        =   10
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Newspapers"
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
      Left            =   13560
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Student Entry"
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
      Left            =   6120
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Books"
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
      Left            =   2520
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "menu_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_n_btn_Click()
ad_rm_np_frm.Show
End Sub

Private Sub addbook_btn_Click()
ad_bk_frm.Show
End Sub

Private Sub attn_btn_Click()
attn_frm.Show
End Sub

Private Sub Command1_Click()
ad_ent_frm.Show
End Sub

Private Sub Command2_Click()
srch_bk_frm.Show
End Sub

Private Sub Command3_Click()
staff_ent_frm.Show
End Sub

Private Sub Command4_Click()
staff_rt_rn_frm.Show
End Sub

Private Sub Command5_Click()
settings_frm.Show
End Sub

Private Sub no_d_btn_Click()
ret_ren_frm.Show
End Sub

Private Sub return_btn_Click()
ret_ren_frm.Show
End Sub

Private Sub srch_bk_btn_Click()
srch_bk_frm.Show
End Sub
