VERSION 5.00
Begin VB.Form ad_bk_frm 
   Caption         =   "Add Books"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   Picture         =   "ad_rem_frm.frx":0000
   ScaleHeight     =   7485
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox dept_com 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "ad_rem_frm.frx":0342
      Left            =   3000
      List            =   "ad_rem_frm.frx":0349
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton add_btn 
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
      Left            =   1680
      TabIndex        =   13
      Top             =   6600
      Width           =   2295
   End
   Begin VB.TextBox accno_txt 
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
      Left            =   3000
      TabIndex        =   6
      Text            =   "104"
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox title_txt 
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
      Left            =   3000
      TabIndex        =   5
      Text            =   "eee"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox author_txt 
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
      Left            =   3000
      TabIndex        =   4
      Text            =   "ee"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox pub_txt 
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
      Left            =   3000
      TabIndex        =   3
      Text            =   "ee"
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox ed_txt 
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
      Left            =   3000
      TabIndex        =   2
      Text            =   "1"
      Top             =   4800
      Width           =   2055
   End
   Begin VB.ComboBox ref_com 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "ad_rem_frm.frx":0362
      Left            =   3000
      List            =   "ad_rem_frm.frx":036C
      TabIndex        =   1
      Text            =   "NO"
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Book Title"
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
      Left            =   600
      TabIndex        =   12
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Department"
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
      Left            =   600
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Author"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Publications"
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
      Left            =   600
      TabIndex        =   9
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Edition"
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
      Left            =   600
      TabIndex        =   8
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Reference Book"
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
      Left            =   600
      TabIndex        =   7
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Acc No"
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
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "ad_bk_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'add error handle when duplicate value inserted

Dim valid_details As Boolean
Dim is_reference As String
Private Sub add_records()
exec_query ("SELECT * FROM Books")

rs.AddNew
rs.Fields(0) = accno_txt.Text
rs.Fields(1) = title_txt.Text
rs.Fields(2) = dept_com.Text
rs.Fields(3) = author_txt.Text
rs.Fields(4) = pub_txt.Text
rs.Fields(5) = Val(ed_txt.Text)
rs.Fields(6) = is_reference

rs.Update

MsgBox "Book Added"
'clear all fields
accno_txt.Text = ""
title_txt.Text = ""
author_txt.Text = ""
pub_txt.Text = ""
ed_txt.Text = ""
ref_com.Text = ""
End Sub

Private Sub add_btn_Click()
Dim q As String
valid_details = True
'validate details
If (ref_com.Text = "YES") Then
    is_reference = "Y"
Else
If (ref_com.Text = "NO") Then
    is_reference = "N"
Else
    valid_details = False
End If
End If

If dept_com.Text = "Select a Department" Then
valid_details = False
End If
If (accno_txt.Text = "" Or title_txt.Text = "" Or dept_com.Text = "" Or author_txt.Text = "" Or pub_txt.Text = "" Or ed_txt.Text = "") Then
valid_details = False
End If
'adding record to db
If valid_details Then
    Call add_records
Else
MsgBox "Invalid Details"
End If
End Sub

Private Sub Form_Load()
dept_com.Text = "Select a Department"

'update depts
exec_query ("SELECT DISTINCT dept FROM Depts")
If (Not rs.EOF) Then
    While (Not rs.EOF)
        dept_com.AddItem (rs.Fields(0))
        rs.MoveNext
    Wend
Else
    MsgBox "Error in retrieving department list"
End If
End Sub
