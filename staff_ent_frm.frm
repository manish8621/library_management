VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form staff_ent_frm 
   Caption         =   "Staff Entry"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSACAL.Calendar Calendar1 
      Height          =   2895
      Left            =   11280
      TabIndex        =   11
      Top             =   5880
      Visible         =   0   'False
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2022
      Month           =   3
      Day             =   16
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
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
      ItemData        =   "staff_ent_frm.frx":0000
      Left            =   8160
      List            =   "staff_ent_frm.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3960
      Width           =   2895
   End
   Begin VB.CommandButton sbmt_btn 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   9
      Top             =   7920
      Width           =   1695
   End
   Begin VB.TextBox nme_txt 
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
      Left            =   8160
      TabIndex        =   8
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox doi_txt 
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
      Left            =   8160
      TabIndex        =   6
      Top             =   5880
      Width           =   2895
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
      Height          =   615
      Left            =   8160
      TabIndex        =   4
      Top             =   4800
      Width           =   2895
   End
   Begin VB.TextBox id_txt 
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
      Left            =   8160
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Book Entry ( For Staff Only*)"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   12
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Staff Name"
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
      Left            =   5760
      TabIndex        =   7
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Date of Issue"
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
      Left            =   5760
      TabIndex        =   5
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label Acc 
      Caption         =   "Accno"
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
      Left            =   5760
      TabIndex        =   3
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label de 
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
      Left            =   5760
      TabIndex        =   2
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Staff ID"
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
      Left            =   5760
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
End
Attribute VB_Name = "staff_ent_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub add_record()
exec_query ("SELECT * FROM StaffEntry")
rs.AddNew
rs.Fields(0) = id_txt.Text
rs.Fields(1) = dept_com.Text
rs.Fields(2) = accno_txt.Text
rs.Fields(3) = doi_txt.Text
rs.Fields(4) = nme_txt.Text
rs.Update
MsgBox "Book Entry Added"
'clear all fields
id_txt.Text = ""
accno_txt.Text = ""
nme_txt.Text = ""
End Sub




Private Sub doi_txt_Click()
Calendar1.Visible = True
End Sub

Private Sub doi_txt_LostFocus()
Calendar1.Visible = False
End Sub

Private Sub Form_Load()
doi_txt.Text = Format(Now, "dd-mm-yyyy")
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


Private Sub sbmt_btn_Click()
Dim valid_details As Boolean
valid_details = True
'validate details
If (id_txt.Text = "" Or nme_txt.Text = "" Or dept_com.Text = "Select a Department" Or doi_txt.Text = "" Or accno_txt.Text = "") Then
    valid_details = False
End If

If valid_details Then
    ' add checking if book exist and its not a reference book
    'checking if book exists
    exec_query ("select Title from Books WHERE AccNo = '" + accno_txt.Text + "'")
    If (Not rs.EOF) Then
        exec_query ("select isReference from Books WHERE AccNo = '" + accno_txt.Text + "'")
        If (rs.Fields(0) = "N") Then
        Dim ans As Integer
        ans = MsgBox("Book Title:" + bk + ",  Are you sure", vbQuestion + vbYesNo + vbDefaultButton1, "confirmation")
        If (ans = vbYes) Then
            'finally after all blocking walls
            Call add_record
        End If
        Else
            MsgBox ("it is a Reference Book")
        End If
    Else
        MsgBox ("no book found with acc no:" + accno_txt.Text)
    End If
Else
MsgBox "Invalid Details"
End If

End Sub
