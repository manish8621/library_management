VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form ad_ent_frm 
   Caption         =   "Make Entry"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox dgr_com 
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
      ItemData        =   "ad_ent_frm.frx":0000
      Left            =   8400
      List            =   "ad_ent_frm.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3960
      Width           =   2055
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
      ItemData        =   "ad_ent_frm.frx":001C
      Left            =   8400
      List            =   "ad_ent_frm.frx":0023
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   3360
      Width           =   2055
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2295
      Left            =   10560
      TabIndex        =   13
      Top             =   6240
      Visible         =   0   'False
      Width           =   3975
      _Version        =   524288
      _ExtentX        =   7011
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2022
      Month           =   3
      Day             =   14
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   0
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
      Height          =   495
      Left            =   8400
      TabIndex        =   12
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton s_btn 
      Caption         =   "submit"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   11
      Top             =   7440
      Width           =   1815
   End
   Begin VB.TextBox regno_txt 
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
      Left            =   8400
      TabIndex        =   10
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox name_txt 
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
      Left            =   8400
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox yr_txt 
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
      Left            =   8400
      TabIndex        =   1
      Top             =   4680
      Width           =   2055
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
      Left            =   8400
      TabIndex        =   0
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Admission /Register number"
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
      Left            =   5880
      TabIndex        =   9
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
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
      Left            =   6480
      TabIndex        =   8
      Top             =   2520
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
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Degree"
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
      Left            =   6480
      TabIndex        =   6
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Year"
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
      Left            =   6480
      TabIndex        =   5
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Acc no"
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
      Left            =   6480
      TabIndex        =   4
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label7 
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
      Left            =   6480
      TabIndex        =   3
      Top             =   6360
      Width           =   1695
   End
End
Attribute VB_Name = "ad_ent_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Calendar1_Click()
doi_txt.Text = Calendar1.Value
End Sub


Private Sub dept_com_Click()

'update degrees
If (dept_com.Text <> "Select a Department") Then
    'clear before adding
    dgr_com.Clear
    exec_query ("SELECT Degree FROM Depts WHERE dept='" + dept_com.Text + "'")
    If (Not rs.EOF) Then
        While (Not rs.EOF)
            dgr_com.AddItem (rs.Fields(0))
            rs.MoveNext
        Wend
    Else
        MsgBox "Error in retrieving degrees list"
    End If
End If
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
dgr_com.Text = "Select a Degree"
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

Private Sub add_record()
exec_query ("SELECT * FROM StudentEntry")

rs.AddNew
rs.Fields(0) = regno_txt.Text
rs.Fields(1) = name_txt.Text
rs.Fields(2) = dept_com.Text
rs.Fields(3) = dgr_com.Text
rs.Fields(4) = Val(yr_txt.Text)
rs.Fields(5) = accno_txt.Text
rs.Fields(6) = doi_txt.Text
rs.Update
MsgBox "Book Entry Added"
'clear all fields
regno_txt.Text = ""
name_txt.Text = ""
yr_txt.Text = ""
accno_txt.Text = ""
End Sub

Private Sub s_btn_Click()
Dim valid_details As Boolean
Dim q As String
valid_details = True
'validate details
If (regno_txt.Text = "" Or name_txt.Text = "" Or dept_com.Text = "" Or dgr_com.Text = "" Or yr_txt.Text = "" Or accno_txt.Text = "") Then
    valid_details = False
End If

If dept_com.Text = "Select a Department" Or dgr_com.Text = "Select a Degree" Then
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
            ans = MsgBox("Book Title:" + rs.Fields(0) + ",  Are you sure", vbQuestion + vbYesNo + vbDefaultButton1, "confirmation")
            If ans = vbYes Then
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
