VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ret_ren_frm 
   Caption         =   "Renew / Return a Book"
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15765
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   15765
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton ret_btn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15120
      TabIndex        =   5
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton ren_btn 
      Caption         =   "Renew"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15120
      TabIndex        =   4
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton search_btn 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   3
      Top             =   7440
      Width           =   1935
   End
   Begin VB.TextBox regno_txt 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   1
      Top             =   6600
      Width           =   2895
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4935
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   8705
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   28
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myanmar Text"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Rs."
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   11640
      TabIndex        =   9
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label fine_lbl 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   12120
      TabIndex        =   8
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Fine :"
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
      Left            =   10920
      TabIndex        =   7
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Borrowed Books"
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
      Left            =   8280
      TabIndex        =   6
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Register Number / Admission Number"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   6600
      Width           =   3135
   End
End
Attribute VB_Name = "ret_ren_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim doi As Date
'used to store difference days b/w nowdate and date of issue
Dim d As Integer

Private Sub DataGrid1_Click()
    Call show_fine
End Sub

Private Sub Form_Load()
    exec_query ("select * from StudentEntry")
Set DataGrid1.DataSource = rs
    DataGrid1.Refresh
    Call show_fine
End Sub

Private Sub show_fine()
If (Not rs.EOF) Then
    doi = CDate(rs.Fields("DOI"))
    d = DateDiff("d", doi, Format(Now, "dd-mm-yyyy"))
    If (d > day_limit) Then
        fine_lbl.Caption = CStr(d * fine_amount)
        fine_lbl.ForeColor = &HFF&
    Else
        fine_lbl.Caption = "0"
        fine_lbl.ForeColor = &H80000012
    End If
Else
    fine_lbl.Caption = "0"
    fine_lbl.ForeColor = &H80000012
End If
End Sub

Private Sub ren_btn_Click()
 Dim today As Date
 today = Format(Now, "dd/mm/yyyy")
 rs.Fields("DOI") = today
'rs.Fields("DOR") = DateAdd("d", 15, today)
 rs.Update
 MsgBox "RENEWED"
End Sub

Private Sub search_btn_Click()
If regno_txt.Text <> "" Then
    sqlquery = "SELECT * FROM StudentEntry WHERE RegisterNo = '" & regno_txt.Text & "'"
    exec_query (sqlquery)
    If (Not rs.EOF) Then
    Set DataGrid1.DataSource = rs
    DataGrid1.Refresh
    Else
        MsgBox ("Not Found\nTry Again")
    End If
End If
Call show_fine
End Sub

Private Sub ret_btn_Click()
Dim ans As Integer
ans = MsgBox("Are you sure ", vbQuestion + vbYesNo + vbDefaultButton1, "confirmation")
If ans = vbYes Then
rs.Delete
End If
End Sub
