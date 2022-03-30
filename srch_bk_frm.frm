VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form srch_bk_frm 
   Caption         =   "Search"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15135
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Book Details"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   840
      TabIndex        =   5
      Top             =   6480
      Width           =   17655
      Begin VB.CommandButton rm_btn 
         Caption         =   "Remove"
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
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label auth_lbl 
         Caption         =   "********************"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label lbl 
         Caption         =   "Author  :"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label nm_lbl 
         Caption         =   "*** **** ********"
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
         Left            =   1200
         TabIndex        =   7
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "Name     :"
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
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Options"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   14760
      TabIndex        =   1
      Top             =   720
      Width           =   3735
      Begin VB.CommandButton srch_btn 
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
         Left            =   1200
         TabIndex        =   4
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ComboBox dept_com 
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         ItemData        =   "srch_bk_frm.frx":0000
         Left            =   240
         List            =   "srch_bk_frm.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5055
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   8916
      _Version        =   393216
      AllowUpdate     =   0   'False
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
End
Attribute VB_Name = "srch_bk_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub showBookDet()
If (Not rs.EOF) Then
nm_lbl.Caption = rs.Fields(1)
auth_lbl.Caption = rs.Fields(2)
Else
nm_lbl.Caption = " "
auth_lbl.Caption = " "
End If
End Sub

Private Sub DataGrid1_Click()
Call showBookDet
End Sub

Private Sub Form_Load()

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

'data grid
exec_query ("SELECT * FROM Books ORDER BY AccNo ASC")
Set DataGrid1.DataSource = rs
    DataGrid1.Refresh
Call showBookDet
End Sub


Private Sub rm_btn_Click()

Dim ans As Integer
ans = MsgBox("Are you sure?", vbQuestion + vbYesNo + vbDefaultButton1, "Remove Book")
If (ans = vbYes) Then
rs.Delete
DataGrid1.Refresh
End If

End Sub

Private Sub srch_btn_Click()

If (dept_com.Text <> "") Then
If (dept_com.Text = "All") Then
    exec_query ("SELECT * FROM Books ORDER BY AccNo ASC")
    Set DataGrid1.DataSource = rs
    DataGrid1.Refresh
Else
    exec_query ("SELECT * FROM Books WHERE Department='" + dept_com.Text + "' ORDER BY AccNo ASC")
    Set DataGrid1.DataSource = rs
    DataGrid1.Refresh
End If
Else
MsgBox "Select the department first !"
End If
Call showBookDet
End Sub
