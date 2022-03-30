VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form staff_rt_rn_frm 
   Caption         =   "s"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton rt_btn 
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
      Left            =   15480
      TabIndex        =   6
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton rn_btn 
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
      Left            =   15480
      TabIndex        =   5
      Top             =   6600
      Width           =   1935
   End
   Begin VB.TextBox id_txt 
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      Top             =   6600
      Width           =   2895
   End
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
      Left            =   2520
      TabIndex        =   1
      Top             =   7680
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4935
      Left            =   2160
      TabIndex        =   0
      Top             =   1080
      Width           =   15255
      _ExtentX        =   26908
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
      Height          =   495
      Left            =   9480
      TabIndex        =   9
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Fine : Rs."
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
      Left            =   8520
      TabIndex        =   8
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   9480
      TabIndex        =   7
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Staff ID"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Borrowed Books (STAFFs)"
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
      Left            =   9000
      TabIndex        =   3
      Top             =   600
      Width           =   2895
   End
End
Attribute VB_Name = "staff_rt_rn_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim doi As Date
'used to store difference days b/w nowdate and date of issue
Dim d As Integer

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

Private Sub DataGrid1_Click()
Call show_fine
End Sub

Private Sub Form_Load()
    exec_query ("select * from StaffEntry")
Set DataGrid1.DataSource = rs
    DataGrid1.Refresh
    Call show_fine
End Sub

Private Sub rn_btn_Click()
 Dim today As Date
 today = Format(Now, "dd/mm/yyyy")
 rs.Fields("DOI") = today
 rs.Update
 MsgBox "RENEWED"
End Sub

Private Sub srch_btn_Click()
If id_txt.Text <> "" Then
    sqlquery = "SELECT * FROM StaffEntry WHERE id = '" & id_txt.Text & "'"
    exec_query (sqlquery)
    If (Not rs.EOF) Then
    Set DataGrid1.DataSource = rs
    DataGrid1.Refresh
    Else
        MsgBox ("Not Found , Try Again !")
    End If
    Call show_fine
End If

End Sub

Private Sub rt_btn_Click()
Dim ans As Integer
ans = MsgBox("Are you sure ", vbQuestion + vbYesNo + vbDefaultButton1, "confirmation")
If ans = vbYes Then
rs.Delete
Call show_fine
End If
End Sub

