VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form returnedbk 
   Caption         =   "Returned Books Log"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Choose Which record to show"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   15480
      TabIndex        =   1
      Top             =   960
      Width           =   3495
      Begin VB.OptionButton stf_op 
         Caption         =   "Staff"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton stu_op 
         Caption         =   "Student"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6735
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   11880
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
   Begin VB.Label ban_lbl 
      Alignment       =   2  'Center
      Caption         =   "Returned Books - Students"
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
      Left            =   6360
      TabIndex        =   4
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "returnedbk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
stu_op.Value = True
exec_query ("SELECT * FROM Student_Entry_Log ORDER BY DOI asc")
Set DataGrid1.DataSource = rs
    DataGrid1.Refresh
End Sub

Private Sub stf_op_Click()
exec_query ("SELECT * FROM Staff_Entry_Log ORDER BY DOI asc")
Set DataGrid1.DataSource = rs
    DataGrid1.Refresh
    ban_lbl.Caption = "Returned Books - Staff"
End Sub

Private Sub stu_op_Click()
exec_query ("SELECT * FROM Student_Entry_Log ORDER BY DOI asc")
Set DataGrid1.DataSource = rs
    DataGrid1.Refresh
    ban_lbl.Caption = "Returned Books - Students"
End Sub
