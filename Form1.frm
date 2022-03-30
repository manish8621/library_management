VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form attn_frm 
   Caption         =   "Newspaper Attendance Report"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Attendance  Report"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   7440
      TabIndex        =   1
      Top             =   600
      Width           =   10215
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2940
         ItemData        =   "Form1.frx":0000
         Left            =   7440
         List            =   "Form1.frx":0002
         TabIndex        =   3
         Top             =   3840
         Width           =   1935
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   3375
         Left            =   480
         TabIndex        =   2
         Top             =   3600
         Width           =   5055
         _Version        =   524288
         _ExtentX        =   8916
         _ExtentY        =   5953
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
      Begin VB.Label date_lbl 
         Alignment       =   2  'Center
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
         Left            =   7440
         TabIndex        =   7
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Select a Date"
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
         Left            =   600
         TabIndex        =   6
         Top             =   2640
         Width           =   4815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Attendance"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   4935
      Begin VB.CommandButton clr_btn 
         Caption         =   "Click here"
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
         Left            =   2880
         TabIndex        =   9
         Top             =   7800
         Width           =   1455
      End
      Begin VB.CommandButton sbmt_btn 
         Caption         =   "Submit"
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
         Left            =   1560
         TabIndex        =   5
         Top             =   6720
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Thandhi"
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
         Index           =   0
         Left            =   1200
         TabIndex        =   4
         Top             =   1250
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "To clear today attendance "
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
         Left            =   720
         TabIndex        =   8
         Top             =   7920
         Width           =   2055
      End
   End
End
Attribute VB_Name = "attn_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim npapers() As String
Dim n As Integer

Private Sub Calendar1_Click()
List1.Clear
date_lbl.Caption = CStr(Calendar1.Value)
exec_query ("select NewspaperName from Attendance where ddate= #" + CStr(Calendar1.Value) + "#")
If (Not rs.EOF) Then
While (Not rs.EOF)
    List1.AddItem (rs.Fields(0))
    rs.MoveNext
Wend
Else
    List1.AddItem ("--no record found--")
End If
End Sub


Private Sub clr_btn_Click()
'clear the todays records
Dim ans As Integer
ans = MsgBox("Are you sure?", vbQuestion + vbYesNo + vbDefaultButton1, "Clearing Today Record")
If ans = vbYes Then
exec_query ("DELETE FROM Attendance WHERE ddate=#" + CStr(Format(Now, "dd-mm-yyyy")) + "#")
End If
End Sub

Private Sub Form_Load()
'To get list of nps
exec_query ("SELECT NewspaperName FROM Newspapers")
n = rs.RecordCount
rs.MoveFirst
'setting size of nps list by num of record
ReDim npapers(n) As String
'getting newsspaper names from record to a local arry
For i = 0 To (n - 1)
    npapers(i) = rs.Fields(0)
    rs.MoveNext
Next i

'Dynamically creating checkboxes
For i = 0 To (n - 1)
If (i <> 0) Then
Load Check1(i)
Check1(i).Height = Check1(0).Height
Check1(i).Top = Check1(i - 1).Height + Check1(i - 1).Top
Check1(i).Left = 1200
Check1(i).Width = Check1(0).Width
Check1(i).Caption = npapers(i)
Check1(i).Visible = True
Else
Check1(i).Caption = npapers(i)
Check1(i).Visible = True
End If
Next i
'setting button to bottom position of last checkbox
sbmt_btn.Top = Check1(n - 1).Top + Check1(n - 1).Height

End Sub

Private Sub sbmt_btn_Click()
Dim ans As Integer
ans = MsgBox("Are you sure ", vbQuestion + vbYesNo + vbDefaultButton1, "confirmation")
If ans = vbYes Then
    exec_query ("SELECT * FROM Attendance")
    For i = 0 To n - 1
    If Check1(i).Value Then
        rs.AddNew
        rs.Fields(0) = Check1(i).Caption
        rs.Fields(1) = Format(Now, "dd-mm-yyyy")
        rs.Update
    End If
    Next i
    MsgBox "ADDED"
End If
End Sub
