VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form ad_mag_jr_frm 
   Caption         =   "Add Magazine or Journal"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19020
   BeginProperty Font 
      Name            =   "Myanmar Text"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   19020
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   2640
      TabIndex        =   0
      Top             =   600
      Width           =   12255
      Begin MSACAL.Calendar Calendar1 
         Height          =   2655
         Left            =   7200
         TabIndex        =   14
         Top             =   4560
         Visible         =   0   'False
         Width           =   3975
         _Version        =   524288
         _ExtentX        =   7011
         _ExtentY        =   4683
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2022
         Month           =   3
         Day             =   23
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
         Height          =   615
         Left            =   3240
         TabIndex        =   7
         Top             =   6600
         Width           =   1695
      End
      Begin VB.ComboBox sub_type_com 
         Height          =   420
         ItemData        =   "ad_mag_jr_frm.frx":0000
         Left            =   4200
         List            =   "ad_mag_jr_frm.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   5400
         Width           =   2775
      End
      Begin VB.TextBox lb_d_txt 
         Height          =   495
         Left            =   4200
         TabIndex        =   5
         Top             =   4560
         Width           =   2775
      End
      Begin VB.TextBox sub_fee_txt 
         Height          =   495
         Left            =   4200
         TabIndex        =   4
         Top             =   3600
         Width           =   2775
      End
      Begin VB.ComboBox dept_com 
         Height          =   420
         ItemData        =   "ad_mag_jr_frm.frx":0004
         Left            =   4200
         List            =   "ad_mag_jr_frm.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2760
         Width           =   2775
      End
      Begin VB.TextBox nm_txt 
         Height          =   495
         Left            =   4200
         TabIndex        =   2
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox sn_txt 
         Height          =   465
         Left            =   4200
         TabIndex        =   1
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Type of Subscription"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         TabIndex        =   13
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Bought Date"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         TabIndex        =   12
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Subscription Fees"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   11
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Height          =   615
         Left            =   2280
         TabIndex        =   10
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Name Of Journal or Magazine"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   9
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Serial Number"
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
         Left            =   2160
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
   End
End
Attribute VB_Name = "ad_mag_jr_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n_moth As Integer
Dim due_date As Date

Private Sub add_record()
    exec_query ("select * from MagazinesAndJ")
    rs.AddNew
    rs.Fields(0) = sn_txt.Text
    rs.Fields(1) = nm_txt.Text
    rs.Fields(2) = dept_com.Text
    rs.Fields(3) = sub_fee_txt.Text
    rs.Fields(4) = lb_d_txt.Text
    rs.Fields(5) = sub_type_com.Text
    rs.Fields(6) = due_date
    rs.Update
    MsgBox ("ADDED")
    sn_txt.Text = ""
    nm_txt.Text = ""
    sub_fee_txt.Text = ""
    
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
sub_type_com.AddItem ("Monthly")
sub_type_com.AddItem ("Quaterly")
sub_type_com.AddItem ("Halfly")
sub_type_com.AddItem ("Annualy")
End Sub

Private Sub lb_d_txt_GotFocus()
Calendar1.Visible = True
End Sub

Private Sub Calendar1_Click()
lb_d_txt.Text = Calendar1.Value
End Sub

Private Sub lb_d_txt_LostFocus()
Calendar1.Visible = False
End Sub

Private Sub sbmt_btn_Click()

Dim valid_detail As Boolean
valid_detail = True
If (sn_txt.Text = "" Or nm_txt.Text = "" Or dept_com.Text = "" Or sub_fee_txt.Text = "" Or lb_d_txt.Text = "" Or sub_type_com.Text = "") Then
valid_detail = False
End If
If valid_detail Then
    'calculating due date
    If (sub_type_com = "Monthly") Then
        n_month = 1
    Else
    If (sub_type_com = "Quaterly") Then
        n_month = 3
    Else
    If (sub_type_com = "Halfly") Then
        n_month = 6
    Else
        n_month = 12
    End If
    End If
    End If
    due_date = DateAdd("m", n_month, CDate(lb_d_txt.Text))
    
    Dim ans As Integer
    ans = MsgBox("Are you sure?", vbQuestion + vbYesNo + vbDefaultButton1, "Add Magazine/Journal")
    If ans = vbYes Then
        Call add_record
    End If
    
    Else
        MsgBox "Fill all the details"
    End If

End Sub
