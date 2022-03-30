VERSION 5.00
Begin VB.Form wait_frm 
   BackColor       =   &H0000FFFF&
   Caption         =   "DataBase Maintanance"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   8760
      Shape           =   2  'Oval
      Top             =   840
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   8760
      Shape           =   2  'Oval
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Shape Shape4 
      Height          =   1935
      Left            =   8760
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000004&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   10440
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "WARNING ! Dont Close or minimize the Application"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   5760
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cleaning the Database Please wait...."
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
      Left            =   8280
      TabIndex        =   0
      Top             =   4320
      Width           =   2655
   End
End
Attribute VB_Name = "wait_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
