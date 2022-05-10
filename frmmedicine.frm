VERSION 5.00
Begin VB.Form frmmedicine 
   Caption         =   "medicine"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17565
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   17565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   " &Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   9840
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   2640
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   4200
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2640
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   3120
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   645
      Left            =   2640
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1920
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label Label5 
      Caption         =   "Total"
      Height          =   375
      Left            =   9840
      TabIndex        =   8
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Drug Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Index no"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frmmedicine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

