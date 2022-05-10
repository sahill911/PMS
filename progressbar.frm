VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13350
   FillColor       =   &H00800000&
   Icon            =   "progressbar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   13350
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   10320
      Top             =   3120
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   3720
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   1085
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400000&
      Caption         =   "WELCOME TO PHARMACY MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   10335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   2400
      Width           =   8295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
If ProgressBar1 < 99 Then ProgressBar1 = ProgressBar1 + 1 / 8
If ProgressBar1 = 10 Then Label1.Caption = "  WELCOME TO PHARMACY MANAGEMENT SYSTEM"
If ProgressBar1 = 20 Then Label1.Caption = "  Loading program....."
If ProgressBar1 = 30 Then Label1.Caption = "  Validating your data"
If ProgressBar1 = 40 Then Label1.Caption = "  Scanning system restore"
If ProgressBar1 = 70 Then Label1.Caption = "  Creating restore points"
If ProgressBar1 = 85 Then Label1.Caption = "  Alomost done!!"
If ProgressBar1 = 98 Then frmLogin.Show
If ProgressBar1 = 99 Then Unload Me

End Sub
