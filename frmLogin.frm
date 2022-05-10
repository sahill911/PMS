VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   6810
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   10155
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4023.573
   ScaleMode       =   0  'User
   ScaleWidth      =   9534.995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   4800
      TabIndex        =   0
      Top             =   1320
      Width           =   2445
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FF0000&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3600
      TabIndex        =   2
      Top             =   2520
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5040
      TabIndex        =   3
      Top             =   2520
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   4800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1920
      Width           =   2445
   End
   Begin VB.Line Line5 
      X1              =   7098.431
      X2              =   6873.083
      Y1              =   1843.399
      Y2              =   1843.399
   End
   Begin VB.Line Line4 
      X1              =   7098.431
      X2              =   7098.431
      Y1              =   638.1
      Y2              =   1843.399
   End
   Begin VB.Line Line3 
      X1              =   2704.164
      X2              =   7098.431
      Y1              =   638.1
      Y2              =   638.1
   End
   Begin VB.Line Line2 
      X1              =   2704.164
      X2              =   6985.757
      Y1              =   1843.399
      Y2              =   1843.399
   End
   Begin VB.Line Line1 
      X1              =   2704.164
      X2              =   2704.164
      Y1              =   638.1
      Y2              =   1843.399
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "&Username :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "&Passoword :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   6525
      Left            =   120
      Picture         =   "frmLogin.frx":1084A
      Top             =   120
      Width           =   9840
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Rem validate input
    If IsNumeric(txtUserName.Text) = True Then
    MsgBox "Enter text data", vbInformation
    txtUserName.Text = ""
    txtUserName.SetFocus
    End If
       
    Rem check user name
    If Len(txtUserName.Text) = 0 Then
    MsgBox "Input user name", vbInformation
        txtUserName.SetFocus
    Exit Sub
    
    ElseIf txtUserName.Text = "sahil" Then
        If txtPassword.Text = "azim" Then
         Rem unload password form
            Load frmmain
            frmmain.Show
         Rem load navigation form
            Unload frmLogin
            frmLogin.Hide
    Else
        MsgBox "Incorrect password", vbCritical
        txtPassword.Text = ""
        txtPassword.SetFocus
        End If
    End If
    
End Sub

