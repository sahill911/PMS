VERSION 5.00
Begin VB.Form frmreports 
   BackColor       =   &H00400000&
   Caption         =   "Reports"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12135
   Icon            =   "frmreports.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   12135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Choose the report to View or Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   9495
      Begin VB.CommandButton Command1 
         Caption         =   "Open the report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2040
         TabIndex        =   2
         Top             =   1560
         Width           =   4455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmreports.frx":1084A
         Left            =   1680
         List            =   "frmreports.frx":1085A
         TabIndex        =   1
         Top             =   960
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmreports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1.Text = "MEDICINE" Then
DataReport1.Show
ElseIf Combo1.Text = "PAYMENTS" Then
DataReport2.Show
ElseIf Combo1.Text = "RE-ORDER" Then
DataReport3.Show
ElseIf Combo1.Text = "TRANSACTION" Then
Datareportt4.Show
End If
End Sub

