VERSION 5.00
Begin VB.Form frmmain 
   ClientHeight    =   7395
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15735
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7395
   ScaleWidth      =   15735
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   10500
      Left            =   0
      Picture         =   "frmmain.frx":1084A
      Top             =   -2520
      Width           =   15750
   End
   Begin VB.Menu MENUTASK 
      Caption         =   "&TASK"
      Begin VB.Menu MNUTRANSACTIONS 
         Caption         =   "&TRANSACTIONS"
         Shortcut        =   ^T
      End
      Begin VB.Menu MNUPAYMENT 
         Caption         =   "&PAYMENTS"
         Shortcut        =   ^P
      End
      Begin VB.Menu MNUMEDICINE 
         Caption         =   "&MEDICINE"
         Shortcut        =   ^M
      End
      Begin VB.Menu MNUREORDER 
         Caption         =   "&RE-ORDER"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu MNUADMIN 
      Caption         =   "&ADMIN"
      Begin VB.Menu MNUPRINT 
         Caption         =   "&PRINT"
         Shortcut        =   {F9}
      End
      Begin VB.Menu MNUEXIT 
         Caption         =   "&EXIT"
         Shortcut        =   {F4}
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub MNUEXIT_Click()
Unload Me
End Sub

Private Sub MNUMEDICINE_Click()
frmmedicine.Show

End Sub

Private Sub MNUPAYMENT_Click()
frmpayments.Show

End Sub

Private Sub MNUPRINT_Click()
frmreports.Show

End Sub

Private Sub MNUREORDER_Click()
frmreorder.Show

End Sub

Private Sub MNUTRANSACTIONS_Click()
Frntransaction.Show

End Sub
