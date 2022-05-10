VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frntransaction 
   BackColor       =   &H00400000&
   Caption         =   "Transactions of the drugs"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14115
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frntransaction.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   14115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      TabIndex        =   19
      Top             =   6000
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4080
      Top             =   6720
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\project\mytable.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\project\mytable.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Transactionone"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Update"
      Height          =   615
      Left            =   9960
      TabIndex        =   18
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Delete"
      Height          =   615
      Left            =   7320
      TabIndex        =   17
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Add New"
      Height          =   615
      Left            =   5040
      TabIndex        =   16
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Previous"
      Height          =   615
      Left            =   3000
      TabIndex        =   14
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Next"
      Height          =   615
      Left            =   1080
      TabIndex        =   13
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Transactions done here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4935
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   13215
      Begin VB.CommandButton Command3 
         Caption         =   "&Costing"
         Height          =   735
         Left            =   8400
         TabIndex        =   15
         ToolTipText     =   "Displays the cost of the drug"
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         DataField       =   "AmntPayable"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   4080
         Width           =   3135
      End
      Begin VB.TextBox Text5 
         DataField       =   "Cost"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   11
         Top             =   3360
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         DataField       =   "Quantity"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         DataField       =   "DrugName"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         DataField       =   "Mobile"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         DataField       =   "CustName"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   720
         Width           =   5415
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808080&
         Caption         =   "  Amount payable"
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
         Left            =   120
         TabIndex        =   6
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808080&
         Caption         =   "  Cost"
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
         Left            =   480
         TabIndex        =   5
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         Caption         =   "  Quantity"
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
         Left            =   360
         TabIndex        =   4
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         Caption         =   "  Drug Name"
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
         Left            =   360
         TabIndex        =   3
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         Caption         =   "  Contact No."
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
         Left            =   360
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   " Customer name"
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
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Frntransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.BOF = True Then
Adodc1.Recordset.MoveNext
End If
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command3_Click()
Text6 = Val(Text4.Text) * (Text5.Text)
MsgBox "COSTING IS: " & Text6.Text

End Sub

Private Sub Command4_Click()
Adodc1.Recordset.AddNew

End Sub

Private Sub Command5_Click()
Dim E As String
E = MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion, "Please Confirm!!")
If E = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Recoed Erased!!"
Else
MsgBox "Record Not Deleted!!"
If Adodc1.Recordset.BOF = True Then
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
Adodc1.Recordset.MovePrevious
End If
End If
End If
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.Update
MsgBox "SAVED TO THE DATABESE"
End Sub

Private Sub Command7_Click()
Unload Me

End Sub
