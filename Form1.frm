VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmmedicine 
   BackColor       =   &H00400000&
   Caption         =   "medicine"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17565
   FillColor       =   &H00800000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   17565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "SAVE"
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
      Left            =   6360
      TabIndex        =   17
      Top             =   5880
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1680
      Top             =   6600
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
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
      RecordSource    =   "MEDICINE"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   8520
      TabIndex        =   12
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "DELETE"
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
      Left            =   4440
      TabIndex        =   11
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   5880
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   5880
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Medicine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   6735
      Begin VB.TextBox Text5 
         DataField       =   "PRICE"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox Text4 
         DataField       =   "QUANTITY"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         DataField       =   "DRUGNAME"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         DataField       =   "INDEXNO"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   2040
         TabIndex        =   13
         Top             =   480
         Width           =   3015
      End
      Begin VB.CommandButton Command3 
         Caption         =   "CLOSE"
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
         Left            =   4560
         TabIndex        =   7
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "NEW"
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
         Left            =   2520
         TabIndex        =   6
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "TOTAL"
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
         Left            =   600
         TabIndex        =   5
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         Caption         =   "PRICE"
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
         Left            =   480
         TabIndex        =   4
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         Caption         =   "QUANTITY"
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
         Left            =   360
         TabIndex        =   3
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Caption         =   "DRUG NAME"
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
         Left            =   360
         TabIndex        =   2
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   "INDEX NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Label Label5 
      Caption         =   "TOTAL"
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
      Left            =   8520
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "frmmedicine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Text1 = Val(Text4.Text) * (Text5.Text)
MsgBox "AMOUNT PAYABLE IS: " & Text1.Text + " TO THE CASHIER!!"
End Sub


Private Sub Command2_Click()
Adodc1.Recordset.AddNew

End Sub

Private Sub Command3_Click()
Unload Me
  
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.BOF = True Then
Adodc1.Recordset.MoveNext
End If
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command6_Click()
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

Private Sub Command7_Click()
Adodc1.Recordset.Update
MsgBox "SAVED TO THE DATABESE"
End Sub

