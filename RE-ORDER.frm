VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmreorder 
   BackColor       =   &H00400000&
   Caption         =   "Re-Oreder"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13125
   Icon            =   "RE-ORDER.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   13125
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4080
      Top             =   7440
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
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
      RecordSource    =   "Reorder"
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
   Begin VB.CommandButton Command6 
      Caption         =   "&Close"
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
      Left            =   10440
      TabIndex        =   20
      Top             =   6840
      Width           =   1935
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
      Left            =   9120
      TabIndex        =   19
      Top             =   6840
      Width           =   855
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
      Left            =   7920
      TabIndex        =   18
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
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
      Left            =   5040
      TabIndex        =   17
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Add New"
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
      Left            =   3000
      TabIndex        =   16
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Deficit"
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
      Left            =   1080
      TabIndex        =   15
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Re-Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   11175
      Begin VB.TextBox Text8 
         Height          =   735
         Left            =   8400
         TabIndex        =   21
         Top             =   4320
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         DataField       =   "PurchasedQty"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         DataField       =   "CurrentQty"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2520
         TabIndex        =   13
         Top             =   3960
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         DataField       =   "Drug name"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         Top             =   3120
         Width           =   4455
      End
      Begin VB.TextBox Text4 
         DataField       =   "Supplier name"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   2280
         Width           =   4455
      End
      Begin VB.TextBox Text3 
         DataField       =   "Invoice no"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   2520
         TabIndex        =   10
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         DataField       =   "Purchase no"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         DataField       =   "Index no"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   9360
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808080&
         Caption         =   "Index No"
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
         Left            =   8160
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808080&
         Caption         =   " Purchese QTY"
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
         TabIndex        =   6
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         Caption         =   " Currnet QTY"
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
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         Caption         =   " Drug Name"
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
         TabIndex        =   4
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         Caption         =   " Suplier Name"
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
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Caption         =   " Invoice No"
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
         TabIndex        =   2
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   " Purchase No"
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
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmreorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text8.Text = Val(Text7.Text) - (Text6.Text)
MsgBox "DEFICIT IS: " & Text8.Text

End Sub

Private Sub Command2_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.Update

End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MovePrevious

End Sub

Private Sub Command6_Click()
Unload Me
End Sub
