VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "HMS  Reports"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   5355
   ScaleWidth      =   8775
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   4455
      Left            =   2040
      TabIndex        =   35
      Top             =   960
      Width           =   6735
      Begin VB.ListBox List3 
         Height          =   3570
         Index           =   0
         ItemData        =   "Form4.frx":030A
         Left            =   120
         List            =   "Form4.frx":030C
         TabIndex        =   40
         Top             =   600
         Width           =   1455
      End
      Begin VB.ListBox List3 
         Height          =   3570
         Index           =   1
         ItemData        =   "Form4.frx":030E
         Left            =   1560
         List            =   "Form4.frx":0310
         TabIndex        =   39
         Top             =   600
         Width           =   1455
      End
      Begin VB.ListBox List3 
         Height          =   3570
         Index           =   2
         ItemData        =   "Form4.frx":0312
         Left            =   3000
         List            =   "Form4.frx":0314
         TabIndex        =   38
         Top             =   600
         Width           =   1215
      End
      Begin VB.ListBox List3 
         Height          =   3570
         Index           =   3
         ItemData        =   "Form4.frx":0316
         Left            =   4200
         List            =   "Form4.frx":0318
         TabIndex        =   37
         Top             =   600
         Width           =   1455
      End
      Begin VB.ListBox List3 
         Height          =   3570
         Index           =   4
         ItemData        =   "Form4.frx":031A
         Left            =   5640
         List            =   "Form4.frx":031C
         TabIndex        =   36
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Name"
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Address"
         Height          =   255
         Left            =   1920
         TabIndex        =   44
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "Contact No"
         Height          =   255
         Left            =   3120
         TabIndex        =   43
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Reservation Date"
         Height          =   255
         Left            =   4200
         TabIndex        =   42
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "Confirmed"
         Height          =   255
         Left            =   5760
         TabIndex        =   41
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "D:\ashik\HMS\hotel2.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "reservation"
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame3 
      Height          =   4455
      Left            =   2040
      TabIndex        =   22
      Top             =   960
      Width           =   6735
      Begin VB.ListBox List2 
         Height          =   3765
         Index           =   5
         ItemData        =   "Form4.frx":031E
         Left            =   5520
         List            =   "Form4.frx":0320
         TabIndex        =   28
         Top             =   480
         Width           =   1095
      End
      Begin VB.ListBox List2 
         Height          =   3765
         Index           =   4
         ItemData        =   "Form4.frx":0322
         Left            =   4680
         List            =   "Form4.frx":0324
         TabIndex        =   27
         Top             =   480
         Width           =   855
      End
      Begin VB.ListBox List2 
         Height          =   3765
         Index           =   3
         ItemData        =   "Form4.frx":0326
         Left            =   3360
         List            =   "Form4.frx":0328
         TabIndex        =   26
         Top             =   480
         Width           =   1335
      End
      Begin VB.ListBox List2 
         Height          =   3765
         Index           =   2
         ItemData        =   "Form4.frx":032A
         Left            =   2160
         List            =   "Form4.frx":032C
         TabIndex        =   25
         Top             =   480
         Width           =   1215
      End
      Begin VB.ListBox List2 
         Height          =   3765
         Index           =   1
         ItemData        =   "Form4.frx":032E
         Left            =   1560
         List            =   "Form4.frx":0330
         TabIndex        =   24
         Top             =   480
         Width           =   615
      End
      Begin VB.ListBox List2 
         Height          =   3765
         Index           =   0
         ItemData        =   "Form4.frx":0332
         Left            =   120
         List            =   "Form4.frx":0334
         TabIndex        =   23
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Amount"
         Height          =   255
         Left            =   5760
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Duration"
         Height          =   255
         Left            =   4800
         TabIndex        =   33
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Checkout date"
         Height          =   255
         Left            =   3480
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Checkin Date"
         Height          =   255
         Left            =   2280
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Room "
         Height          =   255
         Left            =   1560
         TabIndex        =   30
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Name"
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   2040
      TabIndex        =   9
      Top             =   960
      Width           =   6735
      Begin VB.ListBox List1 
         Height          =   3765
         Index           =   5
         ItemData        =   "Form4.frx":0336
         Left            =   2520
         List            =   "Form4.frx":0338
         TabIndex        =   20
         Top             =   480
         Width           =   615
      End
      Begin VB.ListBox List1 
         Height          =   3765
         Index           =   0
         ItemData        =   "Form4.frx":033A
         Left            =   120
         List            =   "Form4.frx":033C
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Height          =   3765
         Index           =   1
         ItemData        =   "Form4.frx":033E
         Left            =   1440
         List            =   "Form4.frx":0340
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Height          =   3765
         Index           =   2
         ItemData        =   "Form4.frx":0342
         Left            =   3120
         List            =   "Form4.frx":0344
         TabIndex        =   12
         Top             =   480
         Width           =   1575
      End
      Begin VB.ListBox List1 
         Height          =   3765
         Index           =   3
         ItemData        =   "Form4.frx":0346
         Left            =   4680
         List            =   "Form4.frx":0348
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Height          =   3765
         Index           =   4
         ItemData        =   "Form4.frx":034A
         Left            =   5760
         List            =   "Form4.frx":034C
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Age"
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Arival Date"
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Address"
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Phone"
         Height          =   255
         Left            =   4800
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Room No"
         Height          =   255
         Left            =   5760
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "D:\ashik\HMS\hotel2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "checkout"
         Top             =   4800
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "D:\ashik\HMS\hotel2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "checkin"
         Top             =   4800
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Print"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   4440
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Close"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Reservation Report"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CheckOut Report"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Checkin Report"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   2280
      ScaleHeight     =   555
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.Label Label1 
         Caption         =   "REPORT "
         BeginProperty Font 
            Name            =   "One Stroke Script LET"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   1
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Label Label8 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frame3.Visible = False
Frame2.Visible = True
Frame4.Visible = False
Frame2.Caption = "CheckIn Report"
Call listclear
Call checkinreport

   
End Sub

Private Sub Command2_Click()
Frame3.Visible = True
Frame2.Visible = False
Frame4.Visible = False
Frame3.Caption = "CheckOut Reports"
Call listclear1
Call checkoutreport
End Sub

Private Sub Command3_Click()
Frame4.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Frame4.Caption = "Reservation Report"
Call listclear2

Call reservreport
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Form4.PrintForm
End Sub

Private Sub Form_Load()
Label8.Caption = Format(Date, "long date")
Frame3.Visible = False
Frame4.Visible = False
End Sub

Private Sub listclear()
List1(0).Clear
List1(1).Clear
List1(2).Clear
List1(3).Clear
List1(4).Clear
List1(5).Clear
End Sub
Private Sub checkinreport()
Dim sql As String
Dim duration As Integer

sql = "select * from checkin"

With Data1
    .RecordSource = "select * from checkin"
    .Refresh
    With .Recordset
    .MoveFirst
    Do Until .EOF
    List1(0).AddItem .Fields(2)
    List1(1).AddItem .Fields(0)
    List1(2).AddItem .Fields(5)
    List1(3).AddItem .Fields(8)
    List1(4).AddItem .Fields(9)

    'duration = DateValue(Format(Now, "Short Date")) - .Fields(0)
    List1(5).AddItem .Fields(4)
    .MoveNext
    Loop
    
End With
End With
End Sub
Private Sub listclear1()
List2(0).Clear
List2(1).Clear
List2(2).Clear
List2(3).Clear
List2(4).Clear
List2(5).Clear
End Sub
Private Sub checkoutreport()

With Data2
    .RecordSource = "select * from checkout"
    .Refresh
    With .Recordset
    .MoveFirst
    Do Until .EOF
    List2(0).AddItem .Fields(0)
    List2(1).AddItem .Fields(3)
    List2(2).AddItem .Fields(4)
    List2(3).AddItem .Fields(5)
    List2(4).AddItem .Fields(6)

    'duration = DateValue(Format(Now, "Short Date")) - .Fields(0)
    List2(5).AddItem .Fields(7)
    .MoveNext
    Loop
    
End With
End With
End Sub

Private Sub listclear2()
List3(0).Clear
List3(1).Clear
List3(2).Clear
List3(3).Clear
List3(4).Clear
End Sub
Private Sub reservreport()

With Data3
    .RecordSource = "select * from reservation"
    .Refresh
    With .Recordset
    .MoveFirst
    Do Until .EOF
    List3(0).AddItem .Fields(1)
    List3(1).AddItem .Fields(2)
    List3(2).AddItem .Fields(3)
    List3(3).AddItem .Fields(4)
    List3(4).AddItem .Fields(5)
    .MoveNext
    Loop
    
End With
End With
End Sub

