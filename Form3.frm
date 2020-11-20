VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "HMS STATISTICS "
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5250
   ScaleWidth      =   4230
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton Command2 
         Caption         =   "&Print"
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Close"
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Hotel Statistics"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label2 
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
         Left            =   480
         TabIndex        =   14
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "Total Guest Checkin Today"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Total Reservations Done"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Total Guest Checkout Today"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Total Reservations Confirmed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Occupied Rooms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Vacant Rooms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   255
         Left            =   2880
         TabIndex        =   7
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Label14"
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Label15"
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Label16"
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   3960
         Width           =   615
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "checkout"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "D:\ashik\HMS\hotel2.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "reservation"
      Top             =   4800
      Width           =   1065
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\ashik\HMS\hotel2.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "checkin"
      Top             =   4800
      Width           =   1065
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim chkin As Integer
Dim reserv As Integer
Dim occupied As Integer
Dim vacant As Integer
Dim chkout As Integer
Dim restoday As Integer

Private Sub Command1_Click()
Let chkin = 0
Let reserv = 0
Let occupied = 0
Let vacant = 0
Let chkout = 0
Let restoday = 0
Unload Me

End Sub

Private Sub Command2_Click()
Form3.PrintForm
End Sub

Private Sub Form_Load()
Label2.Caption = Format(Now, "Long Date")
Call roomoccupied
Label15.Caption = occupied
Label16.Caption = vacant
Call statchkin
Label10.Caption = chkin
Call statreserv
Label11.Caption = reserv
Call statchkout
Label12.Caption = chkout
Call statrestoday
Label14.Caption = restoday
End Sub

Private Sub roomoccupied()
Dim sql As String

sql = "Select * from room"

With Form2.Data1
   .RecordSource = sql
   .Refresh
   .Recordset.MoveFirst
    Do Until .Recordset.EOF
    On Error Resume Next
    If .Recordset.Fields(1) = True Then
        occupied = occupied + 1
        Else: vacant = vacant + 1     'List4(0).AddItem .Recordset("roomno")
        End If
        
        .Recordset.MoveNext
        Loop
        End With
End Sub



Private Sub statchkin()
Dim sql As String

sql = "Select * from checkin"

With Form2.Data1
   .RecordSource = sql
   .Refresh
   .Recordset.MoveFirst
    Do Until .Recordset.EOF
    On Error Resume Next
    If .Recordset.Fields(0) = Date Then
        chkin = chkin + 1
        End If
        
        .Recordset.MoveNext
        Loop
        End With
End Sub
Private Sub statreserv()
Dim sql As String

sql = "Select * from reservation"

With Data2
   .RecordSource = sql
   .Refresh
   .Recordset.MoveFirst
    Do Until .Recordset.EOF
    On Error Resume Next
    If .Recordset.Fields(0) = Date Then
        reserv = reserv + 1
        End If
        
        .Recordset.MoveNext
        Loop
        End With
End Sub
Private Sub statchkout()
Dim sql As String

sql = "Select * from checkout"

With Data3
   .RecordSource = sql
   .Refresh
   .Recordset.MoveFirst
    Do Until .Recordset.EOF
    On Error Resume Next
    If .Recordset.Fields(5) = Date Then
        chkout = chkout + 1
        End If
        
        .Recordset.MoveNext
        Loop
        End With
End Sub
Private Sub statrestoday()
Dim sql As String

sql = "Select * from reservation"

With Data2
   .RecordSource = sql
   .Refresh
   .Recordset.MoveFirst
    Do Until .Recordset.EOF
    
    On Error Resume Next
      If .Recordset.Fields(5) = True Then
        restoday = restoday + 1
        .Recordset.MoveNext
        Exit Sub
        End If
        Loop
        End With
End Sub

