VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room Status -- HMS"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6405
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   4455
      Begin VB.CommandButton Command2 
         Caption         =   "&Close"
         Height          =   375
         Left            =   3000
         TabIndex        =   22
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "D:\ashik\HMS\hotel2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   -120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "checkin"
         Top             =   4680
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   375
         Left            =   3000
         TabIndex        =   20
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label17 
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label16 
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label15 
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label14 
         Height          =   255
         Left            =   2040
         TabIndex        =   16
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label13 
         Height          =   255
         Left            =   2040
         TabIndex        =   15
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label12 
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label11 
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label10 
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Duration"
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
         Left            =   360
         TabIndex        =   11
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Arrival Date"
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
         Left            =   360
         TabIndex        =   10
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Phone"
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
         Left            =   360
         TabIndex        =   9
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Address"
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
         Left            =   360
         TabIndex        =   8
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Sex"
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
         Left            =   360
         TabIndex        =   7
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Age"
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
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
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
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Room Number"
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
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1575
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "D:\ashik\HMS\hotel2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "room"
         Top             =   4680
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.ListBox List1 
         Height          =   3765
         ItemData        =   "Form5.frx":030A
         Left            =   240
         List            =   "Form5.frx":0347
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Room Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label18 
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
      Left            =   3240
      TabIndex        =   21
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase(App.Path + "/hotel2.mdb")
Set rs = db.OpenRecordset("checkin")
Label18.Caption = Format(Now, "long date")

End Sub

Private Sub List1_Click()
With Data1
    .RecordSource = "select * from room"
    .Refresh
    .Recordset.MoveFirst
    Do Until .Recordset.EOF
    If .Recordset.Fields(0) = List1.Text And .Recordset.Fields(1) = False Then
    Label10.Caption = .Recordset.Fields(0)
    Label11.Caption = ""
    Label12.Caption = ""
    Label13.Caption = ""
    Label14.Caption = ""
    Label15.Caption = ""
    Label16.Caption = ""
    Label17.Caption = ""
    Exit Sub
    Else
    'Call roomstatus
    .Recordset.MoveNext
    End If
    Loop
    
   Call roomstatus
   
   End With
   
End Sub

Private Sub roomstatus()
With Data2
    .RecordSource = "select * from checkin"
    .Refresh
    .Recordset.MoveFirst
    Do Until .Recordset.EOF
    Label10.Caption = ""
    Label11.Caption = ""
    Label12.Caption = ""
    Label13.Caption = ""
    Label14.Caption = ""
    Label15.Caption = ""
    Label16.Caption = ""
    Label17.Caption = ""
    If .Recordset.Fields(9) = List1.Text Then
    Label10.Caption = .Recordset.Fields(9)
    Label11.Caption = .Recordset.Fields(2)
    Label12.Caption = .Recordset.Fields(4)
    Label13.Caption = .Recordset.Fields(3)
    Label14.Caption = .Recordset.Fields(5)
    Label15.Caption = .Recordset.Fields(6)
    Label16.Caption = .Recordset.Fields(0)
    Label17.Caption = DateValue(Format(Now, "Short Date")) - .Recordset.Fields(0)
    
    Exit Sub
    Else
    .Recordset.MoveNext
    
    End If
    
    Loop
    End With
    
End Sub
