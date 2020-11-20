VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Administrator Login..."
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   ScaleHeight     =   5355
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "New user registration"
      Height          =   2535
      Left            =   120
      TabIndex        =   21
      Top             =   480
      Width           =   4455
      Begin VB.CommandButton Command10 
         Caption         =   "Close"
         Height          =   375
         Left            =   3120
         TabIndex        =   34
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Clear"
         Height          =   375
         Left            =   3120
         TabIndex        =   33
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Add"
         Height          =   375
         Left            =   3120
         TabIndex        =   32
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text9 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   31
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   960
         TabIndex        =   30
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   3120
         TabIndex        =   29
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   960
         TabIndex        =   28
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   960
         TabIndex        =   27
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "UserId"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Phone"
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Address"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "D:\ashik\HMS\hotel2.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "admin"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   4455
      Begin VB.CommandButton Command12 
         Caption         =   "Update"
         Height          =   375
         Left            =   3240
         TabIndex        =   36
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Edit"
         Height          =   375
         Left            =   3240
         TabIndex        =   35
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Admin Password"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Admin UserId"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Administrator ID Management"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\ashik\HMS\hotel2.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "admin"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4455
      Begin VB.CommandButton Command6 
         Caption         =   "Update"
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Clear"
         Height          =   375
         Left            =   3360
         TabIndex        =   12
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Delete User"
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Edit User"
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add User"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "User List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Main"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "ADMINISTRATOR AREA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer



Private Sub Combo1_Click()
With Data1
.Refresh
.Recordset.MoveFirst
Do Until .Recordset.EOF
If .Recordset.Fields(0) = Combo1.Text Then
Text1.Text = .Recordset.Fields(3)
Text2.Text = .Recordset.Fields(4)
Exit Sub
Else
.Recordset.MoveNext
End If
Loop
End With
End Sub

Private Sub Command1_Click()
Unload Me
Form2.Show
Form2.Caption = "Welcome Administrator..."
End Sub

Private Sub Command10_Click()
Frame3.Visible = False
Frame1.Visible = True

End Sub

Private Sub Command11_Click()
Data2.RecordSource = "select * from admin"
Data2.Recordset.MoveFirst
Text3.Text = Data2.Recordset.Fields(0)
Text4.Text = Data2.Recordset.Fields(1)
Text4.SetFocus
Text3.Enabled = False
MsgBox "please edit the password", vbInformation, "HMS"
End Sub

Private Sub Command12_Click()
With Data2
.RecordSource = "select * from admin"
.Refresh
.Recordset.Edit
.Recordset.Fields(0) = Text3.Text
.Recordset.Fields(1) = Text4.Text
.Recordset.Update
End With
MsgBox "Administrator password has been changed", vbInformation, "HMS"
Command1.SetFocus
End Sub

Private Sub Command2_Click()
Frame3.Visible = True
Frame1.Visible = False
Text5.SetFocus
End Sub

Private Sub Command3_Click()
Text1.Enabled = True
Text2.Enabled = True
Command2.Enabled = False
Command4.Visible = False
Command6.Visible = True
Command3.Enabled = False


End Sub

Private Sub Command4_Click()
MsgBox "Are u sure want to delete user", vbYesNoCancel, "HMS"
Data1.Recordset.Delete
End Sub

Private Sub Command5_Click()
Text1 = ""
Text2 = ""
Command2.Enabled = True
Command3.Enabled = True
Command6.Visible = False
Command4.Visible = True
Command4.Enabled = True
End Sub

Private Sub Command6_Click()
Data1.Recordset.Edit
Data1.Recordset.Fields(3) = Text1.Text
Data1.Recordset.Fields(4) = Text2.Text
Data1.Recordset.Update
MsgBox "User Id or password updated", vbInformation, "HMS"
Command6.Visible = False
Command4.Visible = True
End Sub

Private Sub Command7_Click()
End
End Sub

Private Sub Command8_Click()
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Text5.Text
Data1.Recordset.Fields(1) = Text6.Text
Data1.Recordset.Fields(2) = Text7.Text
Data1.Recordset.Fields(3) = Text8.Text
Data1.Recordset.Fields(4) = Text9.Text
Data1.Recordset.Update
MsgBox "New User added", vbInformation, "HMS"
Command8.enable = False
Command10.SetFocus
End Sub

Private Sub Command9_Click()
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text5.SetFocus
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\hotel2.mdb"
Data1.RecordSource = "select * from user"

Data2.DatabaseName = App.Path & "\hotel2.mdb"
Data2.RecordSource = "select * from admin"
'Data2.DatabaseName = App.Path & "\hotel2.mdb"
'Data2.RecordSource = "select * from admin"

With Data1
.Refresh
.Recordset.MoveFirst
Do Until .Recordset.EOF
For i = 0 To 3 'len(.Recordset.BatchSize)
Combo1.List(i) = .Recordset.Fields(0)
.Recordset.MoveNext
Next i
Loop
End With

Text1.Enabled = False
Text2.Enabled = False
Command6.Visible = False
'Command8.Visible = False
Frame3.Visible = False
'Call admindetails

End Sub

Private Sub admindetails()
Text3.Text = Data2.Recordset.Fields(0)
Text4.Text = Data2.Recordset.Fields(1)
Exit Sub
End Sub
