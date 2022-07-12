VERSION 5.00
Begin VB.Form employeeattendance 
   Caption         =   "Employee Attendance"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16020
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9060
   ScaleWidth      =   16020
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Main"
      Height          =   10815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   20175
      Begin VB.Frame Frame2 
         BackColor       =   &H0000C0C0&
         Caption         =   "Employee Work Schedule"
         Height          =   8415
         Left            =   480
         TabIndex        =   10
         Top             =   600
         Width           =   19335
         Begin VB.PictureBox Picture2 
            Height          =   4935
            Left            =   -720
            Picture         =   "employeeattendance.frx":0000
            ScaleHeight     =   4875
            ScaleWidth      =   8355
            TabIndex        =   22
            Top             =   3600
            Width           =   8415
         End
         Begin VB.PictureBox Picture1 
            Height          =   6615
            Left            =   7800
            Picture         =   "employeeattendance.frx":ABB0
            ScaleHeight     =   6555
            ScaleWidth      =   7515
            TabIndex        =   21
            Top             =   1680
            Width           =   7575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "employeeattendance.frx":2868A
            Left            =   2280
            List            =   "employeeattendance.frx":28694
            TabIndex        =   20
            Text            =   "SHIFT"
            Top             =   3120
            Width           =   4815
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   2280
            TabIndex        =   19
            Top             =   1080
            Width           =   4815
         End
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   2280
            TabIndex        =   18
            Top             =   2280
            Width           =   4815
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   2280
            TabIndex        =   17
            Top             =   1680
            Width           =   4815
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   2280
            TabIndex        =   16
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Shift"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   15
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Date"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   14
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   " Name"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   13
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Employee Id"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Department"
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   11
            Top             =   1680
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0000C0C0&
         Caption         =   "Control Buttons"
         Height          =   1335
         Left            =   480
         TabIndex        =   1
         Top             =   9240
         Width           =   19335
         Begin VB.CommandButton cmd2 
            Height          =   495
            Index           =   0
            Left            =   3120
            Picture         =   "employeeattendance.frx":286A8
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd9 
            Height          =   495
            Index           =   1
            Left            =   4560
            Picture         =   "employeeattendance.frx":2DF5F
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd5 
            Height          =   495
            Index           =   2
            Left            =   8880
            Picture         =   "employeeattendance.frx":337E6
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd4 
            Height          =   495
            Index           =   3
            Left            =   7440
            Picture         =   "employeeattendance.frx":39070
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd3 
            Height          =   495
            Index           =   4
            Left            =   6000
            Picture         =   "employeeattendance.frx":3E91B
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd6 
            Height          =   495
            Index           =   5
            Left            =   10320
            Picture         =   "employeeattendance.frx":441A5
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd10 
            Height          =   495
            Index           =   6
            Left            =   13200
            Picture         =   "employeeattendance.frx":49A82
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd8 
            Height          =   495
            Index           =   8
            Left            =   11760
            Picture         =   "employeeattendance.frx":4F349
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   480
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "employeeattendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd10_Click(Index As Integer)

Text1.Text = ""
Text7.Text = ""
Text5.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub cmd2_Click(Index As Integer)
Unload employeeattendance
End Sub

Private Sub cmd3_Click(Index As Integer)

On Error GoTo errordesc
rk.MoveFirst
End Sub

Private Sub cmd4_Click(Index As Integer)


If (rk.EOF = True) Then
 rk.MoveFirst
Else
rk.MoveNext
End If
var = Text1.Text
 sqlcmd = "select * from doctor where doctor_id=" & Val(var)
 
 Set rs = New ADODB.Recordset
 rs.CursorType = adOpenDynamic
 rs.LockType = adLockOptimistic
 rs.Open sqlcmd, cn, , , adCmdText
 Set Text2.DataSource = rs
 Text2.DataField = "doct_name"
 Set Text3.DataSource = rs
 Text3.DataField = "dept_id"
End Sub

Private Sub cmd5_Click(Index As Integer)

On Error GoTo errordesc
rk.MoveLast
rk.Update
End Sub

Private Sub cmd6_Click(Index As Integer)

On Error GoTo errordesc
If (rk.BOF = True) Then
 rk.MoveLast
Else
 rk.MovePrevious
End Sub

Private Sub cmd8_Click(Index As Integer)
Dim aid As String
aid = InputBox("Enter the doctor id to search", "search", "id")
sqk = "select * from doctor where doctor_id = " & Val(aid)
Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open sqk, cn, , , adCmdText
Set Text1.DataSource = rk
Text1.DataField = "EMP_ID"
Set Text7.DataSource = rk
Text7.DataField = "DATES"
Set Text5.DataSource = rk
Text5.DataField = "SHIFT"



Set rs = New ADODB.Recordset
 rs.CursorType = adOpenDynamic
 rs.LockType = adLockOptimistic
 rs.Open sqlcmd, cn, , , adCmdText
 Set Text2.DataSource = rs
 Text2.DataField = "doc_name"
 Set Text3.DataSource = rs
 Text3.DataField = "dept_id"
End Sub

Private Sub cmd9_Click(Index As Integer)
rk.AddNew
End Sub

Private Sub Form_Load()
Dim var As String
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott"
cn.CursorLocation = adUseClient
cn.Open

Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open "schedule", cn, , , adCmdTable
MsgBox "loading please wait !!.............. ", , "Loading Message......."


Set Text1.DataSource = rk
Text1.DataField = "EMP_ID"
Set Text7.DataSource = rk
Text7.DataField = "DATES"
Set Combo1.DataSource = rk
Combo1.DataField = "SHIFT"
 
 var = Text1.Text
 sqlcmd = "select * from doctor where doctor_id=" & Val(var)
 
 Set rs = New ADODB.Recordset
 rs.CursorType = adOpenDynamic
 rs.LockType = adLockOptimistic
 rs.Open sqlcmd, cn, , , adCmdText
 Set Text2.DataSource = rs
 Text2.DataField = "doct_name"
 Set Text3.DataSource = rs
 Text3.DataField = "dept_id"
End Sub

