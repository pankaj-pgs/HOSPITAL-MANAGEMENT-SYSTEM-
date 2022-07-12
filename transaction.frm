VERSION 5.00
Begin VB.Form transaction 
   BackColor       =   &H00000000&
   Caption         =   "Transaction"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15990
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   15990
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H0000C0C0&
      Caption         =   "Details"
      Height          =   4455
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   20055
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   8160
         TabIndex        =   29
         Top             =   480
         Width           =   3015
      End
      Begin VB.PictureBox Picture2 
         Height          =   4335
         Left            =   13440
         Picture         =   "transaction.frx":0000
         ScaleHeight     =   4275
         ScaleWidth      =   6435
         TabIndex        =   27
         Top             =   120
         Width           =   6495
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   2160
         TabIndex        =   23
         Top             =   3000
         Width           =   2895
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   2160
         TabIndex        =   21
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   2160
         TabIndex        =   20
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   2160
         TabIndex        =   19
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Total Amount "
         Height          =   375
         Left            =   6000
         TabIndex        =   28
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Bill No"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Patient Name"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   17
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Patient Id"
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Dept Name"
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Contact Number"
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H0000C0C0&
      Caption         =   "Control  Buttons"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   9120
      Width           =   20055
      Begin VB.CommandButton Command1 
         Height          =   495
         Index           =   2
         Left            =   7680
         Picture         =   "transaction.frx":2B976
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command9 
         Height          =   495
         Index           =   1
         Left            =   10920
         Picture         =   "transaction.frx":311F1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Height          =   495
         Index           =   0
         Left            =   9240
         Picture         =   "transaction.frx":36AA8
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Transaction Details"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   20055
      Begin VB.PictureBox Picture1 
         Height          =   3375
         Left            =   9960
         Picture         =   "transaction.frx":3C323
         ScaleHeight     =   3315
         ScaleWidth      =   9915
         TabIndex        =   26
         Top             =   360
         Width           =   9975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "transaction.frx":5639D
         Left            =   2160
         List            =   "transaction.frx":563AA
         TabIndex        =   25
         Text            =   "cash"
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Mode"
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Medicine Amount"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Diagnosis Fee"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Doctor Fee"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
   End
End
Attribute VB_Name = "transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub cmd2_Click(Index As Integer)
Unload Me


End Sub






Private Sub Command1_Click(Index As Integer)



Set rs = New ADODB.Recordset
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Open "bill", cn, , , adCmdTable
rs.AddNew
        rs.Fields(0) = Text10.Text
        rs.Fields(1) = Text2.Text
        rs.Fields(2) = Text5.Text
        rs.Fields(3) = Text4.Text
        rs.Fields(4) = Text3.Text
        rs.Fields(5) = Date
        rs.Fields(6) = Combo1.Text
        rs.Update
       Me.Hide
       
MsgBox "transaction complete"
End Sub

Private Sub Command9_Click(Index As Integer)
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = Now()
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott"
cn.CursorLocation = adUseClient
cn.Open
sqlcmd = " select e.pat_name,e.phn_no,e.doc_fee,d.dept_name from patient e,department d Where e.dept_id=d.dept_id  and  e.patient_id = " & dischargpatient.Text1.Text
Set rk = New ADODB.Recordset

rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open sqlcmd, cn, , , adCmdText
If (rk.EOF = True) Then
MsgBox "error"
End If
MsgBox "loading please wait !!.............. ", , "Loading Message......."
Text2.Text = dischargpatient.Text1.Text
Set Text6.DataSource = rk
Text6.DataField = "dept_name"
Set Text7.DataSource = rk
Text7.DataField = "pat_name"
Set Text9.DataSource = rk
Text9.DataField = "phn_no"
Set Text3.DataSource = rk
Text3.DataField = "doc_fee"



End Sub





Private Sub Text8_GotFocus()
Dim amount As Integer
amount = Text3.Text + Text4.Text + Text5.Text
Text8.Text = amount
End Sub
