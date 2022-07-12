VERSION 5.00
Begin VB.Form addmissionpatient 
   Caption         =   "Addmission  Patient"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15900
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Main"
      Height          =   10815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   20175
      Begin VB.Frame Frame3 
         BackColor       =   &H0000C0C0&
         Caption         =   "CONTROL BUTTONS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   600
         TabIndex        =   22
         Top             =   9240
         Width           =   18975
         Begin VB.CommandButton Command2 
            Caption         =   "allotbed"
            Height          =   495
            Left            =   1440
            TabIndex        =   33
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmd8 
            Height          =   495
            Index           =   8
            Left            =   11640
            Picture         =   "frm4.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd10 
            Height          =   495
            Index           =   6
            Left            =   13080
            Picture         =   "frm4.frx":587C
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd6 
            Height          =   495
            Index           =   5
            Left            =   10200
            Picture         =   "frm4.frx":B143
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd3 
            Height          =   495
            Index           =   4
            Left            =   6000
            Picture         =   "frm4.frx":10A20
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd4 
            Height          =   495
            Index           =   3
            Left            =   7320
            Picture         =   "frm4.frx":162AA
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd5 
            Height          =   495
            Index           =   2
            Left            =   8760
            Picture         =   "frm4.frx":1BB55
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmd2 
            Height          =   495
            Index           =   1
            Left            =   4680
            Picture         =   "frm4.frx":213DF
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmd9 
            Height          =   495
            Index           =   0
            Left            =   14520
            Picture         =   "frm4.frx":26C66
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Left            =   3240
            Picture         =   "frm4.frx":2C51D
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0000C0C0&
         Caption         =   "PATIENT INFORMATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8415
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Width           =   18975
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "frm4.frx":31DC8
            Left            =   2280
            List            =   "frm4.frx":31DD2
            TabIndex        =   32
            Top             =   1680
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   2280
            TabIndex        =   11
            Top             =   480
            Width           =   2295
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   2280
            TabIndex        =   10
            Top             =   1080
            Width           =   4575
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   2280
            TabIndex        =   9
            Top             =   2280
            Width           =   4575
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   2280
            TabIndex        =   8
            Top             =   4320
            Width           =   4575
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   2280
            TabIndex        =   7
            Top             =   4920
            Width           =   4575
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   2280
            TabIndex        =   6
            Top             =   5520
            Width           =   4575
         End
         Begin VB.TextBox Text8 
            Height          =   375
            Left            =   2280
            TabIndex        =   5
            Top             =   3720
            Width           =   4575
         End
         Begin VB.TextBox Text9 
            Height          =   375
            Left            =   2280
            TabIndex        =   4
            Top             =   3000
            Width           =   4575
         End
         Begin VB.PictureBox Picture1 
            Height          =   6615
            Left            =   7200
            Picture         =   "frm4.frx":31DE4
            ScaleHeight     =   6555
            ScaleWidth      =   7515
            TabIndex        =   3
            Top             =   240
            Width           =   7575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frm4.frx":739FD
            Left            =   2280
            List            =   "frm4.frx":73A0D
            TabIndex        =   2
            Top             =   6120
            Width           =   4575
         End
         Begin VB.Label Label1 
            Caption         =   "Gender"
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   21
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Patient Id"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Patient Name"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   19
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Patient Disease"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   18
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Date of Birth"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   17
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "PatientDepartment"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   16
            Top             =   3720
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Doctor Fee"
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   15
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Fathers Name"
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   14
            Top             =   4920
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Phone Number"
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   13
            Top             =   5520
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Blood Group"
            Height          =   375
            Index           =   9
            Left            =   240
            TabIndex        =   12
            Top             =   6120
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "addmissionpatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim flag As Integer

Private Sub patient()
Set rs = New ADODB.Recordset
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Open "patient", cn, , , adCmdTable
Set rk = rs
Set Text1.DataSource = rk
Text1.DataField = "patient_id"
Set Text2.DataSource = rk
Text2.DataField = "pat_name"
Set Text3.DataSource = rk
Text3.DataField = "sex"
Set Text4.DataSource = rk
Text4.DataField = "doc_fee"
Set Text5.DataSource = rk
Text5.DataField = "fathers_name"
Set Text6.DataSource = rk
Text6.DataField = "phn_no"

Set Text8.DataSource = rk
Text8.DataField = "dept_id"
Set Text9.DataSource = rk
Text9.DataField = "d_o_b"
Set Text10.DataSource = rk
Text10.DataField = "disease"
End Sub

Private Sub cmd10_Click(Index As Integer)
         Dim patientno As Integer
patientno = 1
sqk = "SELECT patient_id FROM (SELECT patient_id FROM patient ORDER BY patient_id desc) WHERE ROWNUM <= 1 "

Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open sqk, cn, , , adCmdText
patientno = rk.Fields(0)
Text1.Text = patientno + 1

         
 rk.Close
         
         Text2.Text = " "
         Text3.Text = " "
         
         Text8.Text = " "
         Text9.Text = " "
         Combo2.Text = " "
         Text6.Text = " "
         Text5.Text = " "
         Text4.Text = " "
End Sub


Private Sub cmd2_Click(Index As Integer)


        Set rk = New ADODB.Recordset
        rk.CursorType = adOpenDynamic
        rk.LockType = adLockOptimistic
        sqk = "select * from patient "
        rk.Open sqk, cn, , , adCmdText
        rk.AddNew
        rk.Fields(0) = Text1.Text
        rk.Fields(1) = Text2.Text
        rk.Fields(2) = Text9.Text
        rk.Fields(3) = Text4.Text
        rk.Fields(4) = Text5.Text
        rk.Fields(5) = Text6.Text
        rk.Fields(6) = Combo1.Text
        rk.Fields(7) = Text8.Text
        rk.Fields(8) = Combo2.Text
        rk.Fields(9) = Text3.Text
       

    
        Set Text1.DataSource = rk
Text1.DataField = "patient_id"
Set Text2.DataSource = rk
Text2.DataField = "pat_name"
Set Combo2.DataSource = rk
Combo2.DataField = "sex"
Set Text4.DataSource = rk
Text4.DataField = "doc_fee"
Set Text5.DataSource = rk
Text5.DataField = "fathers_name"
Set Text6.DataSource = rk
Text6.DataField = "phn_no"
Set Combo1.DataSource = rk
Combo1.DataField = "bld_grp"
Set Combo2.DataSource = rk
Combo2.DataField = "dept_id"
Set Text9.DataSource = rk
Text9.DataField = "d_o_b"
Set Text3.DataSource = rk
Text3.DataField = "disease"

        
       
        appoinment.Show
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call patient

End Sub

Private Sub cmd3_Click(Index As Integer)
On Error GoTo errordesc
rk.MoveFirst
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call patient
End Sub

Private Sub cmd4_Click(Index As Integer)
On Error GoTo errordesc
If rk.EOF = True Then
    rk.MoveFirst
    
End If
rk.MoveNext
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call patient
End Sub

Private Sub cmd5_Click(Index As Integer)
On Error GoTo errordesc
rk.MoveLast
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call patient
End Sub

Private Sub cmd6_Click(Index As Integer)
On Error GoTo errordesc
If (rk.BOF = True) Then
  rk.MoveLast
Else
 rk.MovePrevious
End If
Exit Sub
errordesc:
i = MsgBox(Err.Description, vbCritical)
Call patient
End Sub

Private Sub cmd8_Click(Index As Integer)
Dim aid As Integer
aid = InputBox("Enter the patientid to search", "search", "id")
sqk = "select * from patient where patient_id= " & Val(aid)
Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open sqk, cn, , , adCmdText
If rk.EOF = True Then
MsgBox "id not exist!........... "
Call patient
End If
Set Text1.DataSource = rk
Text1.DataField = "patient_id"
Set Text2.DataSource = rk
Text2.DataField = "pat_name"
Set Combo2.DataSource = rk
Combo2.DataField = "sex"
Set Text4.DataSource = rk
Text4.DataField = "doc_fee"
Set Text5.DataSource = rk
Text5.DataField = "fathers_name"
Set Text6.DataSource = rk
Text6.DataField = "phn_no"
Set Text7.DataSource = rk
Text7.DataField = "bld_grp"
Set Text8.DataSource = rk
Text8.DataField = "dept_id"
Set Text9.DataSource = rk
Text9.DataField = "d_o_b"
Set Text3.DataSource = rk
Text3.DataField = "disease"

End Sub

Private Sub cmd9_Click(Index As Integer)
Unload Me
End Sub

Private Sub Command1_Click()


rk.Update
MsgBox " UPDATED DATA COMPLETED!!"
rk.MoveFirst
End Sub

Private Sub Command2_Click()
beds.Show
End Sub

Private Sub Form_Load()
Dim z As Integer

Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott"
cn.CursorLocation = adUseClient
cn.Open

Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open "patient", cn, , , adCmdTable
MsgBox "loading please wait !!.............. ", , "Loading Message......."


Set Text1.DataSource = rk
Text1.DataField = "patient_id"
Set Text2.DataSource = rk
Text2.DataField = "pat_name"
Set Combo2.DataSource = rk
Combo2.DataField = "sex"
Set Text4.DataSource = rk
Text4.DataField = "doc_fee"
Set Text5.DataSource = rk
Text5.DataField = "fathers_name"
Set Text6.DataSource = rk
Text6.DataField = "phn_no"
Set Combo1.DataSource = rk
Combo1.DataField = "bld_grp"
Set Text8.DataSource = rk
Text8.DataField = "dept_id"
Set Text9.DataSource = rk
Text9.DataField = "d_o_b"
Set Text3.DataSource = rk
Text3.DataField = "disease"

rk.MoveFirst

flag = 1

End Sub


Private Sub Text1_LostFocus()
Dim aid As Integer
Dim z As Integer
aid = Text1.Text
sqk = "select * from patient where patient_id= " & Val(aid)
Set rs = New ADODB.Recordset
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Open sqk, cn, , , adCmdText
If (rs.EOF = True) Then
Else
z = rs.Fields(0)
End If
    If (z = aid) Then
        MsgBox "Enter Different Id, this id already exist!!........"
        Text1.Text = ""
        Text1.SetFocus
    End If

End Sub
