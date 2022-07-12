VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form beds 
   Caption         =   "Beddings"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15555
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   15555
   WindowState     =   2  'Maximized
   Begin VB.Frame roompatient 
      Caption         =   "Room Patient"
      Height          =   10815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   20175
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "beds.frx":0000
         Height          =   5895
         Left            =   9960
         TabIndex        =   9
         Top             =   2400
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   10398
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataMember      =   "patient"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "beds.frx":001F
         Height          =   6015
         Left            =   720
         TabIndex        =   5
         Top             =   2400
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   10610
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataMember      =   "ROOMS"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         Caption         =   "Navigation Buttons "
         Height          =   1815
         Left            =   720
         TabIndex        =   2
         Top             =   8640
         Width           =   18375
         Begin VB.CommandButton Command3 
            Caption         =   "refresh"
            Height          =   375
            Left            =   10200
            TabIndex        =   11
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Caption         =   "back"
            Height          =   375
            Left            =   8400
            TabIndex        =   10
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Allocate Room"
            Height          =   375
            Left            =   6480
            TabIndex        =   6
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Beds"
         Height          =   1575
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Width           =   18495
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   4920
            TabIndex        =   8
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "beds.frx":003E
            Left            =   11760
            List            =   "beds.frx":0048
            TabIndex        =   4
            Top             =   600
            Width           =   3735
         End
         Begin VB.Label Label2 
            Caption         =   "Bed Number"
            Height          =   375
            Left            =   1800
            TabIndex        =   7
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Select Bed Room Type"
            Height          =   375
            Left            =   9120
            TabIndex        =   3
            Top             =   600
            Width           =   2055
         End
      End
   End
End
Attribute VB_Name = "beds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    
        Set rs = New ADODB.Recordset
        rs.CursorType = adOpenDynamic
        rs.LockType = adLockOptimistic
        sqk = "select * from rooms"
        rs.Open sqk, cn, , , adCmdText
        
        rs.AddNew
       rs.Fields(0) = addmissionpatient.Text1.Text
        rs.Fields(1) = Text1.Text
        rs.Fields(2) = Combo1.Text
        rs.Fields(3) = 100
        rs.Fields(4) = Date
        rs.Update
        
        
       DataGrid1.Refresh
    

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
DataGrid1.Refresh
End Sub

Private Sub Form_Load()
Dim roomno As Integer
roomno = 1
sqk = "SELECT bed_no FROM (SELECT bed_no FROM rooms ORDER BY bed_no desc) WHERE ROWNUM <= 1 "
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott"
cn.CursorLocation = adUseClient
cn.Open
Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open sqk, cn, , , adCmdText
roomno = rk.Fields(0)
Text1.Text = roomno + 1

End Sub
