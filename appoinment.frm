VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form appoinment 
   Caption         =   "Appointment"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Doctor Appointment"
      Height          =   10815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   20175
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "appoinment.frx":0000
         Height          =   2655
         Left            =   600
         TabIndex        =   7
         Top             =   5280
         Width           =   18735
         _ExtentX        =   33046
         _ExtentY        =   4683
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
         DataMember      =   "appointment"
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
         Caption         =   "Set Appointment"
         Height          =   4575
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   10095
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   1920
            TabIndex        =   16
            Top             =   3840
            Width           =   4335
         End
         Begin VB.TextBox Text4 
            Height          =   735
            Left            =   1920
            TabIndex        =   14
            Top             =   2640
            Width           =   4335
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1920
            TabIndex        =   10
            Top             =   1920
            Width           =   4335
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1920
            TabIndex        =   9
            Top             =   1320
            Width           =   4335
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   1920
            TabIndex        =   8
            Top             =   600
            Width           =   4335
         End
         Begin VB.Label Label5 
            Caption         =   "Prescription"
            Height          =   495
            Left            =   360
            TabIndex        =   15
            Top             =   3840
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Discription"
            Height          =   375
            Left            =   360
            TabIndex        =   13
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Date"
            Height          =   375
            Left            =   360
            TabIndex        =   6
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Doctor Id"
            Height          =   495
            Left            =   360
            TabIndex        =   5
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Patient Id"
            Height          =   375
            Left            =   360
            TabIndex        =   4
            Top             =   600
            Width           =   2175
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "appoinment.frx":001F
         Height          =   4215
         Left            =   11160
         TabIndex        =   2
         Top             =   600
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   7435
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
         DataMember      =   "doctor"
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
      Begin VB.Frame Frame3 
         Caption         =   "Controls Buttons"
         Height          =   1455
         Left            =   600
         TabIndex        =   1
         Top             =   8400
         Width           =   18855
         Begin VB.CommandButton Command2 
            Caption         =   "Delete Appointment"
            Height          =   375
            Left            =   10800
            TabIndex        =   12
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Set Appointment"
            Height          =   375
            Left            =   8760
            TabIndex        =   11
            Top             =   600
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "appoinment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
      
        Set rk = New ADODB.Recordset
        rk.CursorType = adOpenDynamic
        rk.LockType = adLockOptimistic
        rk.Open "appointment", cn, , , adCmdTable
        
        rk.AddNew
        rk.Fields(0) = Text1.Text
        rk.Fields(1) = Text2.Text
        rk.Fields(2) = Text3.Text
        rk.Fields(3) = Text5.Text
        rk.Update
        
        
        
        Unload Me
        
        
        

        
       
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.Text = addmissionpatient.Text1.Text
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott"
cn.CursorLocation = adUseClient
cn.Open

Set rk = New ADODB.Recordset
rk.CursorType = adOpenDynamic
rk.LockType = adLockOptimistic
rk.Open "appointment", cn, , , adCmdTable
MsgBox "loading please wait !!.............. ", , "Loading Message......."
Text3.Text = Date
Set Text2.DataSource = rk
        Text2.DataField = "doctor_id"
        
        Set Text5.DataSource = rk
        Text5.DataField = "prescription"




End Sub
