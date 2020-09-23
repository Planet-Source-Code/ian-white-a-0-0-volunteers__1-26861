VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   -195
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Gilroy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Gilroy.frx":1002
   ScaleHeight     =   8985
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   5160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   78
      Text            =   "Gilroy.frx":B0CC4
      Top             =   4320
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox txtYear 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10320
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      TabIndex        =   11
      Top             =   3360
      Width           =   1335
   End
   Begin VB.ListBox lstID 
      Height          =   300
      Left            =   720
      TabIndex        =   76
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lst 
      Height          =   300
      Left            =   0
      TabIndex        =   75
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtFirstName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      TabIndex        =   3
      Top             =   3720
      Width           =   6135
   End
   Begin VB.CommandButton cmdNoOfVolunteers 
      Caption         =   "No. Of Volunteers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   73
      Top             =   3600
      Width           =   1575
   End
   Begin VB.OptionButton Opt2001Vol 
      BackColor       =   &H00400000&
      Caption         =   "2001"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   72
      Top             =   3600
      Width           =   735
   End
   Begin VB.OptionButton OptAllVol 
      BackColor       =   &H00400000&
      Caption         =   "All Volunteers"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   71
      Top             =   3240
      Width           =   1575
   End
   Begin VB.OptionButton Opt2000Vol 
      BackColor       =   &H00400000&
      Caption         =   "2000"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   70
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00400000&
      Height          =   735
      Left            =   10200
      Picture         =   "Gilroy.frx":B0D5A
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Height          =   735
      Left            =   8640
      Picture         =   "Gilroy.frx":B42DC
      Style           =   1  'Graphical
      TabIndex        =   68
      ToolTipText     =   "Goes To Previous Record In Database."
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtPassChange 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'DISABLE
      Index           =   11
      Left            =   7560
      PasswordChar    =   "*"
      TabIndex        =   53
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdPassChange 
      Caption         =   "CHANGE"
      Height          =   255
      Index           =   21
      Left            =   10560
      TabIndex        =   67
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdPassChange 
      Caption         =   "CHANGE"
      Height          =   255
      Index           =   20
      Left            =   7560
      TabIndex        =   55
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPassChange 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   15
      Left            =   10560
      PasswordChar    =   "*"
      TabIndex        =   61
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPassChange 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'DISABLE
      Index           =   14
      Left            =   10560
      PasswordChar    =   "*"
      TabIndex        =   60
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPassChange 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'DISABLE
      Index           =   13
      Left            =   10560
      PasswordChar    =   "*"
      TabIndex        =   59
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton ChkOptionTeach 
      BackColor       =   &H00400000&
      Caption         =   "Teacher"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   47
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtPassChange 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   12
      Left            =   7560
      PasswordChar    =   "*"
      TabIndex        =   54
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPassChange 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'DISABLE
      Index           =   10
      Left            =   7560
      PasswordChar    =   "*"
      TabIndex        =   52
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton ChkOptionAdmin 
      BackColor       =   &H00400000&
      Caption         =   "Admin"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   48
      Top             =   1800
      Width           =   975
   End
   Begin VB.OptionButton ChkOptionUser 
      BackColor       =   &H00400000&
      Caption         =   "User"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   46
      Top             =   1080
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Tuesday:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   22
      Left            =   10080
      TabIndex        =   26
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Monday:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   21
      Left            =   10200
      TabIndex        =   25
      Top             =   6130
      Width           =   1455
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Wednesday:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   23
      Left            =   9600
      TabIndex        =   27
      Top             =   7320
      Width           =   2055
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Thursday:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Index           =   24
      Left            =   10080
      TabIndex        =   28
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Friday:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   25
      Left            =   10440
      TabIndex        =   29
      Top             =   8370
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Advanced Search"
      Height          =   375
      Left            =   5160
      TabIndex        =   45
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Uniforms Shop:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   3
      Left            =   3960
      TabIndex        =   14
      Top             =   6840
      Width           =   2415
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Major Raffle:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   4
      Left            =   4320
      TabIndex        =   15
      Top             =   7320
      Width           =   2055
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Library Work:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   5
      Left            =   4200
      TabIndex        =   16
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Careers Night:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   6
      Left            =   4080
      TabIndex        =   17
      Top             =   8280
      Width           =   2295
   End
   Begin VB.ListBox lstNames 
      BackColor       =   &H00400000&
      ForeColor       =   &H0000FFFF&
      Height          =   4860
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "General Help:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   17
      Left            =   6840
      TabIndex        =   24
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Promotion:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   16
      Left            =   7200
      TabIndex        =   23
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Sales:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   15
      Left            =   7800
      TabIndex        =   22
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Setting Up:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   14
      Left            =   7200
      TabIndex        =   21
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Art:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   13
      Left            =   8160
      TabIndex        =   20
      Top             =   6960
      Width           =   855
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Catering:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   11
      Left            =   7440
      TabIndex        =   18
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Craft:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   12
      Left            =   7920
      TabIndex        =   19
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Ball:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   2
      Left            =   5400
      TabIndex        =   13
      Top             =   6240
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   $"Gilroy.frx":B779E
      OLEDBString     =   $"Gilroy.frx":B782E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Social Commitee:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   12
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00289CFE&
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtHomeRoom 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      TabIndex        =   9
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   3615
   End
   Begin VB.TextBox txtWorkNo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8880
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      TabIndex        =   5
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox txtStudName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      TabIndex        =   8
      Top             =   4800
      Width           =   6135
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7800
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      TabIndex        =   10
      Top             =   5160
      Width           =   3855
   End
   Begin VB.TextBox txtMobile 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8880
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      TabIndex        =   7
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox txtPhone 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      TabIndex        =   4
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox txtCareer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      TabIndex        =   6
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox txtSurName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      TabIndex        =   2
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   9735
      TabIndex        =   77
      Top             =   3360
      Width           =   600
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name(s):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3720
      TabIndex        =   74
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label LabelPassChange 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Teacher Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   9000
      TabIndex        =   66
      Top             =   720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label LabelPassChange 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Admin Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   65
      Top             =   720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label LabelPassChange 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   9240
      TabIndex        =   64
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label LabelPassChange 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   9000
      TabIndex        =   63
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label LabelPassChange 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Original:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   9240
      TabIndex        =   62
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label LabelPassChange 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   6240
      TabIndex        =   58
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label LabelPassChange 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   6000
      TabIndex        =   57
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label LabelPassChange 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Original:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   56
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image ImgPrint 
      Height          =   780
      Left            =   9720
      Picture         =   "Gilroy.frx":B78BE
      ToolTipText     =   "Print Current Volunteer"
      Top             =   0
      Width           =   1050
   End
   Begin VB.Image imgExit 
      Height          =   780
      Left            =   10800
      Picture         =   "Gilroy.frx":BA410
      ToolTipText     =   "Exit"
      Top             =   0
      Width           =   1050
   End
   Begin VB.Image imgNew 
      Height          =   780
      Left            =   7560
      Picture         =   "Gilroy.frx":BCF62
      ToolTipText     =   "Add New Volunteer"
      Top             =   0
      Width           =   1050
   End
   Begin VB.Image ImgDelete 
      Height          =   780
      Left            =   8640
      Picture         =   "Gilroy.frx":BFAB4
      ToolTipText     =   "Delete Current Volunteer"
      Top             =   0
      Width           =   1050
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter A Name To Search For:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Left            =   9480
      TabIndex        =   50
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label lblCanteen 
      BackStyle       =   0  'Transparent
      Caption         =   "Canteen:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   9480
      TabIndex        =   49
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2760
      TabIndex        =   44
      Top             =   6720
      Width           =   3735
   End
   Begin VB.Label lblArtCraft 
      BackStyle       =   0  'Transparent
      Caption         =   "Art && Craft:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6720
      TabIndex        =   43
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Left            =   6720
      TabIndex        =   42
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2760
      TabIndex        =   41
      Top             =   6120
      Width           =   3735
   End
   Begin VB.Label lblPnF 
      BackStyle       =   0  'Transparent
      Caption         =   "P and F:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   40
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Career Field:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   30
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8040
      TabIndex        =   37
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Surname:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   4200
      TabIndex        =   36
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Work:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8160
      TabIndex        =   35
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Home Phone No.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   34
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Students Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3480
      TabIndex        =   33
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   7080
      TabIndex        =   32
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Students Home Room:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   31
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   2760
      TabIndex        =   38
      Top             =   3240
      Width           =   9015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================='
'Year 12 HSC Software Design & Development '
'           Major Assesment                '
'              Ian White                   '
'     Gilroy College Castle Hill           '
'=========================================='
Option Explicit


Dim VolRecCnt As Integer
Public Gilroy As adodb.Connection
Public Volunteers As adodb.Recordset
Public rstTeachers As adodb.Recordset
Public AdvSearchRS As adodb.Recordset
Dim NameQuery As String
Dim Search As String
Public keydrop As Integer
Dim CheckCaption As String
Dim i As Integer
Dim OpIndex As Integer
Dim varReturn As String * 50
Dim varReturn2 As String * 50
Dim varConvert As Double
Dim varConvert2 As Double
Dim CheckValue As Integer
Dim varKey As Integer
Public Sub Form1Display()
    On Error Resume Next
    Dim intValue As Integer
    Dim intValue2 As Integer
    Dim intValue3 As Integer '

    Dim i As Integer
    
    'Below lines put the Database field information
    'into text boxes.
    txtSurName.Text = ""
    txtFirstName.Text = ""
    txtEmail.Text = ""
    txtPhone.Text = "" 'This coes clears the textboxes, so that if you goto a new entry that has a null field, then it will display nothing, not the same field from the previous record.
    txtWorkNo.Text = ""
    txtHomeRoom.Text = ""
    txtCareer.Text = ""
    txtMobile.Text = ""
    txtStudName.Text = ""
        
    On Error Resume Next
    txtSurName.Text = Volunteers.Fields("LastName") & ""
    txtFirstName.Text = Volunteers.Fields("Name") & ""
    txtEmail.Text = Volunteers.Fields("Email") & ""
    txtPhone.Text = Volunteers.Fields("Phone") & ""
    txtWorkNo.Text = Volunteers.Fields("WorkPh") & ""
    txtHomeRoom.Text = Volunteers.Fields("HomeRoom") & "" 'This code displays the correct fields in the corresponding fields.
    txtCareer.Text = Volunteers.Fields("CareerField") & ""
    txtMobile.Text = Volunteers.Fields("Mobile") & ""
    txtStudName.Text = Volunteers.Fields("StudentsName") & ""
    txtYear.Text = Volunteers.Fields("Year") & ""
    
    
    If User = True Then
        User = False    'This code temporarily turns User mode off if it is on so that when it displays the checkboxes you don't get a message box each time saying that you are a user and you can't change them.
        WasUser = True
    Else
        WasUser = False
    End If
    'The below code decodes the binary from the Access DB
    'and displays it in the correct checkboxes.
    intValue = Volunteers.Fields("PandF")
    For i = 0 To 5
        If (intValue And 2 ^ i) Then
            chkOptions(i + 1).Value = vbChecked
        Else
            chkOptions(i + 1).Value = vbUnchecked
        End If
    Next i
    intValue2 = Volunteers.Fields("ArtCraft")
    For i = 0 To 6
        If (intValue2 And 2 ^ i) Then
            chkOptions(i + 11).Value = vbChecked
        Else
            chkOptions(i + 11).Value = vbUnchecked
        End If
    Next i
    intValue3 = Volunteers.Fields("Canteen")
    For i = 0 To 4
        If (intValue3 And 2 ^ i) Then
            chkOptions(i + 21).Value = vbChecked
        Else
            chkOptions(i + 21).Value = vbUnchecked
        End If
    Next i
    If WasUser = True Then
        User = True 'If user mode used to be on, then it turns it back on.
    End If
End Sub
Private Sub ChkOptionAdmin_Click()
    FrmAdminPass.Show
End Sub


Private Sub ChkOptionTeach_Click()
    frmTeacherPass.Show
    AdminHide
End Sub
Private Sub AdminHide()
    'This sub hides the change password fields
    'and locks the textboxes so that user's can not edit them
    Dim i As Integer
             i = 0
             For i = 0 To 7
                 Form1.LabelPassChange(i).Visible = False
             Next i
             For i = 10 To 15
                 Form1.txtPassChange(i).Visible = False
             Next i
             For i = 20 To 21
                 Form1.cmdPassChange(i).Visible = False
             Next i
             
             Form1.txtFirstName.Locked = True
             Form1.txtSurName.Locked = True
             Form1.txtCareer.Locked = True
             Form1.txtPhone.Locked = True
             Form1.txtWorkNo.Locked = True
             Form1.txtMobile.Locked = True
             Form1.txtEmail.Locked = True
             Form1.txtStudName.Locked = True
             Form1.txtHomeRoom.Locked = True
End Sub


Private Sub chkOptions_Click(index As Integer)
    ' the below code works by assigning each set of check boxes to a byte
    'and each checkbox to a bit
    'rather than having one checkbox take up a whole "byte". This is a good idea and works well. For example:

    'If I have 8 checkboxes and the last is turned on, then the byte will be: 00000001
    'ie: the first seven are assigned the value 0, because they are off/unchecked.
    'The computer reads these by getting the overall value, so for this one, the value will be 1, so therefore it knows only the last is turned on.
    
    'For example 00000011 would equal 3 and 00000111 will equal 7. So by looking at the overall byte value, the computer can determine which bits are on and which bits are off.
    
    'I then tell it to display this graphically by assigning the corresponding elements in my chkOptions array a checked or unchecked value. So if the byte is 00000001 then bit 0 is on, therefore chkOptions(0) will be checked, and the rest will be unchecked.

    
    Dim intValue As Integer
    Dim intValue2 As Integer
    Dim intValue3 As Integer
    
    
    If User = True Then
        Form1Display
        MsgBox "Sorry, you do not have adequate access to change these options. Please contact your administrator for any further inquiries.", vbInformation, "Volunteers"
        Exit Sub
    Else
        On Error Resume Next
        Select Case index
            Case 1 To 6 'P and F
                For i = 1 To 6
                        If chkOptions(i).Value = vbChecked Then
                            intValue = intValue Or 2 ^ (i - 1)
                        Else
                            'intValue = intValue Xor 2 ^ (index - 1)
                        End If
                Next i
                Volunteers.Fields("PandF") = intValue
                Volunteers.Update
            Case 11 To 17 'Art and Craft
                For i = 11 To 17
                        If chkOptions(i).Value = vbChecked Then
                            intValue2 = intValue2 Or 2 ^ (i - 11)
                        Else
                            'intValue2 = intValue2 Xor 2 ^ (index - 1)
                        End If
                Next i
                Volunteers.Fields("ArtCraft") = intValue2
                Volunteers.Update
            
            Case 21 To 25 'Canteen - Days of the week
                For i = 21 To 25
                        If chkOptions(i).Value = vbChecked Then
                            intValue3 = intValue3 Or 2 ^ (i - 21)
                        End If
                Next i
                Volunteers.Fields("Canteen") = intValue3
                Volunteers.Update
        End Select
    End If
Exit Sub
End Sub
Private Sub cmdExit_Click()
    End
End Sub

Private Sub ChkOptionUser_Click()
    User = True
    AdminHide
End Sub

Private Sub cmdPassChange_Click(index As Integer)
    'The below code is for administration mode only
    'It is the code that changes the passwords.
    On Error GoTo PassChangeErr
    If index = 20 Then
        If txtPassChange(10).Text = varFinalPass Then 'If the initial password typed = the actual intial password then
            If txtPassChange(11).Text = txtPassChange(12).Text Then 'If the typed and re-typed new password = each other then
                Encrypt (txtPassChange(11).Text) 'Encryot new password
                txtPassChange(10).Text = "" 'Clear text boxes
                txtPassChange(11).Text = ""
                txtPassChange(12).Text = ""
                MsgBox "Password has been succesfully changed.", vbInformation, "Volunteers"
            Else
                MsgBox "New password does not match the re-typed version, please re-type both and try again.", vbInformation, "Volunteers"
            End If
        Else
            MsgBox "Original password appears to be incorrect, please re-type it and try again.", vbInformation, "Volunteers"
        End If
    Call Decrypt 'Calls decrypt so that new password takes effect
    End If
    If index = 21 Then
        If txtPassChange(13).Text = varFinalPass2 Then  'If the initial password typed = the actual intial password then
            If txtPassChange(14).Text = txtPassChange(15).Text Then 'If the typed and re-typed new password = each other then
                Encrypt2 (txtPassChange(14).Text) 'Encryot new password
                txtPassChange(13).Text = "" 'Clear text boxes
                txtPassChange(14).Text = ""
                txtPassChange(15).Text = ""
                MsgBox "Password has been succesfully changed.", vbInformation, "Volunteers"
            Else
                MsgBox "New password does not match the re-typed version, please re-type both and try again.", vbInformation, "Volunteers"
            End If
        Else
            MsgBox "Original password appears to be incorrect, please re-type it and try again.", vbInformation, "Volunteers"
        End If
        Call Decrypt2 'Calls decrypt so that new password takes effect
    End If
    
    Exit Sub
    
PassChangeErr:
    MsgBox "Please enter a new password.", vbInformation, "Volunteers" 'If the new password and re-tpye password fields are empty then it prompts the user.
End Sub
Private Sub cmdPrev_Click()
    If lstNames.ListIndex > 0 Then
        lstNames.ListIndex = lstNames.ListIndex - 1
    End If
End Sub
Private Sub cmdNext_Click()
    If lstNames.ListIndex < lstNames.ListCount - 1 Then 'Because listindex is 0 based, if there is 165 entries, then the last will be 164.
        lstNames.ListIndex = lstNames.ListIndex + 1
        Call lstNames_Click
    End If
End Sub
Private Sub cmdSearch_Click()
     On Error GoTo cmdSearch_Click_Err
    If Text2(1).Text = "" Then
        MsgBox "Please enter a name to search for."
    Else
        'Set Volunteers = Gilroy.OpenRecordset("SELECT * FROM Volunteers")
        NameQuery = Text2(1).Text
        NameQuery = StrConv(NameQuery, vbProperCase)
    
        frmResults.Show 'Brings up results page
        frmResults.lstResultID.Clear 'clears list box ready for new data
        frmResults.lstResults.Clear 'clears lsit box ready for new data
        Volunteers.MoveFirst 'Moves to first record
        Do Until Volunteers.EOF
            Volunteers.Find "lastname like '" & NameQuery & "*'" 'Searches for record via LASTNAME
            If Not Volunteers.EOF Then
                frmResults.lstResultID.AddItem Volunteers.Fields("ID") 'adds the fullname from the record into the list box
                frmResults.lstResults.AddItem Volunteers.Fields("Fullname") 'Adds the ID number into the ID list box
                Volunteers.MoveNext 'Moves to next record ready to search again
            End If
        Loop
        Volunteers.MoveFirst
    
        Do Until Volunteers.EOF
            Volunteers.Find "Name like '" & NameQuery & "*'" 'same as above code, but checks through FIRSTNAMES
            If Not Volunteers.EOF Then
                frmResults.lstResults.AddItem Volunteers.Fields("Fullname")
                Volunteers.MoveNext
            End If
        Loop
    End If
Exit Sub

cmdSearch_Click_Err:
End Sub
Private Sub Command1_Click()
    frmAdSearch.Show
End Sub
Private Sub cmdNoOfVolunteers_Click()
    MsgBox "There are " & lstNames.ListCount & " volunteers.", vbInformation, "Volunteers"
End Sub
Private Sub Form_Load()
  
    Dim ListNamesID As Integer
    ListNamesID = 0
    On Error GoTo Form_Load_Err
          
    i = 1
    'Below code connects prgram to database
    Set Gilroy = New adodb.Connection
    Gilroy.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=MS Access Database;Initial Catalog=gilroy4.mdb"
    
    'below code sets recordet (where to get info from)
    Set Volunteers = New adodb.Recordset
    Volunteers.Open "SELECT * FROM Volunteers", Gilroy, adOpenDynamic, adLockOptimistic
    Set AdvSearchRS = New adodb.Recordset
    AdvSearchRS.Open "SELECT * FROM AdvSearch", Form1.Gilroy, adOpenDynamic, adLockOptimistic
    
    If Volunteers.EOF Then
        MsgBox "Please Add Someone To The Database"
    Else
        Volunteers.MoveFirst
    End If
    
    Form1Display 'displays informtaion of first record
    lstNames.Clear
    
    'below code adds all volunteers to list box on main form
    If Not Volunteers.BOF Then
        Volunteers.MoveFirst
        Do Until Volunteers.EOF
            lstNames.AddItem Volunteers("FullName")
            lstID.AddItem Volunteers("ID")
            Volunteers.MoveNext
        Loop
        Else
            MsgBox "Please add someone to the database."
    End If
    lstNames.ListIndex = 0
    Volunteers.MoveFirst
    User = True
    'the below code sets another recordset that is used to store the encrypted passwords
    Set rstTeachers = New adodb.Recordset
    rstTeachers.Open "SELECT * FROM Teachers", Gilroy, adOpenDynamic, adLockOptimistic
    rstTeachers.MoveFirst
    
    Call Decrypt
    Call Decrypt2
    
           'The following section of code adds the names to the list box
    
    
 Call Sort_it 'sorts list


Form_Load_Err:
    Select Case Err.Number
    Case 94
        Exit Sub
    End Select
End Sub

Private Sub ImgDelete_Click()
    'Below code delets current record if the user is an administrator
    If Admin = False Or User = True Then
        MsgBox "Sorry, unfortunately only someone with Administrator Access can perform that operation. Please contact your administrator for further queries."
        
    Else
        Text1.Visible = True
        If MsgBox("Are you sure you want to delete this volunteer's record?", vbYesNo, "Volunteers") = vbYes Then
            
            Volunteers.Delete
            If lstNames.ListIndex = 0 Then
                lstNames.ListIndex = 1
            End If
            ListNamesUpdate
            
        Else
            Exit Sub
        End If
    End If
   
End Sub

Private Sub ImgPrint_Click()
    'The below code simply asks if the user is ready to print
    If Form1.lstNames.ListIndex = -1 Then
        If MsgBox("Warning: You Do Not Have An Entry Selected. If You Continue, Blank Fields Will Be Printed.", vbCritical) = vbOK Then
        End If
    End If
        If MsgBox("Are You Ready To Print?", vbYesNo, "Volunteers") = vbYes Then Call Print_it
End Sub

Private Sub imgExit_Click()
    End
End Sub

Private Sub AdminShow()
    'Shows relevant administration controls
    'eg password controls
    Dim i As Integer
    i = 0
    For i = 0 To 7
        LabelPassChange(i).Visible = True
    Next i
    For i = 10 To 15
        txtPassChange(i).Visible = True
    Next i
    For i = 20 To 21
        cmdPassChange(i).Visible = True
    Next i
    
End Sub
Private Sub imgNew_Click()
    frmNew.Show
End Sub


Private Sub lstNames_Click()
    'Displays the name when you click on it.
    lstID.ListIndex = lstNames.ListIndex
    Volunteers.MoveFirst
    Volunteers.Find "ID like '" & lstID.Text & "'"
    Form1Display
    
End Sub

Private Sub Opt2000Vol_Click()
    'the below code shows only volunteers involved with the school
    'in the year 2001
    lstNames.Clear
    lstID.Clear
    Volunteers.MoveFirst
        Do Until Volunteers.EOF
            Volunteers.Find "Year like '" & 2000 & "'"
            If Not Volunteers.EOF Then
                lstNames.AddItem Volunteers.Fields("Fullname")
                lstID.AddItem Volunteers.Fields("ID")
                Volunteers.MoveNext
            End If
        Loop
    
    If lstNames.ListCount > 0 Then
        lstNames.ListIndex = 0
    End If
    lstID.ListIndex = lstNames.ListIndex
    
    Call Sort_it
    
End Sub

Private Sub Opt2001Vol_Click()
    'the below code shows only volunteers involved with the school
    'in the year 2001
    lstNames.Clear
    lstID.Clear
    Volunteers.MoveFirst
        Do Until Volunteers.EOF
            Volunteers.Find "Year like '" & 2001 & "'"
            If Not Volunteers.EOF Then
                lstNames.AddItem Volunteers.Fields("Fullname")
                lstID.AddItem Volunteers.Fields("ID")
                Volunteers.MoveNext
            End If
        Loop
    
    lstNames.ListIndex = 0
    lstID.ListIndex = lstNames.ListIndex
    Call Sort_it
End Sub

Private Sub OptAllVol_Click()
    'The below code shows all the volunteers
    lstNames.Clear
    lstID.Clear
    Volunteers.MoveFirst
    Do Until Volunteers.EOF
        lstNames.AddItem Volunteers.Fields("Fullname")
        lstID.AddItem Volunteers.Fields("ID")
        Volunteers.MoveNext
    Loop
    
    lstNames.ListIndex = 0
    lstID.ListIndex = lstNames.ListIndex
    Call Sort_it
    Call lstNames_Click
End Sub

Private Sub Text2_KeyPress(index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case 13
        Text2(1) = StrConv(Text2(1), vbProperCase)
        Call cmdSearch_Click
    End Select
End Sub

Private Sub Text2_LostFocus(index As Integer)
    Text2(1) = StrConv(Text2(1), vbProperCase) 'makes the first letter a capital.
End Sub

'simple just pass the password to it like this
'Encrypt("password")


Private Function Encrypt(varPass As String)
    Dim EncryptedPass As String
    
    For i = 1 To Len(varPass)  'from 1 to the length of the password
        Dim encryptedChar As String, tmp As Integer
        
        'the below code encrypts the password by changing the characters
        tmp = Asc(Mid$(varPass, i, 1))
        encryptedChar = Chr(tmp - 2)
        EncryptedPass = EncryptedPass + encryptedChar
    Next i
    
    varFinalPass = varPass
    
    'Saves encrypted password into database.
    rstTeachers.Fields("Gilroy") = EncryptedPass
    rstTeachers.Update
End Function
'returns the decrypted pass
'like if decrypt() = "password" then
Private Function Encrypt2(varPass2 As String)
    Dim EncryptedPass2 As String
    
    For i = 1 To Len(varPass2) 'from 1 to the length of the password
        Dim encryptedChar2 As String, tmp2 As Integer
        
        'the below code encrypts the password by changing the characters
        tmp2 = Asc(Mid$(varPass2, i, 1))
        encryptedChar2 = Chr(tmp2 - 2)
        EncryptedPass2 = EncryptedPass2 + encryptedChar2
    Next i
    
    varFinalPass2 = varPass2
    
    'Saves encrypted password into database.
    rstTeachers.Fields("Teachers") = EncryptedPass2
    rstTeachers.Update
End Function

'returns the decrypted pass
'like if decrypt() = "password" then

Private Function Decrypt()

    Dim EncryptedPass As String, DecryptedPass As String
    
    EncryptedPass = rstTeachers.Fields("Gilroy")
    
    For i = 1 To Len(EncryptedPass)  'from 1 to the length of the encrypted password
        Dim DecryptedChar As String, tmp As Integer
        
        'the below code decrypts the password by changing the characters back to how they were
        tmp = Asc(Mid$(EncryptedPass, i, 1))
        DecryptedChar = Chr(tmp + 2)
        
        DecryptedPass = DecryptedPass + DecryptedChar
    Next i
    
    'assigns the decypted password to a string for refernce by program.
    varFinalPass = DecryptedPass
End Function

Private Function Decrypt2()

    Dim EncryptedPass2 As String, DecryptedPass2 As String
    
    EncryptedPass2 = rstTeachers.Fields("Teachers")
    
    For i = 1 To Len(EncryptedPass2) 'from 1 to the length of the encrypted password
        Dim DecryptedChar2 As String, tmp2 As Integer
        
        'the below code decrypts the password by changing the characters back to how they were
        tmp2 = Asc(Mid$(EncryptedPass2, i, 1))
        DecryptedChar2 = Chr(tmp2 + 2)
        
        DecryptedPass2 = DecryptedPass2 + DecryptedChar2
    Next i
    
    'assigns the decypted password to a string for refernce by program.
    varFinalPass2 = DecryptedPass2
End Function

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If User = True Then
        MsgBox "Sorry, you do not have adequate access to change these options. Please contact your administrator for any further inquiries.", vbInformation, "Volunteers"
    End If
End Sub

Private Sub Sort_it()
    ' sort them alphabetically
    ' uses a bubble sort
    Dim a As Integer, b As Integer
    For a = 0 To lstNames.ListCount - 2
        For b = a + 1 To lstNames.ListCount - 1
            ' compare and swap if necessary
            If lstNames.List(b) < lstNames.List(a) Then
                Call SwapPeople(a, b)
            End If
        Next b
    Next a
    Call lstNames_Click ' show it all up now again
    Text1.Visible = False
End Sub
Private Sub SwapList(lst As ListBox, a As Integer, b As Integer)
    'This code swaps the listbox information if needed.
    Dim temp As String
    temp = lst.List(a)
    lst.List(a) = lst.List(b)
    lst.List(b) = temp
End Sub
Private Sub SwapPeople(a As Integer, b As Integer)
' used by the sort to swap two values
    Call SwapList(lstNames, a, b)
    Call SwapList(lstID, a, b)
End Sub
Private Sub Print_it()
    Dim D As Integer
    D = 0
    'The below code sets all the printer properties and prints
    'The current record that the user is viewing.
    Printer.Font = "arial"
    Printer.FontSize = 18
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.Print "Gilroy College - Volunteers Database"
    Printer.Print ""
    Printer.Font = "Arial"
    Printer.FontUnderline = False
    Printer.FontSize = 12
    
    Printer.Print "Surname: ", ,
    Printer.FontBold = False
    Printer.Print Form1.txtSurName.Text
    
    Printer.FontBold = True
    Printer.Print "First Name: ", ,
    Printer.FontBold = False
    Printer.Print Form1.txtFirstName.Text
    
    Printer.FontBold = True
    Printer.Print "Home Phone Number: ",
    Printer.FontBold = False
    Printer.Print Form1.txtPhone.Text
    
    Printer.FontBold = True
    Printer.Print "Work Phone Number: ",
    Printer.FontBold = False
    Printer.Print Form1.txtWorkNo.Text
    
    Printer.FontBold = True
    Printer.Print "Mobile Phone Number: ",
    Printer.FontBold = False
    Printer.Print Form1.txtMobile.Text
    
    Printer.FontBold = True
    Printer.Print "Career Field: ",
    Printer.FontBold = False
    Printer.Print Form1.txtCareer.Text
    
    Printer.FontBold = True
    Printer.Print "Students Name: ",
    Printer.FontBold = False
    Printer.Print Form1.txtStudName.Text
    
    Printer.FontBold = True
    Printer.Print "Student's Home Room: ",
    Printer.FontBold = False
    Printer.Print Form1.txtHomeRoom.Text
    
    Printer.FontBold = True
    Printer.Print "Email: ", ,
    Printer.FontBold = False
    Printer.Print Form1.txtEmail.Text
    
    Printer.FontBold = True
    Printer.Print "Year: ", ,
    Printer.FontBold = False
    Printer.Print Form1.txtYear.Text
    
    'The below code sets the printer properties for
    'printing checkboxes.
    'The if statements are to make sure that the information
    'lines up (the Yes' and No's)
    
    For D = 1 To 6
        If D = 1 Or D = 3 Or D = 5 Or D = 6 Then
                Printer.FontBold = True
                Printer.Print chkOptions(D).Caption,
                Printer.FontBold = False
            If chkOptions(D).Value = 1 Then
                 Printer.Print "Yes"
            Else
                Printer.Print "No"
            End If
        Else
            Printer.FontBold = True
            Printer.Print chkOptions(D).Caption, ,
            Printer.FontBold = False
            If chkOptions(D).Value = 1 Then
                Printer.Print "Yes"
            Else
                Printer.Print "No"
            End If
        End If
    Next D
    D = 11
    For D = 11 To 17
        If D = 17 Then
            Printer.FontBold = True
            Printer.Print chkOptions(D).Caption,
            Printer.FontBold = False
            If chkOptions(D).Value = 1 Then
                Printer.Print "Yes"
            Else
                Printer.Print "No"
            End If
        Else
            Printer.FontBold = True
            Printer.Print chkOptions(D).Caption, ,
            Printer.FontBold = False
            If chkOptions(D).Value = 1 Then
                Printer.Print "Yes"
            Else
                Printer.Print "No"
            End If
        End If
    Next D
    D = 21
    For D = 21 To 25
        Printer.FontBold = True
        Printer.Print chkOptions(D).Caption, ,
        Printer.FontBold = False
        If chkOptions(D).Value = 1 Then
            Printer.Print "Yes"
        Else
            Printer.Print "No"
        End If
    Next D
    
    Printer.EndDoc
    End Sub
Public Sub ListNamesUpdate()
     lstNames.Clear
     Text1.Visible = True
     If Not Volunteers.BOF Then
        Volunteers.MoveFirst 'Moves to first record
        Do Until Volunteers.EOF
            lstNames.AddItem Volunteers("FullName") 'Adds records to listbox
            lstID.AddItem Volunteers("ID") 'Adds record ID's to the ID list box
            Volunteers.MoveNext 'Moves to the next record
        Loop
        Else
            MsgBox "Please add someone to the database." 'If there is no one in the database.
    End If
    lstNames.ListIndex = 0 'goes to first name in list box
    Call Sort_it 'Sorts in alphabetical order
End Sub



Private Sub txtCareer_LostFocus()
    Call txtSurName_lostfocus
End Sub
Private Sub txtemail_lostfocus()
    Call txtSurName_lostfocus
End Sub

Private Sub txtHomeRoom_lostfocus()
    Call txtSurName_lostfocus
End Sub
Private Sub txtMobile_lostfocus()
    Call txtSurName_lostfocus
End Sub
Private Sub txtPhone_lostfocus()
    Call txtSurName_lostfocus
End Sub
Private Sub txtStudname_lostfocus()
    Call txtSurName_lostfocus
End Sub
Private Sub txtYear_lostfocus()
    Call txtSurName_lostfocus
End Sub
Private Sub txtWorkNo_lostfocus()
    Call txtSurName_lostfocus
End Sub
Private Sub txtFirstName_lostfocus()
    Call txtSurName_lostfocus
End Sub

Private Sub txtSurName_lostfocus()
    'The code in this sub saves information
    'if it is edited in the textboxes
    'by an admin.
    
    Dim ListIndexTemp As Integer
    If Admin = True Then
        ListIndexTemp = lstNames.ListIndex
        'The below code saves the changed information in to the database.
        Volunteers.Fields("Lastname") = txtSurName.Text & ""
        Volunteers.Fields("Name") = txtFirstName.Text & ""
        Volunteers.Fields("Email") = txtEmail.Text & ""
        Volunteers.Fields("Phone") = txtPhone.Text & ""
        Volunteers.Fields("WorkPh") = txtWorkNo.Text & ""
        Volunteers.Fields("HomeRoom") = txtHomeRoom.Text & ""
        Volunteers.Fields("CareerField") = txtCareer.Text & ""
        Volunteers.Fields("Mobile") = txtMobile.Text & ""
        Volunteers.Fields("StudentsName") = txtStudName.Text & ""
        Volunteers.Fields("Year") = txtYear.Text & ""
        Volunteers.Update
        Volunteers.Fields("FullName") = Volunteers.Fields("Lastname") & ", " & Volunteers.Fields("Name") 'Sets the FullName field in the database
        Volunteers.Update
        Call ListNamesUpdate 'Updates the listbox on the main form
        
        lstNames.ListIndex = ListIndexTemp
        lstNames.Text = Volunteers.Fields("Lastname") & ", " & Volunteers.Fields("Name") 'Changes the text that appears on the main listbox, this is so that if the admin changes the first or last name, this change will appear in the main listbox.
        lstNames.ListIndex = ListIndexTemp
    Else
        Exit Sub
    End If
    
End Sub
