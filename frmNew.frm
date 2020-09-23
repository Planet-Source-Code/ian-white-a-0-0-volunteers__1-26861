VERSION 5.00
Begin VB.Form frmNew 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Volunteer Record"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAddFields 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3840
      MousePointer    =   3  'I-Beam
      TabIndex        =   6
      Text            =   "2001"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add"
      Height          =   855
      Left            =   2640
      TabIndex        =   28
      Top             =   5640
      Width           =   2415
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   855
      Left            =   2640
      TabIndex        =   29
      Top             =   6600
      Width           =   2415
   End
   Begin VB.TextBox txtAddFields 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   2040
      MousePointer    =   3  'I-Beam
      TabIndex        =   8
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox txtAddFields 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3840
      MousePointer    =   3  'I-Beam
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtAddFields 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2040
      MousePointer    =   3  'I-Beam
      TabIndex        =   7
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox txtAddFields 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2040
      MousePointer    =   3  'I-Beam
      TabIndex        =   9
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox txtAddFields 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2040
      MousePointer    =   3  'I-Beam
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtAddFields 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2040
      MousePointer    =   3  'I-Beam
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtAddFields 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2040
      MousePointer    =   3  'I-Beam
      TabIndex        =   2
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txtAddFields 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2040
      MousePointer    =   3  'I-Beam
      TabIndex        =   1
      Top             =   960
      Width           =   3015
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Tuesday:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   22
      Left            =   1320
      TabIndex        =   24
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Monday:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   21
      Left            =   1440
      TabIndex        =   23
      Top             =   5760
      Width           =   975
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Wednesday:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   23
      Left            =   1200
      TabIndex        =   25
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Thursday:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   24
      Left            =   1320
      TabIndex        =   26
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Friday:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   25
      Left            =   1560
      TabIndex        =   27
      Top             =   7200
      Width           =   855
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Uniforms Shop:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   3
      Left            =   960
      TabIndex        =   12
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Major Raffle:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   4
      Left            =   1200
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Library Work:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Index           =   5
      Left            =   960
      TabIndex        =   14
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Careers Night:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   6
      Left            =   960
      TabIndex        =   15
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "General Help:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   17
      Left            =   3600
      TabIndex        =   22
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Promotion:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   16
      Left            =   3840
      TabIndex        =   21
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Sales:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   15
      Left            =   4080
      TabIndex        =   20
      Top             =   4680
      Width           =   855
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Setting Up:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   14
      Left            =   3720
      TabIndex        =   19
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Art:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   13
      Left            =   4320
      TabIndex        =   18
      Top             =   4200
      Width           =   615
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Catering:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   11
      Left            =   3960
      TabIndex        =   16
      Top             =   3705
      Width           =   975
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Craft:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   12
      Left            =   4200
      TabIndex        =   17
      Top             =   3960
      Width           =   735
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Ball:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Index           =   2
      Left            =   1800
      TabIndex        =   11
      Top             =   3960
      Width           =   615
   End
   Begin VB.CheckBox chkOptions 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      Caption         =   "Social Commitee:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   1
      Left            =   840
      TabIndex        =   10
      Top             =   3705
      Width           =   1575
   End
   Begin VB.TextBox txtAddFields 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   3360
      TabIndex        =   46
      Top             =   2040
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Career Field:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   1020
      TabIndex        =   45
      Top             =   1320
      Width           =   915
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   1425
      TabIndex        =   44
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   1530
      TabIndex        =   43
      Top             =   960
      Width           =   450
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Work:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   3390
      TabIndex        =   42
      Top             =   1680
      Width           =   420
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Home Phone No.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   41
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Students Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   795
      TabIndex        =   40
      Top             =   2400
      Width           =   1140
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   1650
      TabIndex        =   39
      Top             =   3120
      Width           =   405
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Students Home Room:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   345
      TabIndex        =   38
      Top             =   2760
      Width           =   1590
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P and F:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00289CFE&
      Height          =   240
      Left            =   120
      TabIndex        =   37
      Top             =   3360
      Width           =   795
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Art && Craft:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00289CFE&
      Height          =   240
      Left            =   2640
      TabIndex        =   36
      Top             =   3360
      Width           =   1035
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Canteen:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00289CFE&
      Height          =   240
      Left            =   120
      TabIndex        =   35
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   120
      TabIndex        =   34
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   2640
      TabIndex        =   33
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   120
      TabIndex        =   32
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Surname:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   1320
      TabIndex        =   31
      Top             =   600
      Width           =   690
   End
   Begin VB.Label lblAddEntry 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add New Gilroy College Volunteer Record."
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
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim AddintValue As Integer
    Dim AddintValue2 As Integer
    Dim AddintValue3 As Integer
    Dim ChkValue1 As Integer
    Dim ChkValue2 As Integer
    Dim ChkValue3 As Integer
    Dim i As Integer
    Dim n As Integer
Private Sub chkOptions_Click(index As Integer)
      i = 1
    '    On Error Resume Next
        Select Case index
            Case 1 To 6 'P and F
                For i = 1 To 6
                        If chkOptions(i).Value = vbChecked Then
                            AddintValue = AddintValue Or 2 ^ (i - 1)
                        Else
                            'AddIntValue = AddIntValue Xor 2 ^ (index - 1)
                        End If
                        ChkValue1 = AddintValue
                Next i
             
            Case 11 To 17 'Art and Craft
                For i = 11 To 17
                        If chkOptions(i).Value = vbChecked Then
                            AddintValue2 = AddintValue2 Or 2 ^ (i - 11)
                        Else
                            'AddIntValue2 = AddIntValue2 Xor 2 ^ (index - 1)
                        End If
                        ChkValue2 = AddintValue2
                Next i
                
            
            Case 21 To 25 'Canteen - Days of the week
                For i = 21 To 25
                        If chkOptions(i).Value = vbChecked Then
                            AddintValue3 = AddintValue3 Or 2 ^ (i - 21)
                        End If
                        ChkValue3 = AddintValue3
                Next i
                
        End Select
Exit Sub
End Sub

Private Sub Command1_Click()
MsgBox ChkValue1
MsgBox ChkValue2
MsgBox ChkValue3
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub txtAddfields_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then 'If they hit enter, it will add the entry to the database
        Call cmdAddNew_Click 'rather than them having to click the Add button
    End If
End Sub
Private Sub cmdExit_Click()
    Unload Me 'Unloads the Add new vounteer form. Ie closes it and deletes information that may be left there in textboxes.
End Sub

Private Sub cmdAddNew_Click()
    If txtAddFields(1).Text = "" Then 'If there is nothing in the surname box
        MsgBox "Please enter a Surname", vbInformation, "Volunteers" 'tell them to pu tone in
    Else
        Form1.Volunteers.AddNew 'Add new record into DB
        Form1.Volunteers!Lastname = txtAddFields(1).Text    'V
        Form1.Volunteers!Name = txtAddFields(0).Text        'The following code
        Form1.Volunteers!Phone = txtAddFields(2).Text       'Adds the corrsesponfing
        Form1.Volunteers!WorkPh = txtAddFields(3).Text      'information into
        Form1.Volunteers!Mobile = txtAddFields(4).Text      'the correct DB fields
        Form1.Volunteers!Email = txtAddFields(5).Text       '^
        Form1.Volunteers!CareerField = txtAddFields(6).Text '^
        Form1.Volunteers!StudentsName = txtAddFields(7).Text '^
        Form1.Volunteers!Homeroom = txtAddFields(8).Text    '^
        If Len(txtAddFields(9).Text) Then
            Form1.Volunteers!Year = txtAddFields(9).Text
        Else
            MsgBox "Please enter a year.", vbInformation, "Volunteers"
        End If
        Form1.Volunteers!PandF = ChkValue1
        Form1.Volunteers!ArtCraft = ChkValue2
        Form1.Volunteers!Canteen = ChkValue3
        
        Form1.Volunteers.Fields("Fullname") = Form1.Volunteers.Fields("Lastname") & ", " & Form1.Volunteers.Fields("Name") 'This fills in the Fullname field by adding the first and last names together, the fullname field is used for displaying purposes.
        Form1.Volunteers.Update 'Updates the database to incorporate new changes.
    'End If
    MsgBox "Your entry has been added."
    Form1.ListNamesUpdate 'This calls a sub that re-fills-in the listbox on the main form so that the new entry is there
    End If
    For n = 0 To 9
        txtAddFields(n).Text = "" 'clears textboxes
    Next n
    For n = 1 To 6
        chkOptions(n).Value = 0
    Next n
    For n = 11 To 17
        chkOptions(n).Value = 0 'These clear the checkboxes
    Next n
    For n = 21 To 25
        chkOptions(n).Value = 0
    Next n

End Sub
