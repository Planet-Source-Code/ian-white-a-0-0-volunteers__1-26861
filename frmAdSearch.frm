VERSION 5.00
Begin VB.Form frmAdSearch 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced Search"
   ClientHeight    =   6060
   ClientLeft      =   1470
   ClientTop       =   1665
   ClientWidth     =   8925
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00289CFE&
   Icon            =   "frmAdSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAdSearch.frx":1002
   ScaleHeight     =   6060
   ScaleWidth      =   8925
   Visible         =   0   'False
   Begin VB.TextBox txtAdFields 
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
      Left            =   4680
      MousePointer    =   3  'I-Beam
      TabIndex        =   45
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtAdFields 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Top             =   1800
      Width           =   3015
   End
   Begin VB.CommandButton cmdAdvSearch 
      Caption         =   "Search"
      Height          =   615
      Left            =   840
      TabIndex        =   27
      Top             =   4200
      Width           =   1335
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
      Left            =   5880
      TabIndex        =   9
      Top             =   1545
      Width           =   1575
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
      Left            =   6840
      TabIndex        =   10
      Top             =   1800
      Width           =   615
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
      Left            =   6720
      TabIndex        =   16
      Top             =   3720
      Width           =   735
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
      Left            =   6480
      TabIndex        =   15
      Top             =   3465
      Width           =   975
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
      Left            =   6840
      TabIndex        =   17
      Top             =   3960
      Width           =   615
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
      Left            =   6240
      TabIndex        =   18
      Top             =   4200
      Width           =   1215
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
      Left            =   6600
      TabIndex        =   19
      Top             =   4440
      Width           =   855
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
      Left            =   6360
      TabIndex        =   20
      Top             =   4680
      Width           =   1095
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
      Left            =   6120
      TabIndex        =   21
      Top             =   4920
      Width           =   1335
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
      Left            =   6000
      TabIndex        =   14
      Top             =   2760
      Width           =   1455
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
      Left            =   6000
      TabIndex        =   13
      Top             =   2520
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
      Left            =   6240
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
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
      Left            =   6000
      TabIndex        =   11
      Top             =   2040
      Width           =   1455
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
      Left            =   4440
      TabIndex        =   26
      Top             =   4920
      Width           =   855
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
      Left            =   4200
      TabIndex        =   25
      Top             =   4560
      Width           =   1095
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
      Left            =   2880
      TabIndex        =   24
      Top             =   4560
      Width           =   1215
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
      Left            =   3120
      TabIndex        =   22
      Top             =   4200
      Width           =   975
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
      Left            =   4200
      TabIndex        =   23
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txtAdFields 
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
      Height          =   240
      Index           =   0
      Left            =   2400
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtAdFields 
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
      Left            =   2400
      MousePointer    =   3  'I-Beam
      TabIndex        =   2
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox txtAdFields 
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
      Left            =   2400
      MousePointer    =   3  'I-Beam
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtAdFields 
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
      Left            =   2400
      MousePointer    =   3  'I-Beam
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtAdFields 
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
      Left            =   4200
      MousePointer    =   3  'I-Beam
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtAdFields 
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
      Left            =   2400
      MousePointer    =   3  'I-Beam
      TabIndex        =   7
      Top             =   3240
      Width           =   3015
   End
   Begin VB.TextBox txtAdFields 
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
      Left            =   4200
      MousePointer    =   3  'I-Beam
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtAdFields 
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
      Left            =   2400
      MousePointer    =   3  'I-Beam
      TabIndex        =   8
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label18 
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
      Left            =   4215
      TabIndex        =   46
      Top             =   1440
      Width           =   405
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
      Left            =   1680
      TabIndex        =   43
      Top             =   1800
      Width           =   690
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   2760
      TabIndex        =   42
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   5640
      TabIndex        =   41
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   5640
      TabIndex        =   40
      Top             =   1440
      Width           =   1935
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
      Left            =   2760
      TabIndex        =   44
      Top             =   3840
      Width           =   855
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
      Left            =   5640
      TabIndex        =   39
      Top             =   3120
      Width           =   1035
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
      Left            =   5640
      TabIndex        =   38
      Top             =   1200
      Width           =   795
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
      Left            =   705
      TabIndex        =   37
      Top             =   3600
      Width           =   1590
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
      Left            =   3810
      TabIndex        =   36
      Top             =   2880
      Width           =   405
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
      Left            =   1155
      TabIndex        =   35
      Top             =   3240
      Width           =   1140
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
      Left            =   1080
      TabIndex        =   34
      Top             =   2520
      Width           =   1275
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
      Left            =   3750
      TabIndex        =   33
      Top             =   2520
      Width           =   420
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
      Left            =   1920
      TabIndex        =   32
      Top             =   1440
      Width           =   450
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
      Left            =   1785
      TabIndex        =   31
      Top             =   2880
      Width           =   495
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
      Left            =   1380
      TabIndex        =   30
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Advanced Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   29
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Gilroy College: Volunteers Database"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00289CFE&
      Height          =   375
      Left            =   2280
      TabIndex        =   28
      Top             =   720
      Width           =   5775
   End
End
Attribute VB_Name = "frmAdSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdField As String
Dim VolName As String
Dim intAdvValue As Integer
Dim intAdvValue2 As Integer
Dim intAdvValue3 As Integer

Private Sub chkOptions_Click(index As Integer)
    On Error Resume Next
    intAdvValue = 0
    intAdvValue2 = 0
    intAdvValue3 = 0
    Select Case index
        Case 1 To 6 'P and F
            For i = 1 To 6
                    If chkOptions(i).Value = vbChecked Then
                        intAdvValue = intAdvValue Or 2 ^ (i - 1)
                    End If
            Next i
        Case 11 To 17 'Art and Craft
            For i = 11 To 17
                    If chkOptions(i).Value = vbChecked Then
                        intAdvValue2 = intAdvValue2 Or 2 ^ (i - 11)
                    Else
                        'intAdvValue2 = intAdvValue2 Xor 2 ^ (index - 1)
                    End If
            Next i
        Case 21 To 25 'Canteen - Days of the week
            For i = 21 To 25
                    If chkOptions(i).Value = vbChecked Then
                        intAdvValue3 = intAdvValue3 Or 2 ^ (i - 21)
                    End If
            Next i
    End Select
    
    
    
    
    
Exit Sub
End Sub

Private Sub cmdAdvSearch_click()
Dim NameSearch As String
Dim LastNameSearch As String
Dim CareerFieldSearch As String
Dim HomePhSearch As String
Dim WorkPhSearch As String
Dim MobileSearch As String
Dim EmailSearch As String
Dim StudentsNameSearch As String
Dim StudentsHomeRoomSearch As String
Dim PandFSearch As String
Dim ArtCraftSearch As String
Dim CanteenSearch As String, strFilter As String
Dim YearSearch As String

    'Sample SQL:
    'SELECT Volunteers.Name, Volunteers.Lastname FROM Volunteers WHERE (((Volunteers.Name) Like "mar*"));
    'Form1.Volunteers.Find "Name like '" & txtAdFields(0).Text & "*'" And "Lastname like '" & txtAdFields(1).Text & "*'"
    
    
    'The below code checks through each search choice (all the textboxes etc) and if the box contains text
    'then it assigns searchcode to the corresponding string. Eg, if the Name field has somthing in it, then
    'it assigns "Name Like '*" & txtAdFields(0).Text & "*'" code to a string, NameSearch
    If Not Form1.Volunteers.BOF Then
        Form1.Volunteers.MoveFirst
    End If
    If txtAdFields(0).Text = "" Then
        NameSearch = ""
    Else
        NameSearch = "Name Like '*" & txtAdFields(0).Text & "*'"
    End If
    If txtAdFields(1).Text = "" Then
        LastNameSearch = ""
    Else
        LastNameSearch = "Lastname Like '*" & txtAdFields(1).Text & "*'"
    End If
    If txtAdFields(2).Text = "" Then
        CareerFieldSearch = ""
    Else
        CareerFieldSearch = "CareerField Like '*" & txtAdFields(2).Text & "*'"
    End If
    If txtAdFields(3).Text = "" Then
        HomePhSearch = ""
    Else
        HomePhSearch = "Phone Like '*" & txtAdFields(3).Text & "*'"
    End If
    If txtAdFields(4).Text = "" Then
        WorkPhSearch = ""
    Else
        WorkPhSearch = "WorkPh Like '*" & txtAdFields(4).Text & "*'"
    End If
    If txtAdFields(5).Text = "" Then
        MobileSearch = ""
    Else
        MobileSearch = "Mobile Like '*" & txtAdFields(5).Text & "*'"
    End If
    If txtAdFields(6).Text = "" Then
        EmailSearch = ""
    Else
        EmailSearch = "Email Like '*" & txtAdFields(6).Text & "*'"
    End If
    If txtAdFields(7).Text = "" Then
        StudentsNameSearch = ""
    Else
        StudentsNameSearch = "StudentsName Like '*" & txtAdFields(7).Text & "*'"
    End If
    If txtAdFields(8).Text = "" Then
        StudentsHomeRoomSearch = ""
    Else
        StudentsHomeRoomSearch = "HomeRoom Like '*" & txtAdFields(8).Text & "*'"
    End If
    
    If txtAdFields(9).Text = "" Then
        YearSearch = ""
    Else
        YearSearch = "Year Like '" & txtAdFields(9).Text & "'"
    End If
    
    'This does the same for checkboxes as the above code does for textboxes.
    'It works on binary numbers
    If intAdvValue = 0 Then
        PandFSearch = ""
    Else
        PandFSearch = "PandF Like '" & intAdvValue & "'"
    End If
    If intAdvValue2 = 0 Then
        ArtCraftSearch = ""
    Else
        ArtCraftSearch = "ArtCraft Like '" & intAdvValue2 & "'"
    End If
    If intAdvValue3 = 0 Then
        CanteenSearch = ""
    Else
        CanteenSearch = "Canteen Like '" & intAdvValue3 & "'"
    End If

'=======================================================================
'The below code builds up a search string.
'It does this by going through all the previous strings,
'eg NameSearch, PhoneSearch etc and if they contain a value
'then the program adds them to a search string and gradually builds up
'a final search string that will search for everything at once, using
'an AND search method.
'This final search string is all assigned to one string, strFilter ready
'to be used further down.
'========================================================================

    If Len(NameSearch) Then 'Name
        strFilter = NameSearch
    End If
    If Len(LastNameSearch) Then 'Surname
        If Len(strFilter) Then
            strFilter = strFilter & " AND "
        End If
        strFilter = strFilter & LastNameSearch 'adds to string
    End If
    If Len(CareerFieldSearch) Then
        If Len(strFilter) Then
            strFilter = strFilter & " AND "
        End If
        strFilter = strFilter & CareerFieldSearch
    End If
    If Len(HomePhSearch) Then
        If Len(strFilter) Then
            strFilter = strFilter & " AND "
        End If
        strFilter = strFilter & HomePhSearch
    End If
    If Len(MobileSearch) Then
        If Len(strFilter) Then
            strFilter = strFilter & " AND "
        End If
        strFilter = strFilter & MobileSearch
    End If
    If Len(EmailSearch) Then
        If Len(strFilter) Then
            strFilter = strFilter & " AND "
        End If
        strFilter = strFilter & EmailSearch
    End If
    If Len(StudentsNameSearch) Then
        If Len(strFilter) Then
            strFilter = strFilter & " AND "
        End If
        strFilter = strFilter & StudentsNameSearch
    End If
    If Len(StudentsHomeRoomSearch) Then
        If Len(strFilter) Then
            strFilter = strFilter & " AND "
        End If
        strFilter = strFilter & StudentsHomeRoomSearch
    End If
    If Len(CanteenSearch) Then
        If Len(strFilter) Then
            strFilter = strFilter & " AND "
        End If
        strFilter = strFilter & CanteenSearch
    End If
    If Len(WorkPhSearch) Then
        If Len(strFilter) Then
            strFilter = strFilter & " AND "
        End If
        strFilter = strFilter & WorkPhSearch
    End If
    If Len(PandFSearch) Then
        If Len(strFilter) Then
            strFilter = strFilter & " AND "
        End If
        strFilter = strFilter & PandFSearch
    End If
    If Len(ArtCraftSearch) Then
        If Len(strFilter) Then
            strFilter = strFilter & " AND "
        End If
        strFilter = strFilter & ArtCraftSearch
    End If
    If Len(CanteenSearch) Then
        If Len(strFilter) Then
            strFilter = strFilter & " AND "
        End If
        strFilter = strFilter & CanteenSearch
    End If
    
    If Len(YearSearch) Then
        If Len(strFilter) Then
            strFilter = strFilter & " AND "
        End If
        strFilter = strFilter & YearSearch
    End If
    Form1.Volunteers.Filter = strFilter 'The program then filters the database to only have the records the user is after.
        If Not Form1.Volunteers.BOF Then
            Form1.Volunteers.MoveFirst 'Moves the recordset back to the beginning
        End If
    frmResults.lstResultID.Clear ' Clears the listboxes
    frmResults.lstResults.Clear  ' Ready for new data to be added
        Do Until Form1.Volunteers.EOF
            frmResults.lstResults.AddItem Form1.Volunteers.Fields("Fullname") 'Adds new data
            frmResults.lstResultID.AddItem Form1.Volunteers.Fields("ID")      'to list boxes
            Form1.Volunteers.MoveNext 'Moves recordset back to beginning
        Loop
    

    frmResults.Show 'Sows results
    strFilter = ""
    Form1.Volunteers.Filter = NullFilter 'Sets filter to nothing, ie all the database is available for viewing.
    Form1.Volunteers.Update
End Sub


Private Sub Command1_Click()
    MsgBox intAdvValue
End Sub

Private Sub Form_Load()
    NullFilter = "" 'defines NullFilter
    'Form1.Volunteers.Filter = ""
    'If Form1.Volunteers.Filter <> 0 Then
    '    Form1.Volunteers.Filter = NullFilter 'Makes sure there is no filter currently on the recordset
    'End If
End Sub
