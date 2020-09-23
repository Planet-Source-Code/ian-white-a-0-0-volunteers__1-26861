VERSION 5.00
Begin VB.Form frmTeacherPass 
   BackColor       =   &H00400000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Please Enter Password"
   ClientHeight    =   1455
   ClientLeft      =   2325
   ClientTop       =   1230
   ClientWidth     =   3105
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   "Teacher"
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pass:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmTeacherPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim varReturn2 As String * 50
Dim varConvert2 As Double
Dim varKey2 As Integer

Private Sub cmdCancel_Click()
    Unload Me 'This closes and unloads any information that was in the password fields.
End Sub

Private Sub cmdOk_Click()
       
    If txtUser.Text = "Teacher" Then 'If the username is correct then
        If txtPassword.Text = varFinalPass2 Then 'If the password equals the decrypted password then
            MsgBox "Teacher Access Allowed.", , "Volunteers" 'Access allowed
                
                User = False
                Teacher = True
                Unload Me 'unloads the form
        Else
            MsgBox "Incorrect Password, please try again.", vbInformation = vbOKOnly, "Volunteers" 'pops up if it is the incorrect password ie the password entered by the user does not match varFinalPass2 (the decrypted password)
        End If
    Else
        MsgBox "Incorrect User Name, please try again.", vbInformation = vbOKOnly, "Volunteers" 'pops up if the user name does is incorrect.
    End If
    txtPassword.Text = ""
    Teacher = True
    Admin = False 'Admin is set to true, this is used in IF statements on other forms
    User = False 'User is set to false, this is used in IF statements on other forms
    
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdOk_Click 'If they push enter, then it calls the OK button
    End If
End Sub


Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdOk_Click 'If they push enter, then it calls the OK button
    End If
End Sub

