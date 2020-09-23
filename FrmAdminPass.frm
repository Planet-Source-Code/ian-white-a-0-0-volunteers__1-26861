VERSION 5.00
Begin VB.Form FrmAdminPass 
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
      Text            =   "Administrator"
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
Attribute VB_Name = "FrmAdminPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim varReturn As String * 50
Dim varConvert As Double
Dim varKey As Integer

Private Sub cmdCancel_Click()
    Unload Me 'This closes and unloads any information that was in the password fields.
End Sub

Private Sub cmdOk_Click()
       
    If txtUser.Text = "Administrator" Then 'If the username is correct then
        If txtPassword.Text = varFinalPass Then 'If the password equals the decrypted password then
            MsgBox "Admin Access Allowed.", , "Volunteers" 'Access allowed
            Dim i As Integer
                i = 0
                For i = 0 To 7
                    Form1.LabelPassChange(i).Visible = True 'Makes all the change password fields and buttons visible.
                Next i
                For i = 10 To 15
                    Form1.txtPassChange(i).Visible = True 'Makes all the change password fields and buttons visible.
                Next i
                For i = 20 To 21
                    Form1.cmdPassChange(i).Visible = True 'Makes all the change password fields and buttons visible.
                Next i
                
                Form1.txtFirstName.Locked = False '
                Form1.txtSurName.Locked = False   '
                Form1.txtCareer.Locked = False    '
                Form1.txtPhone.Locked = False     'These unlock all the textboxes in the centre
                Form1.txtWorkNo.Locked = False    'of the main form so that the user is able to
                Form1.txtMobile.Locked = False    'edit the info
                Form1.txtEmail.Locked = False     '
                Form1.txtStudName.Locked = False  '
                Form1.txtHomeRoom.Locked = False  '
                Form1.txtYear.Locked = False      '
                
                Unload Me 'unloads the form
        Else
            MsgBox "Incorrect Password, please try again.", vbInformation = vbOKOnly, "Volunteers" 'pops up if it is the incorrect password ie the password entered by the user does not match VarFinalPass (the decrypted password)
        End If
    Else
        MsgBox "Incorrect User Name, please try again.", vbInformation = vbOKOnly, "Volunteers" 'pops up if the user name does is incorrect.
    End If
    txtPassword.Text = ""
    Admin = True 'Admin is set to true, this is used in IF statements on other forms
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
