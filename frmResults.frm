VERSION 5.00
Begin VB.Form frmResults 
   BackColor       =   &H00289CFE&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search Results"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5640
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   4800
      Width           =   1695
   End
   Begin VB.ListBox lstResultID 
      Height          =   840
      Left            =   2400
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   1695
   End
   Begin VB.ListBox lstResults 
      BackColor       =   &H00400000&
      ForeColor       =   &H0000FFFF&
      Height          =   4545
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ListIndex As Integer

Private Sub cmdCancel_Click()
'    Form1.Volunteers.Filter = "" 'Makes all of database available. Undoes Filter.
    Unload Me 'This closes and unloads any information that was on the results form.
End Sub

Private Sub cmdOk_Click()
    On Error Resume Next
    lstResultID.ListIndex = lstResults.ListIndex 'This is to match up the name in one listbox, with an ID in another, this is done because I refer to the DB using ID no's
    Form1.Volunteers.MoveFirst 'Moves to the first entry of the DB
    Form1.Volunteers.Find "ID like '" & lstResultID.Text & "'" 'This finds the record in the Access DB that has the corresponding ID no.
    Call Form1.Form1Display 'Then displays it
    frmAdSearch.Hide
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    'This code below sets all the printer variables
    'and prints the results that are in the lstResults ListBox
    
    lstResults.ListIndex = 0
    Printer.Font = "arial"
    Printer.FontSize = 18
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.Print "Gilroy College - Volunteers Database - Search Results"
    Printer.Print ""
    Printer.Font = "Arial"
    Printer.FontUnderline = False
    Printer.FontSize = 12
    Printer.Print "Results: "
    For ListIndex = 0 To lstResults.ListCount - 1
        lstResults.ListIndex = ListIndex
        Printer.FontBold = False
        Printer.Print "",
        Printer.Print lstResults.Text
    Next ListIndex
    Printer.EndDoc
    lstResults.ListIndex = 0 'When it has finished printing, it will be at the last result, this simply moves it back up to the first.
End Sub

Sub lstResults_DblClick()
    On Error Resume Next
    
    lstResultID.ListIndex = lstResults.ListIndex 'This is to match up the name in one listbox, with an ID in another, this is done because I refer to the DB using ID no's
    Form1.Volunteers.MoveFirst 'Moves to first record in DB
    Form1.Volunteers.Find "ID like '" & lstResultID.Text & "'" 'Finds the one with the corresponding ID
    Call Form1.Form1Display 'Displays it
    frmAdSearch.Hide
    Unload Me 'Closes results window
    
End Sub
