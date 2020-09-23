VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBSchool - Lesson 1 | Introduction to VB"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMsgBox 
      Caption         =   "Show MsgBox"
      Height          =   615
      Left            =   1980
      TabIndex        =   2
      ToolTipText     =   "It will show the message box"
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdShowOtherForm 
      Caption         =   "Show Other Form"
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "This button will show other form"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblLabel3 
      Alignment       =   1  'Right Justify
      Caption         =   "Third label"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   5055
   End
   Begin VB.Label lblLabel2 
      Alignment       =   2  'Center
      Caption         =   "Second label"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "This is the ToolTipText"
      Top             =   1920
      Width           =   5055
   End
   Begin VB.Label lblLabel1 
      Caption         =   "First label"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   5055
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      Caption         =   "Welcome to VBSchool!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This is a comment
Rem This is a comment too

Private Sub cmdMsgBox_Click()
' This is the basic test of an MsgBox() functin, you can use it
' When application needs to interact with user
MsgBox "Visual Basic 6 is the best", vbOKOnly, "VBSchool"
            ' Message                         ' Buttons    ' Title
End Sub

Private Sub cmdShowOtherForm_Click()
' This will show other form using Show()
frmForm.Show
' Object   ' Procedure
End Sub

Private Sub Form_Load()
' This code will execute when form will appear on screen
End Sub

Private Sub lblLabel2_Click()
' This is a subprogram that will execute when label will be pressed
' Subprogram must be between Sub and End Sub
End Sub
