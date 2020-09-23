VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Pausing the Build"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   1395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim i As Long
    i = Testing
    MsgBox "This is the return: " & i
End Sub
