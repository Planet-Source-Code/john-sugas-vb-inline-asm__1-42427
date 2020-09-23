VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000C0&
   Caption         =   "Dissecting a VB Function"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Call That Function"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   660
      Width           =   1995
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MsgBox "DissectThis(25) returns: " & DissectThis(25)
End Sub
