VERSION 5.00
Begin VB.Form frmInterface 
   Caption         =   "Testing Class"
   ClientHeight    =   2250
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   2760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd2 
      Caption         =   "Class2"
      Height          =   435
      Left            =   540
      TabIndex        =   1
      Top             =   1380
      Width           =   1575
   End
   Begin VB.CommandButton cmdTesting 
      Caption         =   "Ret: 12345678"
      Height          =   435
      Left            =   540
      TabIndex        =   0
      Top             =   540
      Width           =   1575
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ct As CTest, ct2 As CTest

Private Sub cmd2_Click()
    Set ct2 = New CTest
    MsgBox ct2.ClsFunc, , "cls #2"
End Sub

Private Sub cmdTesting_Click()
    Set ct = New CTest
    MsgBox ct.ClsFunc, , "cls #1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ct = Nothing
    Set ct2 = Nothing
    Set frmInterface = Nothing
End Sub
