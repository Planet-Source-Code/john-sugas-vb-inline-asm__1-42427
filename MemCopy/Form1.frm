VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "MemCopy"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "MemCopy Test"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   3900
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   8835
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Private iASM As Long, iVB As Long, iCount As Long, s As String, fVB As Boolean
Private iIter As Long, sTest As String, iPos As Long, ret As Long
Private bBufferFrom() As Byte, bBufferTo() As Byte, iSize As Long, sOutput As String


Private Sub Command2_Click()
    sOutput = ""
    Screen.MousePointer = vbHourglass
    
    ReDim bBufferTo(iSize)
    sOutput = sOutput & "Copy 5Mb using ASM MemCopyD.." & vbCrLf
    If Not fVB Then sOutput = sOutput & "This is running in VB IDE.... Must compile for this to work..." & vbCrLf
    ProfileStart secIn
         MemCopyD ByVal VarPtr(bBufferTo(0)), ByVal VarPtr(bBufferFrom(0)), ByVal iSize
    ProfileStop secIn, secOut
    sOutput = sOutput & secOut & " sec, B(0): " & _
            Chr(bBufferTo(0)) & "  B(2): " & Chr(bBufferTo(2)) & "  Last 3 bytes: " & _
           Chr(bBufferTo(iSize - 3)) & Chr(bBufferTo(iSize - 2)) & Chr(bBufferTo(iSize - 1)) & vbCrLf & vbCrLf

    ReDim bBufferTo(iSize) 'reset dest
    sOutput = sOutput & "Copy 5Mb using Optimized ASM MemCopyDOpz.." & vbCrLf
    If Not fVB Then sOutput = sOutput & "This is running in VB IDE.... Must compile for this to work..." & vbCrLf
    ProfileStart secIn
         MemCopyDOpz ByVal VarPtr(bBufferTo(0)), ByVal VarPtr(bBufferFrom(0)), ByVal iSize
    ProfileStop secIn, secOut
    sOutput = sOutput & secOut & " sec, B(0): " & _
            Chr(bBufferTo(0)) & "  B(2): " & Chr(bBufferTo(2)) & "  Last 3 bytes: " & _
           Chr(bBufferTo(iSize - 3)) & Chr(bBufferTo(iSize - 2)) & Chr(bBufferTo(iSize - 1)) & vbCrLf & vbCrLf



    ReDim bBufferTo(iSize) 'reset dest
    sOutput = sOutput & "Copy 5Mb using Api CopyMemory." & vbCrLf
    ProfileStart secIn
        CopyMemory ByVal VarPtr(bBufferTo(0)), ByVal VarPtr(bBufferFrom(0)), ByVal iSize
    ProfileStop secIn, secOut
    sOutput = sOutput & secOut & " sec, B(0): " & _
            Chr(bBufferTo(0)) & "  B(2): " & Chr(bBufferTo(2)) & "  Last 3 bytes: " & _
           Chr(bBufferTo(iSize - 3)) & Chr(bBufferTo(iSize - 2)) & Chr(bBufferTo(iSize - 1)) & vbCrLf & vbCrLf

    
    
    Text1.Text = sOutput
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    fVB = IsExe
    iSize = 5000000
    ReDim bBufferFrom(iSize)
    bBufferFrom(0) = Asc("J"): bBufferFrom(2) = Asc("S")
    bBufferFrom(iSize - 3) = Asc("E"): bBufferFrom(iSize - 2) = Asc("N"): bBufferFrom(iSize - 1) = Asc("D")

End Sub


