VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   Caption         =   "String Tests"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInstr 
      Caption         =   "InStr"
      Height          =   495
      Left            =   3060
      TabIndex        =   18
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "testing"
      Height          =   315
      Left            =   2700
      TabIndex        =   17
      Top             =   2340
      Width           =   915
   End
   Begin VB.TextBox txtIter 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1980
      TabIndex        =   15
      Text            =   "1000"
      Top             =   2820
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Height          =   315
      Left            =   5100
      TabIndex        =   14
      Text            =   "Text7"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdStrLen 
      Caption         =   "String Len"
      Height          =   435
      Left            =   4740
      TabIndex        =   13
      Top             =   3300
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   4860
      TabIndex        =   12
      Text            =   "Text6"
      Top             =   3780
      Width           =   1035
   End
   Begin VB.CommandButton cmdConCat 
      Caption         =   "Concatenate"
      Height          =   495
      Left            =   180
      TabIndex        =   11
      Top             =   3480
      Width           =   1395
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Text            =   "EndString"
      Top             =   3780
      Width           =   1275
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Text            =   "BeginString"
      Top             =   3360
      Width           =   1275
   End
   Begin VB.TextBox txtPOS 
      Height          =   285
      Left            =   4200
      TabIndex        =   8
      Text            =   "1"
      Top             =   1440
      Width           =   435
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   4140
      TabIndex        =   5
      Text            =   "Testing String Rev."
      Top             =   2520
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FastRevString"
      Height          =   375
      Left            =   4140
      TabIndex        =   4
      Top             =   2100
      Width           =   1635
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4200
      TabIndex        =   3
      Text            =   "e"
      Top             =   840
      Width           =   435
   End
   Begin VB.CommandButton cmdPassString 
      Caption         =   "Pass String"
      Height          =   435
      Left            =   3840
      TabIndex        =   2
      Top             =   180
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   2595
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3555
   End
   Begin VB.CommandButton cmdScanStr 
      Caption         =   "  ScanStr >           a b c d e f  g h  i  1 2 3 4 5 6 7 8 9"
      Height          =   915
      Index           =   0
      Left            =   4680
      TabIndex        =   0
      Top             =   840
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Iterations:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1080
      TabIndex        =   16
      Top             =   2880
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "StartPOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   3780
      TabIndex        =   7
      Top             =   1200
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Char"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   3780
      TabIndex        =   6
      Top             =   900
      Width           =   405
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iASM As Long, iVB As Long, iCount As Long, s As String
Private iIter As Long, s1 As String, s2 As String, iPos As Long

Private Sub cmdConCat_Click()
    Dim s As String, iLen As Long
    iLen = Len(Text4.Text) + Len(Text5.Text)
    s = Space(iLen)
    Call StrConcatPTR(Text4.Text, Text5.Text, ByVal StrPtr(s))
    Text1.Text = s
End Sub

Private Sub cmdInstr_Click()
    
    s1$ = Text4.Text: s2$ = Text5.Text
    s = "Iterations: " & iIter & vbCrLf
    Screen.MousePointer = vbHourglass
    
    'ASM
    ProfileStart secIn
    For iCount = 0 To iIter
        iPos = InstrASM(s1$, s2$)
    Next
    ProfileStop secIn, secOut
    s = s & "   ASM: " & secOut & " sec, Ret: " & iPos & vbCrLf
    
    'VB
    ProfileStart secIn
    For iCount = 0 To iIter
        iPos = InStr(s1$, s2$)
    Next
    ProfileStop secIn, secOut
    s = s & "   VB: " & secOut & " sec, Ret: " & iPos & vbCrLf
    
    Text1.Text = s
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdPassString_Click()
    Text1.Text = PassString(Text7.Text) '("testing PassString")
End Sub

Private Sub cmdScanStr_Click(Index As Integer)
    Dim s As String, ret As Long, b As Byte, c As String
    If Text2.Text = "" Then Exit Sub
    s = "abcdefghi"
    c = Text2.Text
    b = Asc(c)
    ret = ScanString(StrPtr(s), Len(s), b, Val(txtPOS.Text))
    Text1.Text = s & vbCrLf & c & vbCrLf & ">" & ret & "<"
End Sub

Private Sub cmdStrLen_Click()
    'not that much of a difference in speed with this function...
    
    s = "Len of: """ & Text6.Text & """ is: " & StrLen(Text6.Text) & vbCrLf
     'ASM
    ProfileStart secIn
    For iCount = 0 To iIter
        iPos = StrLen(Text6.Text)
    Next
    ProfileStop secIn, secOut
    s = s & "   ASM: " & secOut & " sec" & vbCrLf & vbCrLf
    
    'vb
    s = s & "Len of: """ & Text6.Text & """ is: " & Len(Text6.Text) & vbCrLf
    ProfileStart secIn
    For iCount = 0 To iIter
        iPos = Len(Text6.Text)
    Next
    ProfileStop secIn, secOut
    s = s & "   VB: " & secOut & " sec"
    
    Text1.Text = s
End Sub

Private Sub Command1_Click()
    Dim sBuffer As String
    sBuffer = Text3.Text
    FastBStringReverse sBuffer
    Text1.Text = sBuffer
End Sub

Private Sub Form_Load()
    iIter = Val(txtIter.Text)
End Sub

Private Sub txtIter_Change()
    iIter = Val(txtIter.Text)
End Sub


Private Sub Command2_Click()
    Text1.Text = strTest(Text4.Text, Text5.Text)
End Sub
