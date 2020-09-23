VERSION 5.00
Begin VB.Form frmTestAsm1 
   Caption         =   "TestAsm1"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkByRef 
      Caption         =   "999555 ByRef"
      Height          =   435
      Left            =   1680
      TabIndex        =   27
      Top             =   1560
      Width           =   1035
   End
   Begin VB.CheckBox chkByVal 
      Caption         =   "333999 ByVal"
      Height          =   435
      Left            =   540
      TabIndex        =   26
      Top             =   1560
      Width           =   1035
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      TabIndex        =   24
      Text            =   "77777"
      Top             =   3180
      Width           =   915
   End
   Begin VB.CommandButton cmdDiv2 
      Caption         =   "/2"
      Height          =   435
      Left            =   1920
      TabIndex        =   23
      Top             =   3060
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   4395
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   1080
      Width           =   3315
   End
   Begin VB.CommandButton cmdRunRR 
      Caption         =   "Execute"
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Top             =   600
      Width           =   2115
   End
   Begin VB.TextBox txtIter 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4500
      TabIndex        =   13
      Text            =   "10000"
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Text            =   "3"
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      TabIndex        =   6
      Text            =   "4"
      Top             =   5400
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "num * 2^ pow"
      Height          =   435
      Left            =   480
      TabIndex        =   4
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simple Function Return"
      Height          =   495
      Left            =   540
      TabIndex        =   0
      Top             =   1020
      Width           =   2115
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Dividend:"
      Height          =   195
      Index           =   3
      Left            =   540
      TabIndex        =   25
      Top             =   2940
      Width           =   675
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Label6"
      Height          =   195
      Left            =   540
      TabIndex        =   22
      Top             =   2340
      Width           =   480
   End
   Begin VB.Shape Shape1 
      Height          =   1395
      Index           =   3
      Left            =   180
      Top             =   2220
      Width           =   2835
   End
   Begin VB.Shape LED 
      BorderWidth     =   2
      FillColor       =   &H00639869&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   660
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "RobertRayment's TestMC.VBP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3900
      TabIndex        =   19
      Top             =   300
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   5355
      Index           =   2
      Left            =   3240
      Top             =   180
      Width           =   3555
   End
   Begin VB.Label lblETVB 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5100
      TabIndex        =   18
      Top             =   6720
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "VB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   5
      Left            =   5520
      TabIndex        =   17
      Top             =   6180
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ASM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   4200
      TabIndex        =   16
      Top             =   6180
      Width           =   405
   End
   Begin VB.Label lblReturnVB 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5100
      TabIndex        =   15
      Top             =   6420
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(Try 2+ times after a compile...)"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   14
      Top             =   6660
      Width           =   2160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Iter:"
      Height          =   195
      Index           =   2
      Left            =   4200
      TabIndex        =   12
      Top             =   5820
      Width           =   270
   End
   Begin VB.Label lblETASM 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   6720
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ET:"
      Height          =   195
      Index           =   1
      Left            =   3540
      TabIndex        =   10
      Top             =   6780
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Arg2:"
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   9
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Arg1:"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   5160
      Width           =   495
   End
   Begin VB.Shape Shape1 
      Height          =   3255
      Index           =   1
      Left            =   120
      Top             =   3780
      Width           =   2955
   End
   Begin VB.Label Label4 
      Caption         =   "Label1"
      Height          =   1035
      Left            =   420
      TabIndex        =   5
      Top             =   3960
      Width           =   2115
   End
   Begin VB.Shape Shape1 
      Height          =   1875
      Index           =   0
      Left            =   240
      Top             =   180
      Width           =   2715
   End
   Begin VB.Label lblReturnASM 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   6420
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ret:"
      Height          =   195
      Index           =   0
      Left            =   3510
      TabIndex        =   2
      Top             =   6420
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   675
      Left            =   540
      TabIndex        =   1
      Top             =   300
      Width           =   2175
   End
End
Attribute VB_Name = "frmTestAsm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkByRef_Click()
    If chkByVal.Value = vbChecked Then chkByVal.Value = vbUnchecked
End Sub

Private Sub chkByVal_Click()
    If chkByRef.Value = vbChecked Then chkByRef.Value = vbUnchecked
End Sub

Private Sub cmdDiv2_Click()
    Dim iASM As Long, iVB As Long, iCount As Long, s As String
    Dim iIter As Long, iVal1 As Long, iVal2 As Long
        
    iIter = Val(txtIter)
    iVal1 = Val(Text5)
    Screen.MousePointer = vbHourglass
    
    'ASM
    ProfileStart secIn
    For iCount = 0 To iIter
        iASM = DivideBy2ByShifting(iVal1)
    Next
    ProfileStop secIn, secOut
    s = secOut & " sec"
    
    'VB
    iVal1 = Val(Text5)
    ProfileStart secIn
    For iCount = 0 To iIter
        iVB = DivideBy2Normally(iVal1)
    Next
    ProfileStop secIn, secOut
    lblETVB.Caption = secOut & " sec"
    
    lblReturnASM.Caption = iASM
    lblReturnVB.Caption = iVB
    lblETASM.Caption = s
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command1_Click()
    'clear unneeded lbl's
    lblETVB.Caption = ""
    lblReturnVB.Caption = ""
    lblETASM.Caption = ""
    'get asm result
    If chkByVal.Value = vbChecked Then 'byval
        lblReturnASM.Caption = TestingByVal(333999)
    ElseIf chkByRef.Value = vbChecked Then 'byref
        lblReturnASM.Caption = TestingByRef(999555)
    Else 'hardcoded return
        lblReturnASM.Caption = Testing
    End If
End Sub

Private Sub Command2_Click()
    Dim iASM As Long, iVB As Long, iCount As Long, s As String
    Dim iIter As Long, iVal1 As Long, iVal2 As Long
        
    'initialize vars so we don't in the first test...
    'getting false results on the first click...
    'appears to be the first time the exe is run after a compile...
    'after that the results are consistent...??? Even with restart of exe...
'    ProfileStart secIn
'    secIn = 0: secOut = 0
'    iCount = 0
'    iASM = 0: iVB = 0
'    ProfileStop secIn, secOut
    iIter = Val(txtIter)
    iVal1 = Val(Text1): iVal2 = Val(Text2)
    Screen.MousePointer = vbHourglass
    
    'ASM
    ProfileStart secIn
    For iCount = 0 To iIter
        iASM = Power2ASM(iVal1, iVal2)
    Next
    ProfileStop secIn, secOut
    s = secOut & " sec"
    
    'VB
    iVal1 = Val(Text1): iVal2 = Val(Text2)
    ProfileStart secIn
    For iCount = 0 To iIter
        iVB = Power2VB(iVal1, iVal2)
    Next
    ProfileStop secIn, secOut
    lblETVB.Caption = secOut & " sec"
    
    lblReturnASM.Caption = iASM
    lblReturnVB.Caption = iVB
    lblETASM.Caption = s
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdRunRR_Click()
    Dim TestArray() As Byte, TestArray2() As Byte, k As Long, i As Long
    Dim ALong1() As Long, ALong2() As Long, alen As Long
    Dim s As String, ret As Long, ptTA As Long, ptTA2 As Long
    
    Screen.MousePointer = vbHourglass
    LED.FillColor = &H3F841: LED.Refresh
    'USING EBX as index register
    s = "USING EBX as index register" & vbCrLf
    'Testpntr.asm
    'Move TestArray(2) to al -> res&
    ''InFile$ = "Testpntr.com"
    ''Loadmcode InFile$
    ReDim TestArray(2)
    TestArray(1) = 111
    TestArray(2) = 222
    ''ptmc& = VarPtr(InCode(1))
    ptTA& = VarPtr(TestArray(1))
    ''res& = CallWindowProc(ptmc&, ptTA&, 2&, 3&, 4&)
    ret = Testpntr(ptTA&)
    s = s & "  TestArray(2) Ret= " & ret& & vbCrLf & vbCrLf
    
    'TPt2Pt.asm
    'Move TestArray(1) -> TestArray2(1)
    ''InFile$ = "TPt2Pt.com"
    ''Loadmcode InFile$
    ReDim TestArray(2)
    ReDim TestArray2(2)
    TestArray(1) = 111
    TestArray2(1) = 0
    ''ptmc& = VarPtr(InCode(1))
    ptTA& = VarPtr(TestArray(1))
    ptTA2& = VarPtr(TestArray2(1))
    ''res& = CallWindowProc(ptmc&, ptTA&, ptTA2&, 3&, 4&)
    ret = TPt2Pt(ptTA&, ptTA2&)
    s = s & "  TestArray(1) to ret& " & ret& & vbCrLf
    s = s & "  TestArray(1) to TestArray2(1) " & TestArray2(1) & vbCrLf & vbCrLf
    
    '---------------------------------------------------------
    'USING ESI->EDI as index registers
    s = s & "USING ESI->EDI as index registers" & vbCrLf
    'TA12A2.asm
    'Move TestArray() -> TestArray2()
    ''InFile$ = "TA12A2.com"
    ''Loadmcode InFile$
    ReDim TestArray(2)
    ReDim TestArray2(2)
    TestArray(1) = 111
    TestArray(2) = 222
    TestArray2(1) = 0
    TestArray2(2) = 0
    ''ptmc& = VarPtr(InCode(1))
    ptTA& = VarPtr(TestArray(1))
    ptTA2& = VarPtr(TestArray2(1))
    ''res& = CallWindowProc(ptmc&, ptTA&, ptTA2&, 3&, 4&)
    ret = TA12A2(ptTA&, ptTA2&)
    s = s & "  TestArray2(1)= " & TestArray2(1) & vbCrLf
    s = s & "  TestArray2(2)= " & TestArray2(2) & vbCrLf & vbCrLf
    
    s = s & "COPY BYTE ARRAY TO BYTE ARRAY" & vbCrLf
    'TBA2BA.asm
    'Move TestArray() -> TestArray2()
    'using length of array
    ''InFile$ = "TBA2BA.com"
    ''Loadmcode InFile$
    ReDim TestArray(3)
    ReDim TestArray2(3)
    TestArray(1) = 11
    TestArray(2) = 22
    TestArray(3) = 33
    TestArray2(1) = 0
    TestArray2(2) = 0
    TestArray2(3) = 0
    ''ptmc& = VarPtr(InCode(1))
    ptTA& = VarPtr(TestArray(1))
    ptTA2& = VarPtr(TestArray2(1))
    alen& = 3
    ''res& = CallWindowProc(ptmc&, ptTA&, ptTA2&, alen&, 4&)
    ret = TBA2BA(ptTA&, ptTA2&, alen&)
    s = s & "  TestArray2(1)= " & TestArray2(1) & vbCrLf
    s = s & "  TestArray2(2)= " & TestArray2(2) & vbCrLf
    s = s & "  TestArray2(3)= " & TestArray2(3) & vbCrLf & vbCrLf
'GoTo Skip
    s = s & "COPY LONG ARRAY TO LONG ARRAY" & vbCrLf
    'Equivalent to API CopyMemory used for arrays
    'NB uses the same TBA2BA.asm
    'CPYARRAY.asm  In fact same as TBA2BA.asm
    'Move ALong1() -> ALong2()
    'using length of array
    ''InFile$ = "CPYARRAY.com"   'In fact same as TBA2BA.com
    ''Loadmcode InFile$
    ReDim ALong1(20000)
    ReDim ALong2(20000)
    'NB their length will be 80000 bytes
    ALong1(1) = 1
    ALong1(20000) = 20000
    ALong2(1) = 0
    ALong2(20000) = 0
    s = s & "  COPYING ARRAYS" & vbCrLf
    
    'VB method
    ''t! = Timer
    ProfileStart secIn
    For k = 1 To 1000
    For i = 1 To 20000
       ALong2(i) = ALong1(i)
    Next i
    Next k
    ProfileStop secIn, secOut
    s = s & "  VB took " & secOut & " secs" & vbCrLf
    
    'Machine code method
    ALong2(1) = 0
    ALong2(20000) = 0
    'ptmc& = VarPtr(InCode(1))
    ptTA& = VarPtr(ALong1(1))
    ptTA2& = VarPtr(ALong2(1))
    alen& = 80000
    
    ''t! = Timer
    ProfileStart secIn
    For k = 1 To 1000
       'res& = CallWindowProc(ptmc&, ptTA&, ptTA2&, alen&, 4&)
       ret = TBA2BA(ptTA&, ptTA2&, alen&)
    Next k
    ProfileStop secIn, secOut
    s = s & "  MCode took " & secOut & " secs" & vbCrLf
    'Test transfer
    s = s & "  ALong2(1)= " & ALong2(1) & vbCrLf
    s = s & "  ALong2(20000)= " & ALong2(20000) & vbCrLf
    Erase ALong1, ALong2
Skip:
    Text3.Text = s
    Screen.MousePointer = vbDefault
    LED.FillColor = &H639869
End Sub

Private Sub Form_Load()
    Me.Caption = "Run in the IDE, Then compare to the compiled .exe"
    Label1.Caption = "Simple Return Function:" & vbCrLf & vbCrLf _
                    & "     mov eax,12345678"
    
    Label4.Caption = "2 Arg. Function:" & vbCrLf & vbCrLf _
                    & "     mov eax, Arg1" & vbCrLf & "     mov ecx, Arg2" & vbCrLf _
                    & "     rcl eax, cl"
    
    Label6.Caption = "Divide by 2 Function:" & vbCrLf & "    Sar   eax, 1"
End Sub






