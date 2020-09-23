VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   757
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoop 
      Caption         =   "Loop"
      Height          =   495
      Left            =   300
      TabIndex        =   4
      Top             =   2460
      Width           =   675
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   495
      Left            =   270
      TabIndex        =   3
      Top             =   1755
      Width           =   675
   End
   Begin VB.CommandButton cmdFlip 
      Caption         =   "Flip"
      Height          =   495
      Left            =   255
      TabIndex        =   2
      Top             =   1005
      Width           =   675
   End
   Begin VB.PictureBox pic2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1755
      Left            =   4710
      ScaleHeight     =   117
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   112
      TabIndex        =   1
      Top             =   585
      Width           =   1680
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   1260
      ScaleHeight     =   172
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   172
      TabIndex        =   0
      Top             =   195
      Width           =   2580
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ported to VbInLineASM as a Test Proj. by John Sugas 2002"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   3720
      TabIndex        =   5
      Top             =   5220
      Width           =   3405
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FlipCopy   by Robert Rayment
'Machine code & Direct picture box addressing
'NB ONLY for 24-bit, Equal sized, Mod 4 width bitmaps
'See PSC   -  Non-linear systems for dll way

Option Base 1
DefInt A-T
DefSng U-Z

'To fill BITMAP structure
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" _
(ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Type BITMAP
   bmType As Long              ' Type of bitmap
   bmWidth As Long             ' Bitmap width in pixel
   bmHeight As Long            ' Bitmap height in pixel
   bmWidthBytes As Long        ' Bitmap width in bytes, like the pitch of a DirectDraw surface
   bmPlanes As Integer         ' Color depth of bitmap
   bmBitsPixel As Integer      ' Bits per pixel, must be 24 or we will terminate
   bmBits As Long              ' This is the pointer to our bitmap data  !!!
End Type

'NB PICTURE STORED IN MEMORY UPSIDE DOWN
'WITH INCREASING MEMORY GOING UP THE PICTURE
'bmp.bmBits points to the bottom left of the picture

Dim bmp1 As BITMAP
Dim bmp2 As BITMAP

'Dim InCode() As Byte  'Dim InCode2() As Byte
Dim inparams&()
Dim Done As Boolean, NotLooping As Boolean


Private Sub cmdFlip_Click()
Done = False

''ptmc& = VarPtr(InCode(1))
ptbmp1& = bmp1.bmBits
ptbmp2& = bmp2.bmBits
'Set up parameter list
ReDim inparams&(6)        '***Not sure why you have 6 here Robert??? 3 will work....
inparams&(1) = bmp1.bmWidth
inparams&(2) = bmp1.bmHeight
inparams&(3) = bmp1.bmWidthBytes
ptinparams& = VarPtr(inparams&(1))

Call FlipPic(ptbmp1&, ptbmp2&, ptinparams&)
''res& = CallWindowProc(ptmc&, ptbmp1&, ptbmp2&, ptinparams&, 4&)
pic2.Refresh
NotLooping = True
End Sub

Private Sub cmdCopy_Click()
Done = False

''ptmc& = VarPtr(InCode2(1))
ptbmp1& = bmp1.bmBits
ptbmp2& = bmp2.bmBits
'Set up parameter list
ReDim inparams&(6)
inparams&(1) = bmp1.bmWidth
inparams&(2) = bmp1.bmHeight
inparams&(3) = bmp1.bmWidthBytes
ptinparams& = VarPtr(inparams&(1))

Call CopyPic(ptbmp1&, ptbmp2&, ptinparams&)
''res& = CallWindowProc(ptmc&, ptbmp1&, ptbmp2&, ptinparams&, 4&)
pic2.Refresh
NotLooping = True
End Sub

Private Sub cmdLoop_Click()
Done = True
NotLooping = False

ptbmp1& = bmp1.bmBits
ptbmp2& = bmp2.bmBits
'Set up parameter list
ReDim inparams&(6)
inparams&(1) = bmp1.bmWidth
inparams&(2) = bmp1.bmHeight
inparams&(3) = bmp1.bmWidthBytes
ptinparams& = VarPtr(inparams&(1))

'ptmc& = VarPtr(InCode(1))
'ptmc2& = VarPtr(InCode2(1))

Do
Call FlipPic(ptbmp1&, ptbmp2&, ptinparams&)
'res& = CallWindowProc(ptmc&, ptbmp1&, ptbmp2&, ptinparams&, 4&)
pic2.Refresh
DoEvents
Call CopyPic(ptbmp1&, ptbmp2&, ptinparams&)
'res& = CallWindowProc(ptmc2&, ptbmp1&, ptbmp2&, ptinparams&, 4&)
pic2.Refresh
DoEvents
Loop While Done
NotLooping = True
End Sub

Private Sub Form_Load()
Form1.Caption = "Machine code demonstration - Flip Copy  -  by Robert Rayment"
Form1.Top = 1000
Form1.Left = 1
Done = True
With pic1
.Width = 300
.Height = 300
.Picture = LoadPicture(App.Path & "/MC.jpg")
.Refresh
End With

With pic2
.Top = pic1.Top
.Left = pic1.Left + pic1.Width + 4
.Width = 300
.Height = 300
.Picture = LoadPicture(App.Path & "/MC.jpg")
.Refresh
End With

GetObjectAPI pic1.Picture, Len(bmp1), bmp1
GetObjectAPI pic2.Picture, Len(bmp2), bmp2
Show
Refresh

'InFile$ = "FlipPic.com"
'Loadmcode (InFile$)

'InFile$ = "CopyPic.com"
'Loadmcode2 (InFile$)

NotLooping = True  'needed only if we exit without playing.....
End Sub

'Private Sub Loadmcode(InFile$)
'Open InFile$ For Binary As #1
'MCSize& = LOF(1)
'ReDim InCode(MCSize&)
'Get #1, , InCode
'Close #1
'End Sub
'Private Sub Loadmcode2(InFile$)
'Open InFile$ For Binary As #1
'MCSize& = LOF(1)
'ReDim InCode2(MCSize&)
'Get #1, , InCode2
'Close #1
'End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Done = False
    DoEvents
    If Not NotLooping Then Cancel = True 'prevent "memory can't be read" crash if we exit
    Set Form1 = Nothing                 'and the loop is still running...
End Sub
