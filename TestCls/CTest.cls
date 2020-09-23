VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function ClsFunc() As Long
'''The "_ClsFunc" variables are added by VB....

'#ASM_START
'  push ebp             ;Save EBP
'  mov ebp, esp         ; Move ESP into EBP so we can refer
'                       ;   to arguments on the stack
'  push ebx
'  push esi             ;save registers
'  push edi
'
' ;Start Add code Here
'
'  mov  DWORD PTR _ClsFunc2$[ebp], 12345678  ;return this number
'
' ;End Add code Here
'
'  mov     eax, DWORD PTR _ClsFunc1$[ebp]
'  mov     ecx, DWORD PTR _ClsFunc2$[ebp]
'  mov     DWORD PTR [eax], ecx
'
'  pop edi              ;restore registers
'  pop esi
'  pop ebx
'  mov esp, ebp         ;MOV/POP is much faster
'  pop ebp              ;on 486 and Pentium than Leave
'  ret 8                ; Return with 8 byte(long)
'#ASM_END

End Function

