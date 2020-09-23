Attribute VB_Name = "MAsmCode"
'This module has both ASM func. and VB func. Noticed that the VB
'functions are very simple and don't include the "On Error" statement.
'More experimentation will be needed to find out the exact stituations
'that cause the Structured Error Handling to be added....JS

Option Explicit
Public Function Testing() As Long
'#ASM_START
'  push ebp             ;Save EBP
'  mov ebp, esp         ; Move ESP into EBP so we can refer
'                       ;   to arguments on the stack
'  mov eax, 12345678    ;Put a number in EAX
'
'  mov esp, ebp         ;MOV/POP is much faster
'  pop ebp              ;on 486 and Pentium than Leave
'  ret
'#ASM_END
End Function
Public Function TestingByVal(ByVal iInput As Long) As Long
'#ASM_START
'  push ebp             ;Save EBP
'  mov ebp, esp         ; Move ESP into EBP so we can refer
'                       ;   to arguments on the stack
'  mov eax, [ebp+8]    ;Put argument in EAX
'
'  mov esp, ebp         ;MOV/POP is much faster
'  pop ebp              ;on 486 and Pentium than Leave
'  ret
'#ASM_END
End Function
Public Function TestingByRef(iInput As Long) As Long
'#ASM_START
'  push ebp             ;Save EBP
'  mov ebp, esp         ;Move ESP into EBP so we can refer to arguments on the stack
'  push ebx             ;
'
'  mov ebx, [ebp+8]     ;put argument ref. in ebx
'  mov eax, [ebx]       ;Put the number in EAX
'
'  pop ebx
'  mov esp, ebp         ;MOV/POP is much faster
'  pop ebp              ;on 486 and Pentium than Leave
'  ret
'#ASM_END
End Function
Public Function DivideBy2Normally(ByVal lngDividend As Long) As Long
    'Don't have to call a function for this easy op but we need to be honest
    'and include the time it takes to call it like the ASM fucntion...
    DivideBy2Normally = lngDividend / 2
End Function
Public Function DivideBy2ByShifting(ByVal lngDividend As Long) As Long
'#ASM_START
'
'  lngDividend equ[ebp+8]
'
'  push ebp             ;Save EBP
'  mov ebp, esp         ; Move ESP into EBP so we can refer
'                       ;   to arguments on the stack
'  mov eax, lngDividend    ;Put  number in EAX
'  Sar   eax, 1         ;shift right 1 place,-> /2
'
'  mov esp, ebp         ;MOV/POP is much faster
'  pop ebp              ;on 486 and Pentium than Leave
'  ret 4
'#ASM_END
End Function

Public Function Testpntr(ByVal long1 As Long) As Long
';From Robert Rayment's TESTMC.vbp
';Testpntr.asm
';Test res& = CallWindowProc(ptr_m/c, long1,Long2,Long3,Long4)
';where Long1 is pointer to TestArray(1),(2)
';This gets param into eax and to res&
';USE32

'#ASM_START
'    long1 equ[ebp+8]
';    long2 equ[ebp+12]   ;comment these out... not needed
';    long3 equ[ebp+16]
';    long4 equ[ebp+20]
'
'    push ebp
'    mov ebp, esp
'    push ebx        ;EBX must be saved for EXE to work
'
'    mov ebx,long1   ;pntr to TestArray(1)
'    xor eax,eax     ;zero eax coz res& is long result & we want a byte
'            ;[ebx]  has a default seg reg [ds:ebx]
'    mov al,[ebx]    ;load a byte (TestArray(1))into al
'    mov al,[ebx+1] ;load a byte (TestArray(2))into al
'
'    pop ebx
'    mov esp, ebp
'    pop ebp
'    ret 4  ;16
'#ASM_END
End Function
Public Function TPt2Pt(ByVal long1 As Long, ByVal long2 As Long) As Long
';TPt2Pt.asm
';Test res& = CallWindowProc(VarPtr(InCode(1)), VarPtr(TestArray(1)),
';                   8
';VarPtr(TestArray2(1))&, 3&, 4&)
';   12                   16   20
';where Long1 is pointer to TestArray(1),(2)
';and Long2 is pointer to TestArray2(1),(2)
';Move value in TestArray(1) to TestArray2(1)
';USE32

'#ASM_START
'    long1 equ[ebp+8]
'    long2 equ[ebp+12]
';    long3 equ[ebp+16]
';    long4 equ[ebp+20]
'
'    push ebp
'    mov ebp, esp
'    push ebx
'
'    mov ebx,long1   ;pntr to TestArray(1)
'    xor eax,eax     ;zero eax coz res& is long result & we want a byte
'            ;[ebx]  has a default seg reg [eds:ebx]
'    mov al,[ebx]    ;load a byte (TestArray(1)) into al
'
'    mov ebx,long2   ;pntr to TestArray2(1)
'    mov [ebx],al    ;put al into TestArray2(1)
'            ;res& will also contain this byte
'    pop ebx
'    mov esp, ebp
'    pop ebp
';    ret 16  returning 16 gives access violation....
'   ret 8
'#ASM_END
End Function
Public Function TA12A2(ByVal long1 As Long, ByVal long2 As Long) As Long
';TA12A2.asm
';Test res& = CallWindowProc(VarPtr(InCode(1)), VarPtr(TestArray(1)),
';                   8
';VarPtr(TestArray2(1))&, 3&, 4&)
';   12                   16   20
';where Long1 is pointer to TestArray(1),(2)
';and Long2 is pointer to TestArray2(1),(2)
';Move all values in TestArray() to TestArray2()
';USE32

'#ASM_START
'    long1 equ[ebp+8]
'    long2 equ[ebp+12]
';    long3 equ[ebp+16]
';    long4 equ[ebp+20]
'
'    push ebp
'    mov ebp, esp
';    push edi, esi ->compile error with the comma
'   push esi
'   push edi
'
'    mov esi,long1   ;pntr to TestArray()
'    mov edi,long2   ;pntr to TestArray2()
'    cld     ;ensure incr
'    movsb       ;byte in long1->byte in long2 esi+1,edi+1
'    movsb
'
';    pop esi, edi -> compile error here also
'   pop edi
'   pop esi
'    mov esp, ebp
'    pop ebp
'    ret 8  ;16
'#ASM_END
End Function
Public Function TBA2BA(ByVal long1 As Long, ByVal long2 As Long, ByVal long3 As Long) As Long
';TBA2BA.asm
';Test res& = CallWindowProc(VarPtr(InCode(1)), VarPtr(TestArray(1)),
';                   8
';VarPtr(TestArray2(1))&, ByteArrayLength, 4&)
';   12                   16                        20
';where Long1 is pointer to TestArray(1),(2)
';and Long2 is pointer to TestArray2(1),(2)
';Move all values in TestArray() to TestArray2()
';USE32

'#ASM_START
'    long1 equ[ebp+8]
'    long2 equ[ebp+12]
'    long3 equ[ebp+16]
';    long4 equ[ebp+20]
'
'    push ebp
'    mov ebp, esp
';    push edi,esi,ecx    ;save regs
'   push ecx
'   push esi
'   push edi
'    push ds
'
'    mov esi,long1   ;pntr to TestArray()
'    mov edi,long2   ;pntr to TestArray2()
'    mov ecx,long3   ;length of byte arrays
'    cld     ;ensure incr
'    rep movsb       ;byte in long1->byte in long2 si+1,di+1 until ecx=0
'
'    pop ds
';    pop ecx,esi,edi ;restore regs
'   pop edi
'   pop esi
'   pop ecx
'    mov esp, ebp
'    pop ebp
';    ret 16 -> won't return last element (33)
'   ret 12
'#ASM_END
End Function
Public Function Power2VB(Num As Long, Pow As Long) As Long
    Power2VB = Num * 2 ^ Pow
End Function
Public Function Power2ASM(ByVal Num As Long, ByVal Power As Long) As Long

'#ASM_START
'  push ebp
'  mov ebp, esp
'
'  mov eax, [ebp+8] ; Get first argument
'  mov ecx, [ebp+12] ; Get second argument
'  rcl eax, cl     ;shl eax, cl     ; EAX = EAX * ( 2 to the power of CL )

'  mov esp, ebp         ;MOV/POP is much faster
'  pop ebp              ;on 486 and Pentium than Leave
'  ret 8
'#ASM_END

''================================================
''; POWER.ASM  'from MSDN
''; Compute the power of an integer
'';
''       PUBLIC _power2
''_TEXT SEGMENT WORD PUBLIC 'CODE'
''_power2 PROC
''
''        push ebp        ; Save EBP
''        mov ebp, esp    ; Move ESP into EBP so we can refer
''                        ;   to arguments on the stack
''        mov eax, [ebp+4] ; Get first argument
''        mov ecx, [ebp+6] ; Get second argument
''        shl eax, cl     ; EAX = EAX * ( 2 ^ CL )
''        pop ebp         ; Restore EBP
''        ret             ; Return with sum in EAX
''
''_power2 ENDP
''_TEXT   ENDS
''        End
''=============================================
''/* POWER2.C */
''#include <stdio.h>
''
''int power2( int num, int power );
''
''void main(void)
''{
''   printf( "3 times 2 to the power of 5 is %d\n", \
''           power2( 3, 5) );
''}
''int power2( int num, int power )
''{
''   __asm
''   {
''      mov eax, num    ; Get first argument
''      mov ecx, power  ; Get second argument
''      shl eax, cl     ; EAX = EAX * ( 2 to the power of CL )
''   }
''   /* Return with result in EAX */
''}

End Function
