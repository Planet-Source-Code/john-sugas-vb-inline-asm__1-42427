Attribute VB_Name = "MPic"
'Robert's external ASM source code converted to VbInLineASM.
'The code that is now not needed has been commented out with
'the ASM comment ";". Of course, Vb comments are needed to
'prevent compile errors.... Notice also-> No longer are the
'input arguments limited to 4.... No special settings are
'needed to compile this- TheParser func. will find the "#ASM_START"
'and do It's "Thing". Of course, the VbInLineASM addin must be loaded
'and hooked...JS
Option Explicit

Public Function CopyPic(ByVal long1 As Long, ByVal long2 As Long, ByVal long3 As Long) As Long
';;CopyPic.asm   For A386
';;Copies picture in pic1 to pic2
';;Out: res&=eax
'
';;VB
';;ptbmp1& = bmp1.bmBits
';;ptbmp2& = bmp2.bmBits
';;ReDim inparams&(6)
';;inparams&(1) = bmp1.bmWidth
';;inparams&(2) = bmp1.bmHeight
';;inparams&(3) = bmp1.bmWidthBytes
';;ptparams = VarPtr(inprams&(1))
';;res& = CallWindowProc(ptmc&, ptbmp1&, ptbmp2&, ptparams, 4&)
';;stack positions              8        12       16        20
'
';USE32

'#ASM_START
'long1 equ[ebp+8]    ;->pic1
'long2 equ[ebp+12]   ;->pic2
'long3 equ[ebp+16]   ;-> params
';long4 equ[ebp+20]   ;4&  Spare
'
'
'bmWidth       equ [ebp-4]
'bmHeight      equ [ebp-8]
'bmWidthBytes  equ [ebp-12]
'totBytes      equ [ebp-16]
'
'    push ebp        ;Arrange for ebp to point to stack params
'    mov ebp, esp
'    sub esp,16
'    push ebx
'    push esi
'    push edi
'
';Get and store bm params
'    mov ebx, long3
'    mov eax, [ebx]
'    mov bmWidth, eax
'    add ebx, 4    ;inc ebx, 4
'    mov eax, [ebx]
'    mov bmHeight, eax
'    mov edx, eax
'    add ebx, 4    ;inc ebx, 4
'    mov eax, [ebx]
'    mov bmWidthBytes, eax
'
'    mul edx
'    add ebx, 4    ;inc ebx, 4
'    mov totBytes, eax
'    mov ecx, bmWidthBytes
'    shr ecx,2           ;/4 for 4 byte moves
'
'    mov edx, ecx
'    mov esi,long1       ;ptr to pic1
'    mov edi,long2       ;ptr to pic2
'    ;add edi,totBytes
'    ;sub edi,bmWidthBytes   ;get to start of scan at top pic2
'    ;mov ebx,edi        ;save
';Set num Y lines
'    mov ecx, bmHeight
'Start:
'    push ecx
'    ;push ebx
'
'    mov ecx, edx
'    cld     ;ensure incr
'    rep movsd   ;[esi]->[edi] esi+4 edi+4 ecx-4
'
'    ;decrement di
'    ;pop ebx
'    ;sub ebx,bmWidthBytes
'    ;mov edi,ebx
'
'
'    pop ecx
'    dec ecx
'    jnz Start
'GETOUT:
'    pop edi
'    pop esi
'    pop ebx
'    mov esp, ebp
'    pop ebp
'    ret 12;16
'#ASM_END

End Function

Public Function FlipPic(ByVal long1 As Long, ByVal long2 As Long, ByVal long3 As Long) As Long

';FlipPic.asm   For A386
';Flips picture in pic1 to pic2
';Out: res&=eax
'
';VB
';ptbmp1& = bmp1.bmBits
';ptbmp2& = bmp2.bmBits
';ReDim inparams&(6)
';inparams&(1) = bmp1.bmWidth
';inparams&(2) = bmp1.bmHeight
';inparams&(3) = bmp1.bmWidthBytes
';ptparams = VarPtr(inprams&(1))
';res& = CallWindowProc(ptmc&, ptbmp1&, ptbmp2&, ptparams, 4&)
';stack positions              8        12       16        20
'
';USE32

'#ASM_START
'long1  equ[ebp+8]    ;->pic1
'long2  equ[ebp+12]   ;->pic2
'long3  equ[ebp+16]   ;-> params
';long4  equ[ebp+20]   ;4&  Spare
'
'bmWidth       equ[ebp-4]
'bmHeight      equ[ebp-8]
'bmWidthBytes  equ[ebp-12]
'totBytes      equ[ebp-16]
'
'    push ebp        ;Arrange for ebp to point to stack params
'    mov ebp, esp
'    sub esp,16
'    push ebx
'    push esi
'    push edi
'
';Get and store bm params
'    mov ebx, long3
'    mov eax, [ebx]
'    mov bmWidth, eax
'    add ebx, 4    ;inc ebx, 4
'    mov eax, [ebx]
'    mov bmHeight, eax
'    mov edx, eax
'    add ebx, 4    ;inc ebx, 4
'    mov eax, [ebx]
'    mov bmWidthBytes, eax
'
'    mul edx
'    add ebx, 4    ;inc ebx, 4
'    mov totBytes, eax
'    mov ecx, bmWidthBytes
'    shr ecx,2           ;/4 for 4 byte moves
'
'    mov edx, ecx
'    mov esi,long1       ;ptr to pic1
'    mov edi,long2       ;ptr to pic2
'    Add edi, totBytes
'    sub edi,bmWidthBytes    ;get to start of scan at top pic2
'    mov ebx,edi     ;save
';Set num Y lines
'    mov ecx, bmHeight
'Start2:
'    push ecx
'    push ebx
'
'    mov ecx, edx
'    cld     ;ensure incr
'    rep movsd   ;[esi]->[edi] esi+4 edi+4 ecx-4
'
'    ;decrement di
'    pop ebx
'    sub ebx,bmWidthBytes
'    mov edi, ebx
'
'
'    pop ecx
'    dec ecx
'    jnz Start2
'GETOUT2:
'    pop edi
'    pop esi
'    pop ebx
'    mov esp, ebp
'    pop ebp
'    ret 12;16

'#ASM_END

End Function
