Attribute VB_Name = "MMemCopy"
Option Explicit


Public Function MemCopyD(ByVal Dest As Long, ByVal Source As Long, ByVal ln As Long) As Long

''Found on the net...JS  **(Had to switch the Des & Src around to match API. Wasn't
''working when I first fired it up and didn't notice it was "backward" from "usual"...
''From: SLH's seventh paper, The sting
'' This is where you take the gloves off and get serious, for only a small
'' increase in overhead, you copy four bytes at a time instead of one. You have
'' to solve the problem of byte lengths that are not equally divisable by 4 and
'' this is done by producing a hybrid function that does the major data
'' transfer in 4 byte chunks and cleans up the remaining bytes in one byte
'' chunks.
'' Big mover, [ movsd ] BURP !
'' ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~
'#ASM_START
'   Dest equ [ebp+8]  ;input args
'   Source equ [ebp+12]
'   ln equ [ebp+16]
'
'   lnth equ [ebp-4]    ;local vars
'   divd equ [ebp-8]
'   rmdr equ [ebp-12]
'
'    push ebp           ;save registers and set up stack
'    mov ebp, esp
'    sub esp,12
'    push ebx
'    push esi
'    push edi
'
'    cmp ln, 4           ; if under 4 bytes long
'    jl Tail             ; jump to label tail
'    mov eax, ln         ; copy length into eax
'    push eax            ; place a copy of eax on the stack
'    shr eax, 2          ; integer divide eax by 4
'    shl eax, 2          ; multiply eax by 4 to get dividend
'    mov divd, eax       ; copy it into variable
'    mov ecx, divd       ; copy variable into ecx
'    pop eax             ; retrieve length in eax off the stack
'    sub eax, ecx        ; subtract dividend from length to get remainder
'    mov rmdr, eax       ; copy remainder into variable
'    cld                 ; copy bytes forward
'    mov ecx, ln         ; put byte count in ecx
'    shr ecx, 2          ; divide by 4 for DWORD data size
'    mov esi, Source     ; copy source pointer into source index
'    mov edi, Dest       ; copy dest pointer into destination index
'    rep movsd         ; repeat while not zero, move string DWORD
'    mov ecx, rmdr       ; put remainder in ecx
'    jmp Over
' Tail:
'    mov ecx, ln         ; set counter if less than 4 bytes in length
'    mov esi, Source     ; copy source pointer into source index
'    mov edi, Dest       ; copy dest pointer into destination index
' Over:
'    rep movsb         ; copy remaining BYTES from source to dest
'    sub ln, ecx         ; calculate return value ( little use )
'    mov eax, ln          ; ' return bytes copied
'
'    pop edi
'    pop esi
'    pop ebx
'    mov esp, ebp
'    pop ebp
'  ret 12
'
'#ASM_END
'' ---------------------------------------------------------------------------
'' Whereas "small mover" had a data transfer rate of about 17 Meg/Sec, "Big
'' mover " clocks at about 42 Meg/Sec on a 166 Meg Pentium, figures that are"
'' only fantasies in high level languages. On a late model fast machine, you
'' will easily see over a hundred Meg/Sec.

End Function


Public Function MemCopyDOpz(ByVal Dest As Long, ByVal Source As Long, ByVal ln As Long) As Long

''A little bit of Optimize added....
'#ASM_START
'
'    push ebp           ;save registers
'    mov ebp, esp
'    push ebx
'    push esi
'    push edi
'
'    mov ecx, [ebp+16]         ; put byte count in ecx
'    mov esi, [ebp+12]     ; copy source pointer into source index
'    mov edi, [ebp+8]       ; copy dest pointer into destination index

'    cmp ecx, 4           ; if under 4 bytes long
'    jl Tail2             ; jump to label tail
'    mov eax, ecx         ; copy length into eax
'    push eax            ; place a copy of eax on the stack
'    shr eax, 2          ; integer divide eax by 4
'    shl eax, 2          ; multiply eax by 4 to get dividend
'    mov ecx, eax       ; copy it into variable
'    pop eax             ; retrieve length in eax off the stack
'    sub eax, ecx        ; subtract dividend from length to get remainder
'    mov ebx, eax       ; copy remainder into variable
'    cld                 ; copy bytes forward
'    shr ecx, 2          ; divide by 4 for DWORD data size
'    rep movsd         ; repeat while not zero, move string DWORD
'    mov ecx, ebx       ; put remainder in ecx
' Tail2:
'    rep movsb         ; copy remaining BYTES from source to dest
'   ; mov eax, [ebp+16]          ; ' return bytes copied
'
'    pop edi
'    pop esi
'    pop ebx
'    mov esp, ebp
'    pop ebp
'  ret 12
'
'#ASM_END
End Function

