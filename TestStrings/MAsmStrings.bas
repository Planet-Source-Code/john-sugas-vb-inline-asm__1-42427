Attribute VB_Name = "MAsmStrings"
Option Explicit
'VB does all kinds of extra calls when returning a string function. I couldn't get
'anything to work in that area since all I got was crashes. Calls were being made
'to system dlls on return to the function. I spent a little time tracing them but
'quickly lost interest in being flopped around all over the place. In short, passing
'in a string reference is easier than returning a string function.... The code I
'was using to test this is at the bottom of this module in case anybody wants
'to try their skills...JS

Public Function StrLen(sIn As String) As Long
';This is working. Returns the length of the input string. JS
'#ASM_START
'
'    push ebp
'    mov ebp, esp
'    push ebx
'    push esi
'    push edi
'
'    mov eax,  [ebp+8]
'    mov ecx, [eax]
'    xor eax, eax
'    mov ax, [ecx-4]
'    cdq
'    sub eax, edx
'    sar eax, 1
'
'    pop edi
'    pop esi
'    pop ebx
'    mov esp, ebp
'    pop ebp
'  ret 4
'
'#ASM_END
End Function

Public Sub FastBStringReverse(TargetString As String)
''this code from compiler controller vbp... it is C++ generated asm. FastString.asm
''I added the local vars. to try and get a better idea what was going on with strings.
''Of course, it would be faster to just reference the [ebp] locations in the code...
''It can probably be optimized in a few other places.... But, it is working...JS

'#ASM_START
' BasicString equ [ebp+8]
' posLeft equ [ebp-4]
' posRight equ [ebp-8]
' Swap equ [ebp-12]
'
'    push ebp
'    mov ebp, esp
'    sub esp, 76                 ; 0000004cH
'    push ebx
'    push esi
'    push edi

'    lea edi, [ebp-76]
'    mov ecx, 19                 ; 00000013H
'    mov eax, -858993460             ; ccccccccH
'    rep stosd
'
'    mov  posLeft, 0
'    mov eax,  BasicString
'    mov ecx, [eax]
'    xor eax, eax
'    mov ax, [ecx-4]
'    cdq
'    sub eax, edx
'    sar eax, 1
'    sub eax, 1
'    mov posRight, eax
'    jmp $L222
'$L223:
'    mov edx, posLeft
'    Add edx, 1
'    mov posLeft, edx
'    mov eax, posRight
'    sub eax, 1
'    mov posRight, eax
'$L222:
'    mov ecx, posLeft
'    cmp ecx, posRight
'    jae $L224
'
'    mov edx, BasicString
'    mov eax, [edx]
'    mov ecx, posLeft
'    mov dx, [eax+ecx*2]
'    mov Swap, dx
'
'    mov eax, BasicString
'    mov ecx, [eax]
'    mov edx, BasicString
'    mov eax, [edx]
'    mov edx, posLeft
'    mov esi, posRight
'    mov cx, [ecx+esi*2]
'    mov [eax+edx*2], cx
'
'    mov edx, BasicString
'    mov eax, [edx]
'    mov ecx, posRight
'    mov dx, Swap
'    mov [eax+ecx*2], dx
'
'    jmp $L223
'$L224:
'
'    pop edi
'    pop esi
'    pop ebx
'    mov esp, ebp
'    pop ebp
'    ret 4
'#ASM_END
End Sub
Public Function PassString(sIn As String) As String
'#ASM_START
'
'  ;nothing needed, ret address is pointing to argument...
'  ret 4
'
'#ASM_END
End Function

Function ScanString(ByVal lpStrng As Long, _
                     ByVal lnStr As Long, _
                     ByVal Char As Byte, _
                     ByVal StartPos As Long) As Long
''found this chunk of code on the net. Similar to vb's Instr function.
''currently not working properly... maybe cuz of UniCode strings, and this is 4 Ansii.....
''Update... Got it working. Had to modify for VB unicode strings...JS

'#ASM_START
'lpStrng equ[ebp+8]      ;1st Argument
'lnStr equ[ebp+12]     ;2nd Argument
'Char equ[ebp+16]  ;byte but gets passed in as a long
'StartPos equ[ebp+20]
'
'    push  ebp
'    mov ebp, esp
'    sub esp, 16
'    push ebx
'    push esi
'    push edi
';int 3
';    inc lnStr     ;**compile err     ; needed so last char is compared
'    add lnStr, 1    ;inc 1 by adding
'    mov eax, lnStr              ; copy length into eax
'    sub StartPos, 1                 ;start pos is actual back 1
'    sub eax, StartPos               ; shorten length by start pos
'    mov lnStr, eax              ; copy result into length
'    mov eax, lpStrng            ; set starting offset in string by
'mov ebx,StartPos ;copy StartPos
'shl ebx, 1  ;multiply by 2 for Unicode
'    add eax, ebx               ; adding the starting position to it
'    mov lpStrng, eax            ; put sum back into variable
'    cld                         ; scan forward in string
'    xor eax,eax
'    mov al, Char                ; copy "search" character into al
'    mov ecx, lnStr              ; set maximum count in ecx
'    mov edi, lpStrng            ; copy pointer into destination index
'    repne scasw                 ; repeat if not equal, scan string BYTE
'    cmp ecx, 0                  ; if no matches found
'    je  zero                    ; jump to zero
'
'    sub lnStr, ecx              ; subtract char pos from string len
'    mov eax, lnStr              ; put value in eax
'  ;  add eax, StartPos               ; add starting pos to it
'  ;  mov lnStr, eax              ; put result back into value
'    jmp TheEnd
' zero:
'    mov lnStr, 0                ; set return value to zero if no match
' TheEnd:
'    mov eax, lnStr

'    pop edi
'    pop esi
'    pop ebx
'    mov esp, ebp
'    pop ebp
'  ret 16
'
'#ASM_END
End Function

Public Function InstrASM(sIn As String, sFind As String) As Long
'';Working now...Doesn't appear to have any advantage over VB Instr with
'';1000 rep timing test....JS

'#ASM_START
'sIn equ[ebp+8]      ;1st Argument
'sFind equ[ebp+12]     ;2nd Argument
'
'len1 equ [ebp-4]   ;len str 1
'len2 equ [ebp-8]   ;len str 2
'ptr2 equ [ebp-12]
'
'
'    push    ebp        ;save registers and setup stack
'    mov ebp, esp
'    sub esp, 16
'    push ebx
'    push edx
'    push esi
'    push edi

'   mov esi, sIn
'   mov esi, [esi]      ;dereference ptr1
'   mov eax, [esi-4]    ;len 1
'   mov len1, eax

'   mov edi, sFind
'   mov edi, [edi]      ;dereference ptr2
'   mov ptr2, edi
'   mov ecx, [edi-4]    ;len 2
'   mov len2, ecx

'   cmp len1, ecx       ;if len2 > len1 then exit
'   jl NotFnd

'   cld
'   mov ebx, len1
'   sub ebx, len2       ;# of reps needed = (len1-len2)/2+1
'   xor edx, edx
'   shr ebx, 1
'   sub esi, 2          ;need to inc in loop so set back
'   add ebx, 3          ;add 1 extra for first decr. + 2 bytes for last char.
'   mov al, byte ptr[edi] ;first char of search str.
'NextLoc:
'   dec ebx             ;loop index
'   je NotFnd           ;not found
'   inc edx             ;loc count
'   add esi, 2          ;loc in string
'   cmp al, byte ptr[esi] ;cmp first char of search str. with
'                         ;each char in target till a match
'   jnz NextLoc

'   push esi            ;save current loc
'   mov edi, ptr2       ;search string loc
'   mov ecx, len2       ;len search str.
'   rep cmpsb           ;  compare string
'   pop esi
'   je  Done            ;found jump
'   jmp NextLoc         ;next try

'NotFnd:
'   xor edx, edx        ;ret 0

'Done:
'   mov eax, edx        ;ret loc of match
'    pop edi
'    pop esi
'    pop edx
'    pop ebx
'    mov esp, ebp
'    pop ebp
'  ret 8
'
'#ASM_END

    
End Function

Public Function StrConcatPTR(s1 As String, s2 As String, ByVal DestStrPtr As Long) As Long
'working finally...JS

'#ASM_START
'
'    push    ebp
'    mov ebp, esp
'    sub esp, 76
'    push ebx
'    push edx
'    push esi
'    push edi

'    mov esi, [ebp+8]   ;load str 1
'    mov esi, [esi]     ;dereference
'    cld                ;direction forward

'    mov ecx, 2
'    mov eax,  [ebp+8]  ;get len1
'StrLen:
'    mov ebx, [eax]     ;string len routine
'    xor eax, eax
'    mov ax, [ebx-4]
'    cdq
'    sub eax, edx
'    sar eax, 1
'    dec ecx
'    jecxz GetStrings   ;done getting both lengths
'    push eax           ;save len1
'    mov eax,  [ebp+12] ;load str 2
'    jmp StrLen         ;get len2
'GetStrings:

'    mov ebx, eax       ;len2 in ebx
'    pop eax            ;len1 in eax

'    mov edx, [ebp+16]  ;return arg loc
'    lea edi, [edx]     ;ret ptr loc
'    mov ecx, eax       ;len 1
'    rep movsw          ;move string 1

'    mov esi, [ebp+12]  ;load str 2
'    mov esi, [esi]     ;dereference
'    mov ecx, ebx       ;len 2
'    rep movsw          ;move string 2

'    xor eax,eax        ;zero eax
'    mov edi, eax       ;null terminators

'    lea eax, [edx]     ;return ptr in eax
'
'    pop edi
'    pop esi
'    pop edx
'    pop ebx
'    mov esp, ebp
'    pop ebp
'  ret 12
'
'#ASM_END
End Function

Public Function strTest(s1 As String, s2 As String) As String
'using this function for testing a string function return. Have the strings
'concat'ed and the total length at -4 from the ptr location. Still crashes on return...
'Something else may be getting passed or ??? Prob-ly sumpin simple....

'#ASM_START
'
'   LenS1 equ [ebp-4]
'   LenS2 equ [ebp-8]
'   TotalLen equ [ebp-12]
'   RetAddr equ [ebp-20]
'
'    push ebp             ;Save EBP
'    mov ebp, esp         ; Move ESP into EBP so we can refer

'    push ebx
'    push edx
'    push esi
'    push edi

'    sub esp, 76        ;   to arguments on the stack

';int 3
';    lea edi, [ebp-76]
'    mov esi, [ebp+8]     ; copy s1 pointer into destination index
';    mov esi, [ebp+12]       ; copy s2 pointer into source index
';mov eax, [ebp+8]
';mov eax, [eax]
';mov RetAddr,eax
';mov RetAddr,edx   ;dest addr has been preloaded by vb into edx, so save it in var

'    cld                 ; copy bytes forward
'    mov ecx, 2
'    mov eax,  [ebp+12]  ;load string 2
'StrLength:
'    mov ebx, [eax]
'    xor eax, eax
'    mov ax, [ebx-4]
';'    cdq
';'    sub eax, edx    ;next 2 ops- conv. from hex 2 dec.
';'    sar eax, 1
'    dec ecx
'    jecxz GetStringz   ;done getting both lengths
'    mov LenS2, eax   ;save len2
'    mov eax,  [ebp+8] ;load str 1
'    jmp StrLength    ;get len2
'GetStringz:
'    mov LenS1, eax   ;save len1
'   mov edx, LenS2      ;len2 in edx
'   add eax, edx        ;add len2 to len1
'   mov TotalLen,eax     ;put total in var
'   mov ecx, LenS1     ;put len2 in ecx
';'   shl ecx,1           ;basic strings are len*2
';'add ecx,4     ;+ 2 null terminators
';'mov eax, LenS1
';'shl eax, 1
';'add eax, ebx
';'mov [edi],eax

'mov edi,[RetAddr]
'mov [edi-4], eax
';int 3
'mov esi,[esi]
'mov edx,2
'MoveS2:
';int 3
'    cmp ecx, 4           ; if under 4 bytes long
'    jl TailEnd             ; jump to label tail
'    mov eax, ecx         ; copy length into eax
'    push eax            ; place a copy of eax on the stack
'    shr eax, 2          ; integer divide eax by 4
'    shl eax, 2          ; multiply eax by 4 to get dividend
'    mov ecx, eax       ; copy it into variable
'    pop eax             ; retrieve length in eax off the stack
'    sub eax, ecx        ; subtract dividend from length to get remainder
'    mov ebx, eax       ; copy remainder into variable
'    shr ecx, 2          ; divide by 4 for DWORD data size
'    rep movsd         ; repeat while not zero, move string DWORD
'    mov ecx, ebx       ; put remainder in ecx
' TailEnd:
'    rep movsb         ; copy remaining BYTES from source to dest
'dec edx
'test edx,1
'je DoneMov
'mov ecx, LenS2
'mov esi, [ebp+12]
'mov esi,[esi]
'jmp MoveS2
'DoneMov:
'xor eax,eax
'mov [edi],eax      ;null terminators
';int 3
'
'mov eax, ebp
'sub eax, 20   ;ptr to return string
';mov esi, eax


'    pop edi
'    pop esi

'    pop edx
'    pop ebx
'  mov esp, ebp         ;MOV/POP is much faster
'  pop ebp              ;on 486 and Pentium than Leave
'  ret  8               ; Return with sum in EAX
'                       ;If Arguments then-> eg. 2 longs use ret 8, 3 longs use ret 12 ...
'#ASM_END


End Function

