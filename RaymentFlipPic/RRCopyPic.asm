;CopyPic.asm   For A386
;Copies picture in pic1 to pic2
;Out: res&=eax

;VB
;ptbmp1& = bmp1.bmBits
;ptbmp2& = bmp2.bmBits
;ReDim inparams&(6)
;inparams&(1) = bmp1.bmWidth
;inparams&(2) = bmp1.bmHeight
;inparams&(3) = bmp1.bmWidthBytes
;ptparams = VarPtr(inprams&(1))
;res& = CallWindowProc(ptmc&, ptbmp1&, ptbmp2&, ptparams, 4&)
;stack positions              8        12       16        20

USE32
long1 equ[ebp+8]	;->pic1
long2 equ[ebp+12]	;->pic2
long3 equ[ebp+16]	;-> params
long4 equ[ebp+20]	;4&  Spare
	
bmwidth       equ[ebp-4]
bmheight      equ[ebp-8]
bmwidthbytes  equ[ebp-12]
totbytes      equ[ebp-16]

	push ebp		;Arrange for ebp to point to stack params
	mov ebp,esp
	sub esp,16	
	push edi,esi,ebx

;Get and store bm params
	mov ebx,long3
	mov eax,[ebx]
	mov bmwidth,eax
	inc ebx,4
	mov eax,[ebx]
	mov bmheight,eax
	mov edx,eax
	inc ebx,4
	mov eax,[ebx]
	mov bmwidthbytes,eax

	mul edx
	inc ebx,4
	mov totbytes,eax
	mov ecx,bmwidthbytes
	shr ecx,2			;/4 for 4 byte moves

	mov edx,ecx
	mov esi,long1		;ptr to pic1
	mov edi,long2		;ptr to pic2
	;add edi,totbytes
	;sub edi,bmwidthbytes	;get to start of scan at top pic2
	;mov ebx,edi		;save
;Set num Y lines
	mov ecx,bmheight
start:
	push ecx
	;push ebx
	
	mov ecx,edx
	cld		;ensure incr
	rep movsd	;[esi]->[edi] esi+4 edi+4 ecx-4

	;decrement di
	;pop ebx
	;sub ebx,bmwidthbytes
	;mov edi,ebx


	pop ecx
	dec ecx
	jnz start
GETOUT:
	pop ebx,esi,edi
	mov esp,ebp
	pop ebp
	ret 16
