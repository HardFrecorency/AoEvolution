%define _WINALL_
%include "Include\Windows.inc"

[BITS 32]
[section .text]

procglobal DllMain, hinstDLL, fdwReason, lpvReserved
	LibMainPrologue
	mov	eax, 1
	LibMainEpilogue
endproc

procglobal GenCrc, _CrcKey, _CrcString
	push esi
	mov ecx, 0x8a9b4e75
	mov edx, ._CrcKey
	mov esi, ._CrcString
ProxByt:
	lodsb
	test al, al
	jz FinCad
	not al
	xor al, dl
	and cl, al
	ror ecx, cl
	not ecx
	rol edx, cl
	or al, cl
	xor dl, al
	jmp ProxByt
FinCad:
	pop esi
	xor edx, ecx
	mov eax, edx
	shr eax, 16
	xor ax, dx
endproc

[section .data]

