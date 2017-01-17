;AOKernel.dll
;Archivo de kernel Argentum 0.9.5
;Compilador: NASM
;Plataforma: Win32
;Función Inicial DLL: DllMain
;Programador: Otto Pérez








%define _WINALL_
%include "Include\Windows.inc"

extern _send@16

[BITS 32]
[section .text]

procglobal DllMain, hinstDLL, fdwReason, lpvReserved
	LibMainPrologue
	mov	eax, 1
	LibMainEpilogue
endproc

procglobal SetKernelMaxUsers, _MaxUsers
	mov eax, ._MaxUsers
	mov [MaxUsers], eax
endproc
;------------------------------------------------------
;Configuro en el kernel la cantidad máxima de usuarios.
;------------------------------------------------------



procglobal GetMapUsers, _InputMapArray, _OutputIndexArray, _MapLocked
	push ebx
	push ecx
	push edx
	push esi
	push edi

	mov eax, ._MapLocked
	mov ebx, [MaxUsers]
	inc ebx
	mov ecx, ebx
	mov esi, ._OutputIndexArray
	mov edi, ._InputMapArray
NextLock:
	repnz scasw
	jnz NoMasUsers
	mov edx, ebx
	sub edx, ecx
	mov word [esi], dx
	add esi, 2
	jmp NextLock
NoMasUsers:
	mov eax, esi
	sub eax, ._OutputIndexArray
	shr eax, 1

	pop edi
	pop esi
	pop edx
	pop ecx
	pop ebx	
endproc
;---------------------------------------------------------------------------------------------
;Toma del InputArray todos los valores iguales a MapLocked y pone el indice en el OutputArray.
;Retorno = Cantidad de items coincidentes.
;---------------------------------------------------------------------------------------------



procglobal GetMapUsersButIndex, _InputMapArray, _OutputIndexArray, _MapLocked, _UserIndex
	push ebx
	push ecx
	push edx
	push esi
	push edi

	mov ebx, [MaxUsers]
	inc ebx
	mov ecx, ebx
	mov esi, ._OutputIndexArray
	mov edi, ._InputMapArray

NextLock1:
	mov eax, ._MapLocked
	repnz scasw
	jnz NoMasUsers1
	mov edx, ebx
	sub edx, ecx
	mov eax, ._UserIndex
	cmp ax, dx
	jz NextLock1
	mov word [esi], dx
	add esi, 2
	jmp NextLock1
NoMasUsers1:
	mov eax, esi
	sub eax, ._OutputIndexArray
	shr eax, 1

	pop edi
	pop esi
	pop edx
	pop ecx
	pop ebx	
endproc
;---------------------------------------------------------------------------------------------
;Toma del InputArray todos los valores iguales a MapLocked y pone el indice en el OutputArray.
;Si algún valor es igual a UserIndex no lo cuento.
;Retorno = Cantidad de items coincidentes.
;---------------------------------------------------------------------------------------------



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
;--------------------------------------------------------------------------------
;Genera un número de control en base al parametro _CrcKey y el string _CrcString.
;--------------------------------------------------------------------------------



procglobal ReadFieldASM, _Pos, _Text, _SepASCII
	push ebx
	push ecx
	push edx
	push esi
	push edi

	mov ecx, ._Pos
	mov ebx, ._SepASCII
	mov esi, ._Text
	mov edi, ReadFieldOutPut
	dec ecx
	jz NoFieldError1
	jc FieldError
NoFieldError:
	lodsb
	test al, al
	je FieldError
	cmp al, bl
	jnz NoFieldError
	dec ecx
	jnz NoFieldError
NoFieldError1:
	lodsb
	test al, al
	je FieldError
	cmp al, bl
	je FieldError
	stosb
	jmp NoFieldError1
FieldError:
	sub edi, ReadFieldOutPut
	mov [ReadFieldOutPutLen], edi
	mov eax, ReadFieldOutPut

	pop edi
	pop esi
	pop edx
	pop ecx
	pop ebx
endproc
;--------------------------------------------------------------------------------------------------------------
;Esta funcion corta el campo especificado en _Pos con el separador _SepASCII sobre el string que esta en _Text.
;--------------------------------------------------------------------------------------------------------------



procglobal SendMapUsers, _UsersMapArray, _UsersSocketArray, _MapLocked, _TxtData
	push ebx
	push ecx
	push edx
	push esi
	push edi

	mov eax, ._MapLocked
	mov ecx, [MaxUsers]
	mov edi, ._UsersMapArray
NextUserMap:
	repnz scasw
	jnz NoMoreUsers
	push edi
	push ecx
	push eax
	mov ebx, edi
	mov esi, ._UsersMapArray
	sub ebx, esi
	sub ebx, 2
	shl ebx, 1
	mov esi, ._UsersSocketArray
	add esi, ebx
	mov eax, [esi]
	mov ebx, ._TxtData
	mov ecx, [ebx - 4]
	sc send, eax, ebx, ecx, NULL
	pop eax
	pop ecx
	pop edi
	jmp NextUserMap
NoMoreUsers:

	pop edi
	pop esi
	pop edx
	pop ecx
	pop ebx
endproc
;-------------------------------------------------------------------------------------
;Envío el string TxtData a los sockets de los usuarios que esten en el mapa MapLocked.
;-------------------------------------------------------------------------------------


procglobal SendMapUsersButIndex, _UsersMapArray, _UsersSocketArray, _MapLocked, _TxtData, _UserIndex
	push ebx
	push ecx
	push edx
	push esi
	push edi

	mov eax, ._MapLocked
	mov ecx, [MaxUsers]
	mov edi, ._UsersMapArray
NextUserMap1:
	repnz scasw
	jnz NoMoreUsers1
	push edi
	push ecx
	push eax
	mov ebx, edi
	mov esi, ._UsersMapArray
	sub ebx, esi
	mov eax, ebx
	shr eax, 1
	mov ecx, ._UserIndex
	cmp ax, cx
	pop eax
	pop ecx
	pop edi
	jz NextUserMap1
	push edi
	push ecx
	push eax
	sub ebx, 2
	shl ebx, 1
	mov esi, ._UsersSocketArray
	add esi, ebx
	mov eax, [esi]
	mov ebx, ._TxtData
	mov ecx, [ebx - 4]
	sc send, eax, ebx, ecx, NULL
	pop eax
	pop ecx
	pop edi
	jmp NextUserMap1
NoMoreUsers1:

	pop edi
	pop esi
	pop edx
	pop ecx
	pop ebx
endproc
;-------------------------------------------------------------------------------------
;Envío el string TxtData a los sockets de los usuarios que esten en el mapa MapLocked.
;-------------------------------------------------------------------------------------

[section .data]

MaxUsers		dd	0
ReadFieldOutPutLen	dd	0
ReadFieldOutPut		times 8192 db (0)
