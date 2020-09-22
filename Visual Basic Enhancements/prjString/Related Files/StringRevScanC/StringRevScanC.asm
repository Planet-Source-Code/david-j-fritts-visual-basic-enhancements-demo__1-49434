%define string     [esp+0ch] 
%define searchchar [esp+10h]
%define return 	   [esp+14h]

[BITS 32]
push esi
mov esi, dword string
mov edx, dword return

mov ecx, [esi-4]
mov ax, word searchchar

Scan_Loop:
cmp ax, word [esi+ecx]
je Found
sub ecx, 2
jnz Scan_Loop

Found:
shr ecx, 1
inc ecx

mov dword [edx], ecx
ExitMethod:
xor eax, eax		;hresult success = 0
pop esi
ret 10h