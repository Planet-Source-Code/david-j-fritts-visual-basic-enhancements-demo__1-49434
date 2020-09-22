%define value  [esp+08h] 
%define shift  [esp+0Ch]
%define return [esp+10h]

[BITS 32]
mov eax, dword value
mov ecx, dword shift
mov edx, dword return

shl eax, cl

mov dword [edx], eax
xor eax, eax		;hresult success = 0
ret 10h