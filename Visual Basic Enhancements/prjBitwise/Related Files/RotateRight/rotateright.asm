%define value  [esp+08h] 
%define shift  [esp+0Ch]
%define return [esp+10h]

[BITS 32]
mov eax, dword value
mov ecx, dword shift
mov edx, dword return

ror eax, cl

mov dword [edx], eax
xor eax, eax		;hresult is returned in eax
ret 10h			;if (eax == 0) then success