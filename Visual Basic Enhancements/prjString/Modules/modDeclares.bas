Attribute VB_Name = "modDeclares"
Public Declare Function GlobalAlloc Lib "KERNEL32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "KERNEL32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "KERNEL32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "KERNEL32" (ByVal hMem As Long) As Long

Public Declare Sub RtlMoveMemory Lib "KERNEL32" (Destination As Any, Source As Any, ByVal Length As Long)
