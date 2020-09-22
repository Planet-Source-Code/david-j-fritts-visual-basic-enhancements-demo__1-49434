VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private seqFreq As Currency
Private seqStart As Currency
Private fAvailable As Boolean

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Sub Class_Initialize()
  If QueryPerformanceFrequency(seqFreq) = 0& Then
    MsgBox "PerfCounter not available", vbExclamation
  Else
    fAvailable = True
  End If
End Sub

Friend Sub Reset()
  QueryPerformanceCounter seqStart
End Sub

Friend Function Elapsed() As Double
  Dim seqNow As Currency
  QueryPerformanceCounter seqNow
  If fAvailable Then
    Elapsed = 1000 * (seqNow - seqStart) / seqFreq
  End If
End Function

