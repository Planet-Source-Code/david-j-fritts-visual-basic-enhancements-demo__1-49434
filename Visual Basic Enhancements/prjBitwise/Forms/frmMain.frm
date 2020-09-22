VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Bitwise Project"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim a As Long, b As Long, i As Long
    
    Dim Timer As New clsTimer
    Dim BitwiseClass As New clsBitwise
        
        b = 1
        Timer.Reset
            For i = 1 To 1000000
                a = b * 4 ^ 2
            Next i
        Debug.Print "(1 LEFTSHIFT 4) visual basic took " & Format(Timer.Elapsed, "0000") & " msec to execute"
        
        Timer.Reset
            For i = 1 To 1000000
                a = BitwiseClass.ShiftLeft(b, 4)
            Next i
        Debug.Print "(1 LEFTSHIFT 4) machine code took " & Format(Timer.Elapsed, "0000") & " msec to execute"
        
        Debug.Print "Bitwise ShiftLeft(2, 1) = " & BitwiseClass.ShiftLeft(2, 1)
        Debug.Print "Bitwise ShiftRight(4, 1) = " & BitwiseClass.ShiftRight(4, 1)
        Debug.Print "Bitwise RotateLeft(4, 1) = " & BitwiseClass.RotateLeft(4, 1)
        Debug.Print "Bitwise RotateRight(8, 1) = " & BitwiseClass.RotateRight(8, 1)
        
    Set BitwiseClass = Nothing
    Set Timer = Nothing
End Sub
