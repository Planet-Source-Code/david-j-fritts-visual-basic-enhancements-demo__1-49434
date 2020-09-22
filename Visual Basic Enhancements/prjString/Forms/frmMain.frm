VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "String Project"
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
Private Sub Form_Load()
    Dim strTesting As String
    
    Dim Timer As New clsTimer
    Dim StringClass As New clsString
    
        strTesting = "hello, how are you? my aunt's name is sally!"
        Timer.Reset
            For i = 1 To 1000000
                InStrRev strTesting, "'"
            Next i
        Debug.Print "(InStrRev) visual basic took " & Format(Timer.Elapsed, "0000") & " msec to execute"

        Timer.Reset
            For i = 1 To 1000000
                StringClass.StringRevScanC strTesting, 39   'Asc("'") = 39
            Next i
        Debug.Print "(StringRevScanC) machine code took " & Format(Timer.Elapsed, "0000") & " msec to execute"

        Timer.Reset
            For i = 1 To 1000000
                InStr strTesting, "'"
            Next i
        Debug.Print "(InStr) visual basic took " & Format(Timer.Elapsed, "0000") & " msec to execute"

        Timer.Reset
            For i = 1 To 1000000
                StringClass.StringScanC strTesting, 39   'Asc("'") = 39
            Next i
        Debug.Print "(StringScanC) machine code took " & Format(Timer.Elapsed, "0000") & " msec to execute"

        Debug.Print Mid$(strTesting, InStrRev(strTesting, "m"))
        Debug.Print Mid$(strTesting, StringClass.StringRevScanC(strTesting, Asc("m")))
        Debug.Print Mid$(strTesting, InStr(strTesting, "m"))
        Debug.Print Mid$(strTesting, StringClass.StringScanC(strTesting, 39))
        
    Set StringClass = Nothing
    Set Timer = Nothing
End Sub
