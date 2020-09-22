Attribute VB_Name = "AI1"
'Tic tac toe
'By arun p

Option Explicit

Public playid As String
Public blnWho(0 To 8) As Integer

Private lngPos As Long
Private incid As String

'simple isn't it!!!!!!!!!!!!
'u can win i have made some holes
Public Function GetAIMove() As String
    Dim incretAI As String
    
    incid = ""
    lngPos = 1

    Open "ai.ai" For Binary Access Read As #1
    Do
        Line Input #1, incid
    Loop While playid <> (Mid(incid, 1, InStr(1, incid, ">") - 1)) And Not EOF(1)
    
    Close #1
    Form1.Label2.Caption = playid
    GetAIMove = Right(incid, 1)
End Function
''''

Public Function won()
    won = 0
    If blnWho(0) = blnWho(1) And blnWho(0) = blnWho(2) Then
        If blnWho(0) = 2 Then
            Highlight 0, 1, 2
            MsgBox "Sorry Computer Won!!!"
            End
        ElseIf blnWho(0) = 1 Then
            Highlight 0, 1, 2
            MsgBox "Wow u won!!!"
            End
        End If
    End If
    
    If blnWho(3) = blnWho(4) And blnWho(3) = blnWho(5) Then
        If blnWho(3) = 2 Then
            Highlight 3, 4, 5
            MsgBox "Sorry Computer Won!!!"
            End
        ElseIf blnWho(3) = 1 Then
            Highlight 3, 4, 5
            MsgBox "Wow u won!!!"
            End
        End If
    End If
    
    If blnWho(6) = blnWho(7) And blnWho(6) = blnWho(8) Then
        If blnWho(6) = 2 Then
            Highlight 6, 7, 8
            MsgBox "Sorry Computer Won!!!"
            End
        ElseIf blnWho(6) = 1 Then
            Highlight 6, 7, 8
            MsgBox "Wow u won!!!"
            End
        End If
    End If
    
    If blnWho(0) = blnWho(3) And blnWho(0) = blnWho(6) Then
        If blnWho(0) = 2 Then
            Highlight 0, 3, 6
            MsgBox "Sorry Computer Won!!!"
            End
        ElseIf blnWho(0) = 1 Then
            Highlight 0, 3, 6
            MsgBox "Wow u won!!!"
            End
        End If
    End If
    If blnWho(1) = blnWho(4) And blnWho(1) = blnWho(7) Then
        If blnWho(1) = 2 Then
            Highlight 1, 4, 7
            MsgBox "Sorry Computer Won!!!"
            End
        ElseIf blnWho(1) = 1 Then
            Highlight 1, 4, 7
            MsgBox "Wow u won!!!"
            End
        End If
    End If
    If blnWho(2) = blnWho(5) And blnWho(2) = blnWho(8) Then
        If blnWho(2) = 2 Then
            Highlight 2, 5, 8
            MsgBox "Sorry Computer Won!!!"
            End
        ElseIf blnWho(2) = 1 Then
            Highlight 2, 5, 8
            MsgBox "Wow u won!!!"
            End
        End If
    End If
    If blnWho(0) = blnWho(4) And blnWho(0) = blnWho(8) Then
        If blnWho(0) = 2 Then
            Highlight 0, 4, 8
            MsgBox "Sorry Computer Won!!!"
            End
        ElseIf blnWho(0) = 1 Then
            Highlight 0, 4, 8
            MsgBox "Wow u won!!!"
            End
        End If
    End If
    If blnWho(2) = blnWho(4) And blnWho(2) = blnWho(6) Then
        If blnWho(2) = 2 Then
            Highlight 2, 4, 6
            MsgBox "Sorry Computer Won!!!"
            End
        ElseIf blnWho(2) = 1 Then
            Highlight 2, 4, 6
            MsgBox "Wow u won!!!"
            End
        End If
    End If
    
    For lngPos = 0 To 8
        If blnWho(lngPos) = 0 Then GoTo dd:
    Next
    
    MsgBox "Draw Game!!!"
    End
dd:
End Function
