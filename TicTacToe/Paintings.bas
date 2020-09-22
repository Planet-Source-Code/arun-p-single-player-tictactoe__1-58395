Attribute VB_Name = "Paintings"
'(c) Arun P
'Part Of Tic tac Toe

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function DrawCircleAt(intInx As Integer, Optional blnDel As Boolean = True, Optional bkcol As Boolean = True)
    If bkcol Then Form1.Picture1(intInx).BackColor = RGB(200, 200, 255)
    
    Dim iphi As Single
    iphi = 0
    
    If blnDel Then
        For iphi = 0 To 2 * 3.14152 Step 0.1 ' 0 to 2PI
            Form1.Picture1(intInx).Line (472 + Cos(iphi) * 300, 472 + Sin(iphi) * 300)-(472 + Cos(iphi) * 250, 472 + Sin(iphi) * 250)
            Sleep 1
        Next
    End If
    
    For iphi = 250 To 300 Step 1
        Form1.Picture1(intInx).Circle (472, 472), iphi
    Next
End Function

Public Function DrawXAt(intInx As Integer, Optional blnDel As Boolean = True, Optional bkcol As Boolean = True)
    If bkcol Then Form1.Picture1(intInx).BackColor = RGB(255, 200, 200)
    
    Dim buf1 As Integer, buf2 As Integer
    buf1 = 200
    While buf1 < 775
        Form1.Picture1(intInx).Line (buf1 + 30, buf1 - 30)-(buf1 - 30, buf1 + 30)
        If blnDel Then
            Sleep 1
        End If
        buf1 = buf1 + 25
    Wend
    
    For buf1 = 200 To 775 Step 25
        Form1.Picture1(intInx).Line (975 - (buf1 + 30), (buf1 - 30))-(975 - (buf1 - 30), (buf1 + 30))
        If blnDel Then
            Sleep 1
        End If
    Next
    
    For buf1 = 0 To 30
        Form1.Picture1(intInx).Line (230 - buf1, 170 + buf1)-(805 - buf1, 745 + buf1)
        Form1.Picture1(intInx).Line (975 - (230 - buf1), 170 + buf1)-(975 - (805 - buf1), 745 + buf1)
    Next
    
End Function


Public Function RefreshPaints(Optional blnDraw As Boolean = True)
    Dim buf1 As Integer
    For buf1 = 0 To 8
        If blnWho(buf1) = 2 Then
            DrawCircleAt buf1, False, blnDraw
        ElseIf blnWho(buf1) = 1 Then
            DrawXAt buf1, False, blnDraw
        End If
    Next
End Function

Public Function Highlight(int1 As Integer, int2 As Integer, int3 As Integer)
    Form1.Picture1(int1).BackColor = RGB(255, 255, 255)
    Form1.Picture1(int2).BackColor = RGB(255, 255, 255)
    Form1.Picture1(int3).BackColor = RGB(255, 255, 255)
    RefreshPaints False
End Function

'Private Function getmin(int1 As Integer, int2 As Integer, int3 As Integer) As Integer
'    getmin = int1
'    If int2 < getmin Then getmin = int2
'    If int3 < getmin Then getmin = int3
'End Function
'
'Private Function getmax(int1 As Integer, int2 As Integer, int3 As Integer) As Integer
'    getmax = int1
'    If int2 > getmax Then getmax = int2
'    If int3 > getmax Then getmax = int3
'End Function

