VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "!__Single Player TicTacToe__!"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      Height          =   375
      Left            =   3960
      MaskColor       =   &H00FFFFC0&
      TabIndex        =   11
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "About"
      Height          =   375
      Left            =   3960
      MaskColor       =   &H00FFFFC0&
      TabIndex        =   10
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   975
      Index           =   8
      Left            =   2520
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   975
      Index           =   7
      Left            =   1440
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   975
      Index           =   6
      Left            =   360
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   975
      Index           =   5
      Left            =   2520
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   975
      Index           =   4
      Left            =   1440
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   975
      Index           =   3
      Left            =   360
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   975
      Index           =   2
      Left            =   2520
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   975
      Index           =   1
      Left            =   1440
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   975
      Index           =   0
      Left            =   360
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debug Info"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim blnOcc(0 To 8) As Boolean

Dim msn As Long
Dim buff As Integer

Dim aival As String



Private Sub Command1_Click()
    MsgBox "By Arun P, arun_pbk@hotmail.com" & vbCrLf & "Entirely New. I created this game over a year before " & vbCrLf & "in TurboC++. Now i made it to VB. I'm only a beginner. If u like it please vote for me (just to know my position, not for contest). " & vbCrLf & "Since i dont have time now to explain the program, please mail me for explanation (IF ANY!!)", vbOKOnly, "TIC TAC TOE"
    RefreshPaints
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
RefreshPaints
End Sub

Private Sub Form_Load()
    msn = 0
    playid = ""
    Me.BackColor = RGB(180, 200, 255)
    
    If Dir(App.Path + "\ai.ai") = "" Then
        MsgBox "AI File not found. Please contact arun_pbk@hotmail.com"
        Unload Me
    End If
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    For buff = 0 To 8
        If blnOcc(buff) = False Then
            Picture1(buff).BackColor = &HE0E0E0
        End If
    Next
End Sub

Private Sub Form_Paint()
    RefreshPaints
End Sub

Private Sub Picture1_Click(Index As Integer)
    If blnOcc(Index) = False Then
        blnOcc(Index) = True
        'Picture1(Index).BackColor = RGB(255, 200, 200)
        DrawXAt (Index)
        blnWho(Index) = 1 'player
        
        won
        
        playid = playid + Chr(Index + 65)
        
        aival = GetAIMove()
        'MsgBox aival
        DrawCircleAt Asc(aival) - 65
        blnOcc(Asc(aival) - 65) = True
        blnWho(Asc(aival) - 65) = 2 'cpu
        
        won
    Else
        Beep
    End If
End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If blnOcc(Index) = False Then
        Picture1(Index).BackColor = RGB(200, 255, 200)
    End If
    'removes green from other column
    For buff = 0 To 8
        If blnOcc(buff) = False And buff <> Index Then
            Picture1(buff).BackColor = &HE0E0E0
        End If
    Next
End Sub

