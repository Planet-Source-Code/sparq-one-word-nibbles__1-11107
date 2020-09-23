VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6360
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   14
      Top             =   3300
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   5220
      TabIndex        =   11
      Top             =   4560
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "(un) &Pause"
      Height          =   495
      Left            =   5220
      TabIndex        =   10
      Top             =   4020
      Width           =   3255
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   180
      ScaleHeight     =   465
      ScaleWidth      =   4965
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   4995
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5940
      Top             =   5040
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5220
      ScaleHeight     =   165
      ScaleWidth      =   705
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5220
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   3
      Top             =   1860
      Width           =   180
   End
   Begin VB.PictureBox picBoard 
      BackColor       =   &H00000000&
      Height          =   4730
      Left            =   180
      ScaleHeight     =   4665
      ScaleWidth      =   4965
      TabIndex        =   0
      Top             =   180
      Width           =   5030
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click Right-Arrow to Start."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1560
         TabIndex        =   13
         Top             =   2580
         Width           =   1815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click for new game."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1560
         TabIndex        =   7
         Top             =   2580
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Label lblHS 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   180
      Width           =   2475
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This is you, Green the snake. "
      Height          =   195
      Left            =   6060
      TabIndex        =   9
      Top             =   1560
      Width           =   2115
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Gather the Berries - as many as you can. In order to do this, you must avoid the walls and yourself!"
      Height          =   1215
      Left            =   5520
      TabIndex        =   8
      Top             =   1860
      Width           =   2970
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6240
      TabIndex        =   5
      Top             =   420
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5220
      TabIndex        =   4
      Top             =   420
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3120
      TabIndex        =   1
      Top             =   6480
      Width           =   45
   End
   Begin VB.Shape Border 
      Height          =   495
      Left            =   3720
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SquareSize = 10
Dim Start As Boolean
Dim Xs(0 To 900)
Dim Ys(0 To 900)
Dim NextX As Integer
Dim NextY As Integer
Dim NextSpot As Integer
Dim Direction As jDirection
Dim LastDir As jDirection
Dim Bool As Boolean
Dim XX As Integer
Dim YY As Integer
Dim PauseCount As Integer
Dim Dead As Boolean
Dim Hs As Integer

Public Enum jDirection
    Left
    Right
    up
    DOWN
End Enum

Private Sub Command1_Click()
    'Stop / Start the timer
    Timer1.Enabled = Not (Timer1.Enabled)
    'Set focus back to the game
    picBoard.SetFocus
End Sub

Private Sub Command2_Click()
    End 'ByeBye
End Sub

Private Sub Form_Load()
    'Load Graphics
    Start = True
    Picture1.Picture = LoadPicture(App.Path & "\green.gif")
    Picture2.Picture = LoadPicture(App.Path & "\fruit.gif")
    Picture3.Picture = LoadPicture(App.Path & "\gameover.gif")
    
    picBoard.Cls
    
    'initialize controls
    Picture3.Visible = False
    Timer1.Enabled = False
    Label3.Visible = False
    Visible = True
    Command1.Enabled = False
    lblStart.Visible = True
    
    
    Do Until Visible
        DoEvents
    Loop
    Direction = DOWN
    Bool = False
    LoadHS     'Load high Score
    InitSnake  'Draw first Screen
    PopUpFruit 'Draw 1st Fruit
End Sub

Sub LoadHS()
  Dim Readln
    'if there is not high score file, create one with the high score of Zero
    If Dir(App.Path & "\hs.ini") = "" Then SaveHS (0)
    
    'Load High Score
    Open App.Path & "\hs.ini" For Input As #1
        Line Input #1, Readln
    Close #1
    Hs = Readln
    
    'display Score
    lblHS = "High Score: " & Hs
End Sub

Sub SaveHS(Score As String)
    'open HS.ini and write the High Score
    Open App.Path & "\hs.ini" For Output As #1
        Print #1, Score
    Close #1
End Sub
Sub PopUpFruit()
10
    'XX: X Coordinate for fruit
    'YY: Y Coordinate for fruit
    'Both must start at a number below 0 that is not divisible by 10 evenly
    XX = -6
    YY = -6

    'Create X/Y-Coords until we find one that is divisible by 10 (so its directly
    'in the path of the snake. they must be greater than zero but less than the Width/Height
    Do Until (XX Mod 10 = 0) And (XX > 0) And (XX < 320)
        Randomize Timer
        XX = Int((330 * Rnd) + 1)
        DoEvents
    Loop
    Do Until (YY Mod 10 = 0) And (YY > 0) And (YY < 300)
        Randomize Timer
        YY = Int((310 * Rnd) + 1)
        DoEvents
    Loop
    
    'Make sure we're not sprouting a fruit on top of the snake!!
    X = 0
    Do Until Xs(X) < 0
        If Xs(X) = XX And Ys(X) = YY Then Randomize Timer: GoTo 10
        X = X + 1
        DoEvents
    Loop
    
    
End Sub

Private Sub Form_Resize()
    'resize border
    With Border
        .Top = 0
        .Left = 0
        .Width = Width
        .Height = Height
    End With
End Sub


Sub InitSnake()
    'Set X/Y - Coords to -1  (-1 indicates an empty slot)
    For X = 0 To 900
        Xs(X) = -1
        Ys(X) = -1
    Next X
    
    'Setup the snake to be 3 Units long (0, 1, and 2)
    For X = 0 To 2
        Xs(X) = 50 + (X * 10)
        Ys(X) = 160
    Next X
    'Start with a direction other than right (since right is the direction we will start in,
    'If the LastDirection is right, when the user hits right to start the game, it will be a
    'disregarded command.
    LastDir = DOWN
    
    'I should hope not
    Dead = False
    
    'Clear Current Score
    lblScore = "00000"
    
    'Actually paint the snake
    DrawSnake False
End Sub

Sub DrawSnake(Optional Pause As Boolean)

  'if pause = true then the snake will not stay the smae length, but will constantly grow
  
  'Clear Board
  'picBoard.Cls
  If Not Start Then BitBlt picBoard.hDC, XX, YY, 10, 10, Picture2.hDC, 0, 0, SRCCOPY
  
  Dim X As Integer
    X = 0
    Do Until Xs(X) < 0
        BitBlt picBoard.hDC, Xs(X), Ys(X), SquareSize, SquareSize, Picture1.hDC, 0, 0, SRCCOPY
        X = X + 1
        DoEvents
    Loop
    
    If Not Pause Then BitBlt picBoard.hDC, Xs(0), Ys(0), SquareSize, SquareSize, Picture4.hDC, 0, 0, SRCCOPY
    
    If Pause Then GoTo Ending
    X = 0
    Do Until Xs(X + 1) < 0
        Xs(X) = Xs(X + 1)
        Ys(X) = Ys(X + 1)
        X = X + 1
        DoEvents
    Loop
    NextSpot = X
    Exit Sub
Ending:
    NextSpot = X
End Sub

Private Sub picBoard_Click()
    If Not Timer1.Enabled Then Form_Load
End Sub

Private Sub picBoard_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyRight Then
        If Direction = Left Or LastDir = Left Then Exit Sub
        Timer1.Enabled = True
        Direction = Right
    ElseIf KeyCode = vbKeyLeft Then
        If Direction = Right Or LastDir = Right Then Exit Sub
        Direction = Left
    ElseIf KeyCode = vbKeyUp Then
        If Direction = DOWN Or LastDir = DOWN Then Exit Sub
        Direction = up
    ElseIf KeyCode = vbKeyDown Then
        If Direction = up Or LastDir = up Then Exit Sub
        Direction = DOWN
    End If
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Command1.Enabled = True
    lblStart.Visible = False
    Start = False
    Select Case Direction
        Case up
            Ys(NextSpot) = Ys(NextSpot - 1) - 10
            Xs(NextSpot) = Xs(NextSpot - 1)
            LastDir = up
        Case DOWN
            Ys(NextSpot) = Ys(NextSpot - 1) + 10
            Xs(NextSpot) = Xs(NextSpot - 1)
            LastDir = DOWN
        Case Left
            Xs(NextSpot) = Xs(NextSpot - 1) - 10
            Ys(NextSpot) = Ys(NextSpot - 1)
            LastDir = Left
        Case Right
            Xs(NextSpot) = Xs(NextSpot - 1) + 10
            Ys(NextSpot) = Ys(NextSpot - 1)
            LastDir = Right
    End Select
    CheckforDead
    If Dead = False Then CheckforBerry
End Sub

Sub CheckforBerry()
    
    If Bool Then
        If PauseCount = 2 Then
            Bool = False
            PauseCount = 1
        Else
            PauseCount = PauseCount + 1
        End If
    End If
    If Xs(NextSpot) = XX And Ys(NextSpot) = YY Then
        Bool = True
        PauseCount = 0
        PopUpFruit
        lblScore = Format(Val(lblScore) + 9, "00000")
    End If
    DrawSnake Bool
End Sub

Sub CheckforDead()
    If Xs(NextSpot) < 0 Then GoTo Endd
    If Xs(NextSpot) > 330 Then GoTo Endd
    If Ys(NextSpot) < 0 Then GoTo Endd
    If Ys(NextSpot) > 300 Then GoTo Endd
    
    X = 0
    Do Until Xs(X + 2) < 0
        If Xs(NextSpot) = Xs(X) And Ys(NextSpot) = Ys(X) Then
            GoTo Endd
        End If
        X = X + 1
    Loop
    Exit Sub

Endd:
    Timer1.Enabled = False
    Dead = True
    Label3.Visible = True
    If Val(lblScore) > Hs Then SaveHS lblScore
    picBoard.Cls
    Picture3.Visible = True
End Sub
