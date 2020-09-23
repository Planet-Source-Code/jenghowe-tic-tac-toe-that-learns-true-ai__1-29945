VERSION 5.00
Begin VB.Form TicTacToe 
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   2895
   ClientLeft      =   4020
   ClientTop       =   4050
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   193
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   392
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Clear History"
      Height          =   375
      Left            =   4440
      TabIndex        =   16
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ListBox ScoreBoard 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      ForeColor       =   &H80000007&
      Height          =   2565
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":000A
      TabIndex        =   15
      Top             =   210
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opponent"
      Height          =   1215
      Left            =   4440
      TabIndex        =   10
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton Gtype 
         Caption         =   "CPU First"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Gtype 
         Caption         =   "CPU"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   8
         Left            =   1800
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   14
         Top             =   1920
         Width           =   615
      End
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   1
         Left            =   960
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   7
         Left            =   960
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   7
         Top             =   1920
         Width           =   615
      End
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   6
         Left            =   120
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   6
         Top             =   1920
         Width           =   615
      End
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   5
         Left            =   1800
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   5
         Top             =   1080
         Width           =   615
      End
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   4
         Left            =   960
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   4
         Top             =   1080
         Width           =   615
      End
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   3
         Left            =   120
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   3
         Top             =   1080
         Width           =   615
      End
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   2
         Left            =   1800
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox Box 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   0
         Left            =   120
         ScaleHeight     =   41
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.Line Line4 
         X1              =   1680
         X2              =   1680
         Y1              =   240
         Y2              =   2520
      End
      Begin VB.Line Line3 
         X1              =   840
         X2              =   840
         Y1              =   240
         Y2              =   2520
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   2400
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2400
         Y1              =   960
         Y2              =   960
      End
   End
End
Attribute VB_Name = "TicTacToe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Tic Tac Toe by JengHowe, with the ability to learn (or true AI)
'Keeps a record of all the games that it has lost and calculates the best possible
'move. If the CPU is allowed to play first, it would use the record to find the best
'move to win. The record file that came with this code 'MovHis.txt' is with all possible
'combinations the CPU would lose, so basically the CPU would not lose, and if allowed to
'play first, would win you if you are not careful. To make the CPU dumb again, just delete
'the record file, "movhis.txt".
'
'This is a revolutionary way of coding tic tac toe AI. the AI is is extremely robust
'and is capable of self improvement, and all this is done in a short time and without
'any really complex algorithms and long code.
'
'Note, the CPU only learns when you are the one who start, and it only gets as smart
'as you do.

Dim CPUplay As Boolean
Dim whowin, playstr As String
Dim NumBoxFilled, NGP, games_played As Integer
Dim MovHis() As String
Dim loseweight(8) As Integer

'Numbering of the Picture Box
' 0 1 2
' 3 4 5
' 6 7 8
'
'Draws X on the desinated Picture Box, Also Sets the tag property to 'X'
Private Sub drawX(Index As Integer)
Box(Index).Line (0, 1)-(38, 39)
Box(Index).Line (0, 0)-(39, 39)
Box(Index).Line (1, 0)-(39, 38)
Box(Index).Line (38, 1)-(1, 38)
Box(Index).Line (38, 0)-(0, 38)
Box(Index).Line (37, 0)-(0, 37)
Box(Index).Tag = "X"
End Sub

'Draws O on the desinated Picture Box, Also sets the tag property to 'O'
Private Sub drawO(Index As Integer)
Box(Index).FillStyle = 0
Box(Index).FillColor = &H0&
Box(Index).Circle (20, 20), 19
Box(Index).FillColor = &H8000000F
Box(Index).Circle (20, 20), 17
Box(Index).Tag = "O"
End Sub

'New Button
'Clears all the Picture Box, and reinitialises the variables
Private Sub Command1_Click()
Dim mv As Integer
For i = 0 To 8
    Box(i).Cls
    Box(i).Tag = ""
Next
CPUplay = False
whowin = ""
playstr = ""
NumBoxFilled = 0
Gtype(1).Enabled = True
Gtype(2).Enabled = True
Gtype(1).Value = True
End Sub

'Function that returns true if a winning combination is found
Private Function win() As Boolean
If rowsame(0, 1, 2) Or rowsame(3, 4, 5) Or rowsame(6, 7, 8) Or _
    rowsame(0, 3, 6) Or rowsame(1, 4, 7) Or rowsame(2, 5, 8) Or _
    rowsame(0, 4, 8) Or rowsame(2, 4, 6) Then win = True
End Function

'Sub Function used by Win() to check wheter the 3 PictureBox has the same Symbol
Private Function rowsame(i1, i2, i3 As Integer) As Boolean
If (Box(i1).Tag = Box(i2).Tag) And (Box(i2).Tag = Box(i3).Tag) And (Box(i1).Tag = "X" Or Box(i1).Tag = "O") Then
    rowsame = True
    whowin = Box(i1).Tag
Else
    rowsame = False
End If
End Function

'Exit Button
'Compacts the move history by only saving losing moves, then saves it into a file
'named 'movhis.txt'
Private Sub Command2_Click()
Dim CompactHis() As String
Dim CNum, i, j, k As Integer
For i = 1 To NGP
    If Right(MovHis(i), 1) = "X" Then   'Saves only losing moves
        CNum = CNum + 1
        ReDim Preserve CompactHis(CNum)
        CompactHis(CNum) = MovHis(i)
    End If
Next
If NGP = 0 Then End
'Open File for output
Open App.Path + "\MoveHis.txt" For Output As #1
Print #1, CNum
For i = 1 To CNum
    Print #1, CompactHis(i)
Next
Close #1
End
End Sub

'Executes when the Picture Boxes are clicked
Private Sub Box_Click(Index As Integer)
Dim i, mvpos, rp As Integer
Dim mvstr As String
Gtype(1).Enabled = False
Gtype(2).Enabled = False

playstr = playstr + CStr(Index)
If Box(Index).Tag = "" Then
    drawX (Index)
    CPUplay = True
    NumBoxFilled = NumBoxFilled + 1
End If

'If either side wins or draw, exit subroutine
If checkforwin Then Exit Sub

'CPU Play
If CPUplay Then
    mvstr = FindMoves
    rp = Int(Rnd * Len(mvstr)) + 1
    mvpos = CDec(Mid(mvstr, rp, 1))
    Debug.Print "Available Moves:"; mvstr; " Choose:"; mvpos
    drawO (mvpos)
    playstr = playstr + CStr(mvpos)
    CPUplay = False
    NumBoxFilled = NumBoxFilled + 1
End If

If checkforwin Then Exit Sub
End Sub

'Function that checks for winning combination
'when CPU loses, it will then record the current game moves into the move history
'Since the game board is 3x3 matrix, it is symmetrical in 8 ways, so 1 game played
'in a particular direction can be interpreted into 8 identical games. Matrix
'manipulation "Transpose" and "Rotate" is used to obtain the extra 7 equivalent games.
'
'Basically, what this does is it speeds up the CPU learning process by 8 fold.
Private Function checkforwin() As Boolean
checkforwin = False
If win Then
    MsgBox whowin + " Won", , "Game Over!"
    games_played = games_played + 1
    ScoreBoard.AddItem ("Game " + CStr(games_played) + " : " + whowin + " Win")
    playstr = playstr + whowin
    Debug.Print playstr
    If Gtype(1).Value Then
        NGP = NGP + 1: ReDim Preserve MovHis(NGP): MovHis(NGP) = playstr
        NGP = NGP + 1: ReDim Preserve MovHis(NGP): MovHis(NGP) = transpose(playstr)
        NGP = NGP + 1: ReDim Preserve MovHis(NGP): MovHis(NGP) = rotate(playstr)
        NGP = NGP + 1: ReDim Preserve MovHis(NGP): MovHis(NGP) = transpose(rotate(playstr))
        NGP = NGP + 1: ReDim Preserve MovHis(NGP): MovHis(NGP) = rotate(rotate(playstr))
        NGP = NGP + 1: ReDim Preserve MovHis(NGP): MovHis(NGP) = transpose(rotate(rotate(playstr)))
        NGP = NGP + 1: ReDim Preserve MovHis(NGP): MovHis(NGP) = rotate(rotate(rotate(playstr)))
        NGP = NGP + 1: ReDim Preserve MovHis(NGP): MovHis(NGP) = transpose(rotate(rotate(rotate(playstr))))
    End If
    Gtype(1).Enabled = True
    Gtype(2).Enabled = True
    Gtype(1).Value = True
    Call Command1_Click
    checkforwin = True
ElseIf NumBoxFilled = 9 Then
    MsgBox "Draw", , "Game Over!"
    games_played = games_played + 1
    ScoreBoard.AddItem ("Game " + CStr(games_played) + " : " + "Draw")
    playstr = playstr + "D"
    Debug.Print playstr
    Call Command1_Click
    checkforwin = True
    Gtype(1).Enabled = True
    Gtype(2).Enabled = True
    Gtype(1).Value = True
End If
End Function

'Matrix Manipulation Function
'
' 012    036
' 345 => 147
' 678    258
Private Function transpose(tstr As String) As String
Dim temp, c, c1 As String
Dim i, rp As Integer
For i = 1 To Len(tstr)
    c = Mid(tstr, i, 1)
    Select Case c
        Case Is = "1": c1 = "3"
        Case Is = "3": c1 = "1"
        Case Is = "2": c1 = "6"
        Case Is = "6": c1 = "2"
        Case Is = "5": c1 = "7"
        Case Is = "7": c1 = "5"
        Case Else: c1 = c
    End Select
    temp = temp + c1
Next
transpose = temp
End Function

'Matrix Manipulation Function
'Rotates the Matrix by 90 degrees, to rotate 180, just use the function twice
'for 270, three times
'
' 012    258
' 345 => 147
' 678    036

Private Function rotate(tstr As String) As String
Dim temp, c, c1 As String
Dim i, rp As Integer
For i = 1 To Len(tstr)
    c = Mid(tstr, i, 1)
    Select Case c
        Case Is = "0": c1 = "2"
        Case Is = "1": c1 = "5"
        Case Is = "2": c1 = "8"
        Case Is = "3": c1 = "1"
        Case Is = "4": c1 = "4"
        Case Is = "5": c1 = "7"
        Case Is = "6": c1 = "0"
        Case Is = "7": c1 = "3"
        Case Is = "8": c1 = "6"
        Case Else: c1 = c
    End Select
    temp = temp + c1
Next
rotate = temp
End Function

'Finds the Appropriate move for the CPU, the function returns a string containing all
'possible moves
Private Function FindMoves() As String
Dim i, j, l, mv, sml As Integer
Dim pst, mvc As String
For i = 0 To 8
    loseweight(i) = 0
Next
For j = 0 To 8
  pst = playstr + CStr(j)
  l = Len(pst)
  For i = 1 To NGP
    If l < Len(MovHis(i)) Then
        If pst = Left(MovHis(i), l) Then
            If Right(MovHis(i), 1) = "X" Then loseweight(j) = loseweight(j) + 1
        End If
    End If
  Next i
Next j

'Different weight calculation when the CPU moves first
If Gtype(1).Value Then  'When CPU moves Second
    sml = 32767
    For i = 0 To 8
        If sml > loseweight(i) And Box(i).Tag = "" Then sml = loseweight(i)
    Next
Else    'When CPU moves First
    sml = 0
    For i = 0 To 8
        If sml < loseweight(i) And Box(i).Tag = "" Then sml = loseweight(i)
    Next
End If

For i = 0 To 8
    If loseweight(i) = sml And Box(i).Tag = "" Then mvc = mvc + CStr(i)
Next

'Checks wheter if the Human Player has a winning move, if there is, then block it
For i = 0 To 8
    If Box(i).Tag = "" Then
        Box(i).Tag = "X"
        If win Then mvc = CStr(i)
        Box(i).Tag = ""
        
    End If
Next

'Checks wheter the CPU has a winning move, if it does, then choose that move.
For i = 0 To 8
    If Box(i).Tag = "" Then
        Box(i).Tag = "O"
        If win Then mvc = CStr(i)
        Box(i).Tag = ""
        
    End If
Next
FindMoves = mvc
End Function

'Clears Score Board
Private Sub Command3_Click()
games_played = 0
ScoreBoard.Clear
End Sub

'Executed when the Program starts, loads the move history from file if available
Private Sub Form_Load()
Dim i As Integer
Randomize Timer
On Error GoTo errHand
Open App.Path + "\MoveHis.txt" For Input As #1
Input #1, NGP
ReDim MovHis(NGP)
For i = 1 To NGP
    Input #1, MovHis(i)
Next
errHand:
Close #1
End Sub

'Executes when the Options menu are Clicked
'Chooses between CPU play first or Human First
Private Sub Gtype_Click(Index As Integer)
If Gtype(2).Value Then
    mv = Int(Rnd * 9)
    drawO (mv)
    NumBoxFilled = NumBoxFilled + 1
    playstr = CStr(mv)
Gtype(1).Enabled = False
Gtype(2).Enabled = False
End If
End Sub
