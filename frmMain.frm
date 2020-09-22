VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MineCat"
   ClientHeight    =   5610
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   1920
      Top             =   0
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   3840
      ScaleHeight     =   4215
      ScaleWidth      =   3975
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4215
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   4
      Left            =   1920
      Picture         =   "frmMain.frx":030A
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   3
      Left            =   1440
      Picture         =   "frmMain.frx":064C
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   960
      Picture         =   "frmMain.frx":098E
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   480
      Picture         =   "frmMain.frx":0CD0
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":1012
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuLevel 
         Caption         =   "Level"
         Begin VB.Menu mnuLvlBeg 
            Caption         =   "Beginner"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuLvlInt 
            Caption         =   "Intermediate"
         End
         Begin VB.Menu mnuLvlExp 
            Caption         =   "Expert"
         End
         Begin VB.Menu mnuLvlNuts 
            Caption         =   "Are you nuts???"
         End
      End
      Begin VB.Menu mnuBestScores 
         Caption         =   "Best scores"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private varSquareSide       ' ~~ Size (in pixels) of a Square
Private varSquaresHor, varSquaresVer
Private varLeft
Private varTop
Private varTimer
Private varLastWinnerName

Private varBeginnerName, varBeginnerScore
Private varIntermediateName, varIntermediateScore
Private varExpertName, varExpertScore
Private varNutsName, varNutsScore

' ~~ Create a Type that will hold Levels variables
' ~~ like the number of bombs and number of
' ~~ horizontal and vertical Squares
Private Type Difficulty
    NumberOfBombs As Integer
    SquaresVertical As Integer
    SquaresHorizontal As Integer
End Type

Private LevelBeginner As Difficulty
Private LevelMedium As Difficulty
Private LevelExpert As Difficulty
Private LevelNuts As Difficulty
Private myLevel As Difficulty

Private arrayDiscover()
Private arrayHiddenField()
Private arrayPlayerField()
Private arrayFieldStatus()

Private Const BASICSQUARE = 0
Private Const EMPTYSQUARE = 1
Private Const BOMBSQUARE = 2
Private Const BOUMSQUARE = 3
Private Const FLAGSQUARE = 4


Private Sub Command1_Click()

' ~~ Depending on what menu is selected, the Difficulty
' ~~ level changes
If mnuLvlBeg.Checked = True Then
    varDifLev = 0
ElseIf mnuLvlInt.Checked = True Then
    varDifLev = 1
ElseIf mnuLvlExp.Checked = True Then
    varDifLev = 2
ElseIf mnuLvlNuts.Checked = True Then
    varDifLev = 3
End If

' ~~ Let's start a new game
Call subNewGame(varDifLev)

End Sub

Private Sub subNewGame(fDifLev)
Picture1.Cls

Call subSetGameDifficulty(fDifLev)
Call subResizeWindow
Call subDrawSquares(0, 0)
Call subInitializeGame

End Sub

Private Sub subDrawSquares(fX, fY)
' ~~ Draw a new Playground with the number of Squares
' ~~ defined by the Difficulty level
If fX = 0 Then defX = myLevel.SquaresHorizontal Else defX = fX
If fY = 0 Then defY = myLevel.SquaresVertical Else defY = fY

For a = 1 To defX
    For b = 1 To defY
        Picture1.PaintPicture Image1(BASICSQUARE).Picture, varLeft + ((a - 1) * varSquareSide), varTop + ((b - 1) * varSquareSide)
    Next b
Next a

End Sub

Private Sub subDrawPictureOnSquare(fX, fY, fPicture, fControl)
' ~~ Draws a particular picture in the specified Square
fControl.PaintPicture Image1(fPicture).Picture, varLeft + ((fX - 1) * varSquareSide), varTop + ((fY - 1) * varSquareSide)

End Sub

Private Sub Form_Load()

Timer1.Interval = 1000
Timer1.Enabled = False

' ~~ We'll firs tinitialize all variables and arrays
Call subInitialize
' ~~ then give the form the right size
Call subResizeWindow
' ~~ and draw the initial Playground
Call subDrawSquares(LevelBeginner.SquaresHorizontal, LevelBeginner.SquaresVertical)

Call subPrintMessage("Press GO to start")

End Sub

Private Sub mnuAbout_Click()
MsgBox "A MineSweeper clone, just to make time pass faster..." & vbCrLf & vbCrLf & "  version " & App.Major & "." & App.Minor & "." & App.Revision, vbInformation, "MineCat"
End Sub

Private Sub mnuBestScores_Click()
' ~~ calls the second form, which displays highest scores
Form2.Show 1
End Sub

Private Sub mnuExit_Click()
' ~~ basta!
Unload Me
End Sub

Private Sub mnuLvlBeg_Click()
' ~~ Beginner level
mnuLvlBeg.Checked = True
mnuLvlInt.Checked = False
mnuLvlExp.Checked = False
mnuLvlNuts.Checked = False
Command1_Click
End Sub

Private Sub mnuLvlExp_Click()
' ~~ Expert level
mnuLvlBeg.Checked = False
mnuLvlInt.Checked = False
mnuLvlExp.Checked = True
mnuLvlNuts.Checked = False
Command1_Click
End Sub

Private Sub mnuLvlInt_Click()
' ~~ Intermediate level
mnuLvlBeg.Checked = False
mnuLvlInt.Checked = True
mnuLvlExp.Checked = False
mnuLvlNuts.Checked = False
Command1_Click
End Sub

Private Sub mnuLvlNuts_Click()
' ~~ Crazy level
mnuLvlBeg.Checked = False
mnuLvlInt.Checked = False
mnuLvlExp.Checked = False
mnuLvlNuts.Checked = True
Command1_Click
End Sub

Private Sub mnuNewGame_Click()
' ~~ Selecting 'New Game' in the menu is like clicking
' ~~ the button
Command1_Click
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' ~~ Has the player clicked inside the Playground???
If X > varLeft And X < varLeft + (myLevel.SquaresHorizontal * varSquareSide) Then
    If Y > varTop And Y < varTop + (myLevel.SquaresVertical * varSquareSide) Then
        ' ~~ If Yes, then let's find what Square was
        ' ~~ clicked
        varTmp = X Mod varSquareSide
        varSquareHor = ((X - varTmp) / varSquareSide) + 1
        varTmp = Y Mod varSquareSide
        varSquareVer = ((Y - varTmp) / varSquareSide) + 1
        
        ' ~~ Left or Right mouse button ?
        If Button = 1 Then
            ' ~~ Left: player takes a chance
            Call subPlayGame(varSquareHor, varSquareVer)
        ElseIf Button = 2 Then
            ' ~~ Right: player places a Flag, a Question
            ' ~~ mark or resets one of these signs
            Select Case arrayFieldStatus(varSquareHor, varSquareVer)
                Case -1   'discovered by App
                
                Case 0    'Nothing
                ' ~~ if there's nothing, draw a Flag
                Call subDrawPictureOnSquare(varSquareHor, varSquareVer, FLAGSQUARE, Picture1)
                arrayFieldStatus(varSquareHor, varSquareVer) = 1
                
                Case 1    'Flag
                ' ~~ if there's a Flag, draw a Question mark
                Call subDrawPictureOnSquare(varSquareHor, varSquareVer, BASICSQUARE, Picture1)
                Call subDrawQuestionMark(varSquareHor, varSquareVer)
                arrayFieldStatus(varSquareHor, varSquareVer) = 2
                
                Case 2    'Question mark
                ' ~~ if there's a Question mark, reset
                Call subDrawPictureOnSquare(varSquareHor, varSquareVer, BASICSQUARE, Picture1)
                arrayFieldStatus(varSquareHor, varSquareVer) = 0
            End Select
            
            ' ~~ if all Bombs have been flagged
            ' ~~ correctly, the Game is over and
            ' ~~ the player has won
            If fnCheckIfFinished = True Then
                Timer1.Enabled = False
                Picture1.Enabled = False
                ' ~~ check if player has beaten the
                ' ~~ highest score for this
                ' ~~ difficulty level
                Call subCheckBestScore
            End If

        End If
    End If
End If

End Sub

Private Sub subPlayGame(X, Y)
' ~~ If there's a bomb......
If arrayHiddenField(X, Y) = 1 Then
    ' ~~ BOUM!
    Call subExplode(X, Y)
Else
    ' ~~ otherwise, see what's around
    Call subMineSweep(X, Y, True)
    ' ~~ if the previous action has made more Squares
    ' ~~ ready to explore, let's do it
    If fnStackCount(True) > 0 Then
        Do
            varSt = fnFirstFromStack
            varY = varSt Mod 1000
            varX = (varSt - varY) / 1000
            Call subMineSweep(varX, varY, True)
            ' ~~ Get the square out of the Stack list
            Call subStackOffList(varX, varY)
        ' ~~ Continue exploring until there are no
        ' ~~ squares left in the Stack
        Loop Until fnStackCount(True) = 0
    End If
    ' ~~ if all Bombs have been flagged
    ' ~~ correctly, the Game is over and
    ' ~~ the player has won
    If fnCheckIfFinished = True Then
        Timer1.Enabled = False
        Picture1.Enabled = False
        ' ~~ check if player has beaten the
        ' ~~ highest score for this
        ' ~~ difficulty level
        Call subCheckBestScore
    End If
End If
End Sub

Private Sub subInitializeGame()

Picture1.Enabled = True
varTimer = 0
Timer1.Enabled = True
Label1 = ""

' ~~ dimension the arrays following the number of
' ~~ vertical and horizontal squares defined by the
' ~~ difficulty level
ReDim arrayHiddenField(myLevel.SquaresHorizontal, myLevel.SquaresVertical)
ReDim arrayPlayerField(myLevel.SquaresHorizontal, myLevel.SquaresVertical)
ReDim arrayFieldStatus(myLevel.SquaresHorizontal, myLevel.SquaresVertical)
' ~~ and initialize them
For a = 1 To myLevel.SquaresHorizontal
    For b = 1 To myLevel.SquaresVertical
        arrayHiddenField(a, b) = 0
        arrayPlayerField(a, b) = 0
        arrayFieldStatus(a, b) = 0
    Next b
Next a

' ~~ let's now place the bombs on the minefield
Randomize Timer
PlacedBombs = 0
Do
    tmpCol = Int(Rnd * myLevel.SquaresHorizontal) + 1
    tmpRow = Int(Rnd * myLevel.SquaresVertical) + 1
    If arrayHiddenField(tmpCol, tmpRow) = 0 Then
        arrayHiddenField(tmpCol, tmpRow) = 1
        PlacedBombs = PlacedBombs + 1
        ' ~~ Debug
        ' ~~ uncomment next line to see bombs appear
        ' ~~ on the playground
        'Call subDrawPictureOnSquare(tmpCol, tmpRow, BOMBSQUARE, Picture1)
    End If
' ~~ continue until the number of bombs defined by the
' ~~ difficulty level is reached
Loop Until PlacedBombs = myLevel.NumberOfBombs

End Sub

Private Sub subInitialize()
varSquareSide = 230
varLeft = 0
varTop = 0
' ~~ let's define each difficulty level and attribute
' ~~ to each of them a number of bombs and a number
' ~~ of vertical and horizontal Squares
With LevelBeginner
    .NumberOfBombs = 16
    .SquaresVertical = 12
    .SquaresHorizontal = 10
End With
With LevelMedium
    .NumberOfBombs = 35
    .SquaresVertical = 16
    .SquaresHorizontal = 14
End With
With LevelExpert
    .NumberOfBombs = 75
    .SquaresVertical = 25
    .SquaresHorizontal = 25
End With
With LevelNuts
    .NumberOfBombs = 350
    .SquaresVertical = 33
    .SquaresHorizontal = 50
End With
End Sub

Private Sub subSetGameDifficulty(fDifficultyLevel)
Select Case fDifficultyLevel
    Case 0
    myLevel.NumberOfBombs = LevelBeginner.NumberOfBombs
    myLevel.SquaresHorizontal = LevelBeginner.SquaresHorizontal
    myLevel.SquaresVertical = LevelBeginner.SquaresVertical
    ReDim arrayDiscover(myLevel.SquaresHorizontal * myLevel.SquaresVertical, 2)
    
    Case 1
    myLevel.NumberOfBombs = LevelMedium.NumberOfBombs
    myLevel.SquaresHorizontal = LevelMedium.SquaresHorizontal
    myLevel.SquaresVertical = LevelMedium.SquaresVertical
    ReDim arrayDiscover(myLevel.SquaresHorizontal * myLevel.SquaresVertical, 2)
    
    Case 2
    myLevel.NumberOfBombs = LevelExpert.NumberOfBombs
    myLevel.SquaresHorizontal = LevelExpert.SquaresHorizontal
    myLevel.SquaresVertical = LevelExpert.SquaresVertical
    ReDim arrayDiscover(myLevel.SquaresHorizontal * myLevel.SquaresVertical, 2)
    
    Case 3
    myLevel.NumberOfBombs = LevelNuts.NumberOfBombs
    myLevel.SquaresHorizontal = LevelNuts.SquaresHorizontal
    myLevel.SquaresVertical = LevelNuts.SquaresVertical
    ReDim arrayDiscover(myLevel.SquaresHorizontal * myLevel.SquaresVertical, 2)

End Select
End Sub

Private Sub subExplode(fX, fY)
' ~~ BOUM!
Timer1.Enabled = False

Call subDrawPictureOnSquare(fX, fY, BOUMSQUARE, Picture1)
' ~~ Wherever there's a Bomb, draw it!
For a = 1 To myLevel.SquaresHorizontal
    For b = 1 To myLevel.SquaresVertical
        If arrayHiddenField(a, b) = 1 And arrayPlayerField(a, b) = 0 And a <> fX And b <> fY Then
            Call subDrawPictureOnSquare(a, b, BOMBSQUARE, Picture1)
        End If
    Next b
Next a

End Sub

Private Sub subMineSweep(fX, fY, fMode)
' ~~ See what's under the Square the player just clicked

' ~~ first count the number of bombs in the (maximum)
' ~~ 8 squares around the clicked one
varAround = fnCountBombsAround(fX, fY)
' ~~ Square has been clicked
arrayPlayerField(fX, fY) = 1
arrayFieldStatus(fX, fY) = -1
' ~~ if at least one bomb around, we write the number
' ~~ of bombs in the Square and stop exploring
If varAround > 0 Then
    Call subDrawPictureOnSquare(fX, fY, EMPTYSQUARE, Picture1)
    Picture1.FontBold = True
    Picture1.ForeColor = vbBlue
    varX = varLeft + ((fX - 1) * varSquareSide) + ((varSquareSide / 2) - (Picture1.TextWidth(varAround) / 2))
    varY = varTop + ((fY - 1) * varSquareSide) + ((varSquareSide / 2) - (Picture1.TextHeight(varAround) / 2))
    Picture1.CurrentX = varX
    Picture1.CurrentY = varY
    If fMode = True Then Picture1.Print varAround
Else
    ' ~~ if there are no bombs around, we'll explore
    ' ~~ each square contiguous to this one as well
    Call subDrawPictureOnSquare(fX, fY, EMPTYSQUARE, Picture1)
    ' ~~ so we'll add all neighbour squares to the Stack
    Call subAddNeighboursToStack(fX, fY)
End If

End Sub

Private Function fnCountBombsAround(fX, fY)
' ~~ count the number of bombs around a specific Square
If fX > 0 Then LeftLimit = fX - 1 Else LeftLimit = fX
If fX < myLevel.SquaresHorizontal Then RightLimit = fX + 1 Else RightLimit = fX
If fY > 0 Then TopLimit = fY - 1 Else TopLimit = fY
If fY < myLevel.SquaresVertical Then BottomLimit = fY + 1 Else BottomLimit = fY

BombsCount = 0
For a = LeftLimit To RightLimit
    For b = TopLimit To BottomLimit
        If Not (a = fX And b = fY) Then
            If arrayHiddenField(a, b) = 1 Then
                BombsCount = BombsCount + 1
            End If
        End If
    Next b
Next a
fnCountBombsAround = BombsCount
End Function

Private Sub subStackInList(fSquareX, fSquareY)
' ~~ put square in the first free Stack dimension
For c = 1 To UBound(arrayDiscover)
    If arrayDiscover(c, 1) = 0 Then
        arrayDiscover(c, 1) = fSquareX
        arrayDiscover(c, 2) = fSquareY
        Exit For
    End If
Next c
End Sub

Private Sub subStackOffList(fSquareX, fSquareY)
' ~~ remove the Square from the Stack
For c = 1 To UBound(arrayDiscover)
    If arrayDiscover(c, 1) = fSquareX And arrayDiscover(c, 2) = fSquareY Then
        arrayDiscover(c, 1) = 0
        arrayDiscover(c, 2) = 0
        Exit For
    End If
Next c
End Sub

Private Function fnStackCount(fMode)
' ~~ count the Squares in the Stack.
' ~~  if fMode is false, we really count them.
' ~~  if fMode is true, we just see if there's at least one
varCount = 0
For c = 1 To UBound(arrayDiscover)
    If arrayDiscover(c, 1) <> 0 Then
        varCount = varCount + 1
        If fMode = True Then
            fnStackCount = 1
            Exit Function
        End If
    End If
Next c
fnStackCount = varCount
End Function

Private Function fnFirstFromStack()
' ~~ retrieve the first Square in the Stack
varFoundX = 0: varFoundY = 0
For c = 1 To UBound(arrayDiscover)
    If arrayDiscover(c, 1) <> 0 Then
        varFoundX = arrayDiscover(c, 1)
        varFoundY = arrayDiscover(c, 2)
        Exit For
    End If
Next c
If varFoundX <> 0 Then
    fnFirstFromStack = (1000 * varFoundX) + varFoundY
Else
    fnFirstFromStack = 0
End If
End Function

Public Sub subAddNeighboursToStack(fX, fY)
' ~~ this sub adds all neighbours of a specific Square
' ~~ to the Stack
If fX > 0 Then LeftLimit = fX - 1 Else LeftLimit = fX
If fX < myLevel.SquaresHorizontal Then RightLimit = fX + 1 Else RightLimit = fX
If fY > 0 Then TopLimit = fY - 1 Else TopLimit = fY
If fY < myLevel.SquaresVertical Then BottomLimit = fY + 1 Else BottomLimit = fY

For a = LeftLimit To RightLimit
    For b = TopLimit To BottomLimit
        If Not (a = fX And b = fY) Then
            If arrayPlayerField(a, b) = 0 And arrayFieldStatus(a, b) < 1 Then
                Call subStackInList(a, b)
            End If
        End If
    Next b
Next a
End Sub

Private Sub subResizeWindow()
' ~~ The number of Squares defines the size of the
' ~~ playground ...
If mnuLvlBeg.Checked = True Then
    Picture1.Width = varLeft + (LevelBeginner.SquaresHorizontal * varSquareSide)
    Picture1.Height = varTop + (LevelBeginner.SquaresVertical * varSquareSide)
ElseIf mnuLvlInt.Checked = True Then
    Picture1.Width = varLeft + (LevelMedium.SquaresHorizontal * varSquareSide)
    Picture1.Height = varTop + (LevelMedium.SquaresVertical * varSquareSide)
ElseIf mnuLvlExp.Checked = True Then
    Picture1.Width = varLeft + (LevelExpert.SquaresHorizontal * varSquareSide)
    Picture1.Height = varTop + (LevelExpert.SquaresVertical * varSquareSide)
ElseIf mnuLvlNuts.Checked = True Then
    Picture1.Width = varLeft + (LevelNuts.SquaresHorizontal * varSquareSide)
    Picture1.Height = varTop + (LevelNuts.SquaresVertical * varSquareSide)
End If
' ~~ ... and of the window
Me.Width = Picture1.Width + 120
Me.Height = Picture1.Height + Command1.Height + 680
Command1.Left = (Me.Width / 2) - (Command1.Width / 2)
End Sub

Private Sub Timer1_Timer()
' ~~ displays the score
varTimer = varTimer + 1
Label1 = varTimer
End Sub

Public Sub subDrawQuestionMark(fX, fY)
    Picture1.FontBold = True
    Picture1.ForeColor = vbBlack
    varX = varLeft + ((fX - 1) * varSquareSide) + ((varSquareSide / 2) - (Picture1.TextWidth("?") / 2))
    varY = varTop + ((fY - 1) * varSquareSide) + ((varSquareSide / 2) - (Picture1.TextHeight("?") / 2))
    Picture1.CurrentX = varX
    Picture1.CurrentY = varY
    Picture1.Print "?"
End Sub

Public Function fnCheckIfFinished()
' ~~ let's check if the game is finished, that is
' ~~ if the player has correctly flagged all bombs
varFlags = 0:  varFlagsPlacedOK = 0: varPlayed = 0
For a = 1 To myLevel.SquaresHorizontal
    For b = 1 To myLevel.SquaresVertical
        If arrayFieldStatus(a, b) <> 0 Then
            varPlayed = varPlayed + 1
        End If
        If arrayFieldStatus(a, b) = 1 Then
            varFlags = varFlags + 1
            If arrayHiddenField(a, b) = 1 Then
                varFlagsPlacedOK = varFlagsPlacedOK + 1
            End If
        End If
    Next b
Next a
If varPlayed = (myLevel.SquaresHorizontal * myLevel.SquaresVertical) And varFlags = myLevel.NumberOfBombs And varFlags = varFlagsPlacedOK Then
    fnCheckIfFinished = True
Else
    fnCheckIfFinished = False
End If
End Function

Public Sub subLoadBestScores()
' ~~ if the file Scores.dat exist in the application
' ~~ directory, let's load its data (highest scores)
varPath = App.Path & "\"
ReDim arrayScores(100)
If Dir(varPath & "scores.dat", vbNormal) <> "" Then
    f = FreeFile
    Open varPath & "scores.dat" For Input As #f
        varCount = 0
        Do Until EOF(f)
            varCount = varCount + 1
            If varCount > 100 Then Exit Do
            Line Input #f, arrayScores(varCount)
        Loop
    Close #f
    For a = 1 To 100
        If InStr(arrayScores(a), "BEGINNERNAME=") <> 0 Then
            varEq = InStr(arrayScores(a), "=")
            varBeginnerName = Right(arrayScores(a), Len(arrayScores(a)) - varEq)
        End If
        If InStr(arrayScores(a), "BEGINNERSCORE=") <> 0 Then
            varEq = InStr(arrayScores(a), "=")
            varBeginnerScore = Right(arrayScores(a), Len(arrayScores(a)) - varEq)
        End If
        
        If InStr(arrayScores(a), "INTERMEDIATENAME=") <> 0 Then
            varEq = InStr(arrayScores(a), "=")
            varIntermediateName = Right(arrayScores(a), Len(arrayScores(a)) - varEq)
        End If
        If InStr(arrayScores(a), "INTERMEDIATESCORE=") <> 0 Then
            varEq = InStr(arrayScores(a), "=")
            varIntermediateScore = Right(arrayScores(a), Len(arrayScores(a)) - varEq)
        End If
                
        If InStr(arrayScores(a), "EXPERTNAME=") <> 0 Then
            varEq = InStr(arrayScores(a), "=")
            varExpertName = Right(arrayScores(a), Len(arrayScores(a)) - varEq)
        End If
        If InStr(arrayScores(a), "EXPERTSCORE=") <> 0 Then
            varEq = InStr(arrayScores(a), "=")
            varExpertScore = Right(arrayScores(a), Len(arrayScores(a)) - varEq)
        End If
                
        If InStr(arrayScores(a), "NUTSNAME=") <> 0 Then
            varEq = InStr(arrayScores(a), "=")
            varNutsName = Right(arrayScores(a), Len(arrayScores(a)) - varEq)
        End If
        If InStr(arrayScores(a), "NUTSSCORE=") <> 0 Then
            varEq = InStr(arrayScores(a), "=")
            varNutsScore = Right(arrayScores(a), Len(arrayScores(a)) - varEq)
        End If
    Next a
Else
    ' ~~ otherwise, initialize the highest scores
    varBeginnerName = "": varBeginnerScore = 0
    varIntermediateName = "": varIntermediateScore = 0
    varExpertName = "": varExpertScore = 0
    varNutsName = "": varNutsScore = 0
End If
End Sub

Public Sub subCheckBestScore()
' ~~ verifies if the player has beaten a highest score

' ~~ first load the Highest scores, if any
Call subLoadBestScores

ReDim arrayHighest(4, 3)
arrayHighest(0, 1) = varBeginnerName
arrayHighest(0, 2) = CLng(varBeginnerScore)
arrayHighest(1, 1) = varIntermediateName
arrayHighest(1, 2) = CLng(varIntermediateScore)
arrayHighest(2, 1) = varExpertName
arrayHighest(2, 2) = CLng(varExpertScore)
arrayHighest(3, 1) = varNutsName
arrayHighest(3, 2) = CLng(varNutsScore)

If mnuLvlBeg.Checked = True Then
    varLevelName = "Beginner"
    varHighScore = CLng(varBeginnerScore)
    varL = 0
ElseIf mnuLvlInt.Checked = True Then
    varLevelName = "Intermediate"
    varHighScore = CLng(varIntermediateScore)
    varL = 1
ElseIf mnuLvlExp.Checked = True Then
    varLevelName = "Expert"
    varHighScore = CLng(varExpertScore)
    varL = 2
ElseIf mnuLvlNuts.Checked = True Then
    varLevelName = "Are you Nuts???"
    varHighScore = CLng(varNutsScore)
    varL = 3
End If
If varTimer < varHighScore Or varHighScore = 0 Then
    ' ~~ if there are no high scores or if the player
    ' ~~ has beaten it
    varWinnerName = ""
    Do
        varWinnerName = InputBox("You have made a Best Score!!! Please enter your name.", "Best Score", varLastWinnerName)
    Loop While varWinnerName = ""
    varLastWinnerName = varWinnerName
    If Len(varWinnerName) > 75 Then varWinnerName = Left(varWinnerName, 75)
    arrayHighest(varL, 1) = varWinnerName
    arrayHighest(varL, 2) = varTimer
    varPath = App.Path & "\"
    f = FreeFile
    ' ~~ Write the new results
    Open varPath & "scores.dat" For Output As #f
        Print #f, "BEGINNERNAME=" & arrayHighest(0, 1)
        Print #f, "BEGINNERSCORE=" & arrayHighest(0, 2)
        Print #f, "INTERMEDIATENAME=" & arrayHighest(1, 1)
        Print #f, "INTERMEDIATESCORE=" & arrayHighest(1, 2)
        Print #f, "EXPERTNAME=" & arrayHighest(2, 1)
        Print #f, "EXPERTSCORE=" & arrayHighest(2, 2)
        Print #f, "NUTSNAME=" & arrayHighest(3, 1)
        Print #f, "NUTSSCORE=" & arrayHighest(3, 2)
    Close #f
    Form2.Show 1
End If

End Sub

Private Sub subPrintMessage(fMsg)
Picture1.ForeColor = vbBlack
Picture1.FontBold = True
Picture1.CurrentX = (Picture1.Width / 2) - (Picture1.TextWidth(fMsg) / 2)
Picture1.CurrentY = (Picture1.Height / 2) - (Picture1.TextHeight(fMsg) / 2)
Picture1.Print fMsg
Picture1.FontBold = False
End Sub
