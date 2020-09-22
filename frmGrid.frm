VERSION 5.00
Begin VB.Form frmGrid 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Same"
   ClientHeight    =   3480
   ClientLeft      =   165
   ClientTop       =   690
   ClientWidth     =   4725
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGrid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGrid 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3165
      Left            =   0
      ScaleHeight     =   3165
      ScaleWidth      =   4725
      TabIndex        =   0
      Top             =   0
      Width           =   4725
      Begin VB.Shape shpHiLite 
         BorderColor     =   &H00FFFFC0&
         Height          =   330
         Left            =   2415
         Top             =   1575
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Config"
      Begin VB.Menu mnuRows 
         Caption         =   "Rows: 10"
      End
      Begin VB.Menu mnuCols 
         Caption         =   "Cols: 40"
      End
      Begin VB.Menu mnuColors 
         Caption         =   "Colors: 2"
      End
      Begin VB.Menu mnuSp 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuBWidth 
         Caption         =   "Block Width: 200"
      End
      Begin VB.Menu mnuBHeight 
         Caption         =   "Block Height: 200"
      End
   End
   Begin VB.Menu mnuGen 
      Caption         =   "Generate"
   End
   Begin VB.Menu mnuScores 
      Caption         =   "Scores"
      Begin VB.Menu mnuScore 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu mnuScore 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu mnuScore 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu mnuScore 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu mnuScore 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu mnuScore 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu mnuScore 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu mnuScore 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu mnuScore 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu mnuScore 
         Caption         =   ""
         Index           =   9
      End
   End
End
Attribute VB_Name = "frmGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BWidth As Integer 'Block Width
Dim BHeight As Integer 'Block Height
Dim aryB() As clsBlock 'Array of game 'pieces'
Dim MaxRows As Integer 'Rows on the board
Dim MaxCols As Integer 'columns on the board
Dim MaxColors As Integer 'Colors in the game
Dim SelColor As Integer 'Currently Selected Color
Dim CurPoints As Long 'Current Total Score
Dim SelPoints As Long 'Number of points attained in a play
Dim SelCount As Integer 'Number of blocks triggered i a play
Dim LastCol As Integer 'Last column affected in a play
Dim KeyRow As Integer 'Row of key-board input
Dim KeyCol As Integer 'Column of keyboard input
Dim CurBlock As Integer 'Current Block selected
Dim HiRow As Integer 'Top most row affected in a play
Dim LowCol As Integer 'Lowermost row affected in a play
Sub DoSelect(BIndex As Integer)
If BIndex < 0 Then Exit Sub 'Just in case a valid block isn't selected
SelColor = aryB(BIndex).Color 'The Selected Blocks color
If SelColor = 0 Then Exit Sub 'If the block is black, then it's disabled
If Not HasMate(BIndex) Then Exit Sub 'Make sure that the selected block can be triggered
SelCount = 0
LowCol = MaxCols
HiRow = 0
SelBlock BIndex 'Select the block
SelPoints = ((SelCount * MaxColors * MaxColors / (MaxRows * MaxCols)) * 100) * SelCount * MaxColors * MaxColors 'This is my scoring algorthym
'----------Re-Arrange the array-----------
SlideDown
SlideLeft
DrawScreen 'Redraw our blocks
End Sub

Sub DrawScreen()
'Itterate through the array of blocks and paint them on the screen
Dim sX As Integer
Dim sY As Integer
Dim i As Integer
picGrid.Cls
For i = 0 To (MaxCols * MaxRows) - 1
    Do While aryB(i).Color = 0 And i < (MaxCols * MaxRows) - 1
        i = i + 1
    Loop
    sX = (aryB(i).Col * BWidth)
    sY = (aryB(i).Row * BHeight)
    picGrid.Line (sX, sY)-(sX + BWidth, sY + BHeight), QBColor(aryB(i).Color), BF
    picGrid.Line (sX, sY)-(sX + BWidth, sY + BHeight), QBColor(0), B
Next i
CurPoints = CurPoints + SelPoints
Me.Caption = "Same - " & CurPoints
Select Case EndSet() 'Check to see if the game is over
    Case 0
        MsgBox "You Did It"
        SavehiScore CurPoints
    Case -1
        MsgBox "No Moves Left"
End Select
End Sub
Function EndSet() As Integer
'this looks for a valid block  .  . . .
Dim temp As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
temp = -1
If LastCol = 0 Then
    temp = 0
    GoTo TheEnd
End If
For i = MaxRows - 1 To 0 Step -1
    For j = 0 To LastCol - 1
        k = RCtoIndex(i, j)
        'If aryB(k).Color <> 0 Then
            If HasMate(k) Then
                temp = 1
                GoTo TheEnd
            End If
        'End If
    Next j
Next i
TheEnd:
EndSet = temp
End Function

Sub Generate()
'This is where we generate an array of playing pieces.
Dim i As Integer
Dim j As Integer
Dim BCount As Integer
'-----------Reset our var's from the last game played--------
SelPoints = 0
CurPoints = 0
Me.Width = (MaxCols * BWidth) + 135
Me.Height = (MaxRows * BHeight) + 700
picGrid.Height = Me.Height - 660
shpHiLite.Width = BWidth
shpHiLite.Height = BHeight
BCount = 0
LastCol = MaxCols
KeyRow = 0
KeyCol = 0
CurBlock = 0
'---------------Initiate a 'TRUE' random number
Randomize
'------------------------------------------
ReDim aryB(MaxRows * MaxCols) As clsBlock 'this is also a reset of the array
'-----Iterate through rows and columns to create the array--------
For i = 0 To MaxRows - 1
    For j = 0 To MaxCols - 1
        Set aryB(BCount) = New clsBlock
        aryB(BCount).Row = i
        aryB(BCount).Col = j
        aryB(BCount).Color = CInt(((MaxColors - 1) * Rnd) + 1)
        BCount = BCount + 1
    Next j
Next i

DrawScreen 'Draw the blocks on the screen
shpHiLite.Visible = True
HiliteBlock (0) 'Hilite the first block
End Sub

Function GetAryIndex(X As Single, Y As Single) As Integer
'This converts X and Y coords into a related index within the array
'by retreiving the related row and column
'Notice it makes use of the RCtoIndex function
Dim Row As Integer
Dim Col As Integer
Dim i As Integer
Row = CInt((Y + (BHeight / 2)) / BHeight) - 1
Col = CInt((X + (BWidth / 2)) / BWidth) - 1
GetAryIndex = RCtoIndex(Row, Col)
End Function

Function HasMate(BIndex As Integer) As Boolean
'This checks to see if there are any valid plays for the selected block
'On Error Resume Next
Dim temp As Boolean
Dim BColor As Integer
temp = False
BColor = aryB(BIndex).Color
If BColor = 0 Then GoTo TheEnd
If UpIndex(BIndex) <> -1 Then If aryB(UpIndex(BIndex)).Color = BColor Then temp = True
If DownIndex(BIndex) <> -1 Then If aryB(DownIndex(BIndex)).Color = BColor Then temp = True
If LeftIndex(BIndex) <> -1 Then If aryB(LeftIndex(BIndex)).Color = BColor Then temp = True
If RightIndex(BIndex) <> -1 Then If aryB(RightIndex(BIndex)).Color = BColor Then temp = True
TheEnd:
HasMate = temp
End Function

Sub HiliteBlock(BIndex As Integer)
'hilite a block
shpHiLite.Left = aryB(BIndex).Col * BWidth
shpHiLite.Top = aryB(BIndex).Row * BHeight
End Sub

Function NZ(MyInput) As Variant
'This is a handy little function to make sure you don't return a null
'I stole the idea from MSAccess and use it everywhere for validation
If MyInput = "" Then
    NZ = 0
Else
    NZ = MyInput
End If
End Function

Function RCtoIndex(Row, Col) As Integer
'This converts rows and columns into a related index within the array
On Error GoTo BadBlock
Dim i As Integer
Dim temp As Integer
temp = -1
If Row >= MaxRows Or Row < 0 Then GoTo BadBlock
If Col >= MaxCols Or Col < 0 Then GoTo BadBlock
temp = (Row * MaxCols) + Col
BadBlock:
RCtoIndex = temp
End Function


Sub SavehiScore(MyScore As Long)
'Again, in order to avoid creating other forms and stuff
'I take the cheap-method and save the scores to a menu
Dim i As Integer
Dim j  As Integer
Dim response As String
For i = 0 To 9
    If NZ(GetSetting("SameGame", "Scores", "Points" & i, 0)) < MyScore Then
        response = InputBox("Please Enter Your Name", "You Have A High Score")
        If response = "" Then Exit Sub
            For j = 9 To i + 1 Step -1
                SaveSetting "SameGame", "Scores", "Points" & j, GetSetting("SameGame", "Scores", "Points" & j - 1)
                SaveSetting "SameGame", "Scores", "Name" & j, GetSetting("SameGame", "Scores", "Name" & j - 1)
                SaveSetting "SameGame", "Scores", "Level" & j, GetSetting("SameGame", "Scores", "Level" & j - 1)
            Next j
                SaveSetting "SameGame", "Scores", "Points" & i, MyScore
                SaveSetting "SameGame", "Scores", "Name" & i, Left(response, 20)
                SaveSetting "SameGame", "Scores", "Level" & i, MaxRows & "X" & MaxCols & "X" & MaxColors
            Exit Sub
        End If
Next i
End Sub


Sub SelBlock(BIndex As Integer)
'This is the HEART of the program.
'Simply put, it recursivly searches through all of the blocks
'surrounding the selected block, and then
'selects those blocks . . .and so on and so on . .
'BTW a selected block ends up with a color = 0 (black)
If BIndex < 0 Then Exit Sub
If aryB(BIndex).Color <> SelColor Then Exit Sub
aryB(BIndex).Color = 0
SelBlock UpIndex(BIndex)
SelBlock DownIndex(BIndex)
SelBlock LeftIndex(BIndex)
SelBlock RightIndex(BIndex)
SelCount = SelCount + 1
HiliteBlock BIndex
'These values are stored for optimization for sliding and drawing.
'There is no need to evaluate the entire array or redraw the entire screen
If aryB(BIndex).Row > HiRow Then HiRow = aryB(BIndex).Row
If aryB(BIndex).Col < LowCol Then LowCol = aryB(BIndex).Col
End Sub
Function ShiftDown(Row As Integer, Col As Integer) As Boolean
'This is the most confusing part of the code
'It looks at the range of rows and columns and
'Shifts their colors so that all of the black ones
'are on the top of each column
Dim i As Integer
Dim j As Integer
Dim AllBlack As Boolean
Dim AllShifted As Boolean
j = 1
AllShifted = False
AllBlack = True
For i = Row - 1 To 0 Step -1
    If aryB(RCtoIndex(i, Col)).Color = 0 Then
        j = j + 1
    Else
        aryB(RCtoIndex(i + j, Col)).Color = aryB(RCtoIndex(i, Col)).Color
        AllBlack = False
    End If
Next i
If aryB(RCtoIndex(Row, Col)).Color <> 0 Then AllShifted = True
For i = 0 To j - 1
    aryB(RCtoIndex(i, Col)).Color = 0
Next i
If AllShifted Or AllBlack Then
    ShiftDown = False
Else
    ShiftDown = True
End If
End Function
Function ShiftLeft(Col As Integer) As Boolean
'This looks at the range of rows and columns and
'Shifts their colors so that if there is a completely
'black column, it moves to the right side of the 'board'
Dim i As Integer
Dim j As Integer
Dim AllShifted As Boolean
Dim AllBlack As Boolean
AllShifted = False
AllBlack = True
For i = Col + 1 To LastCol - 1
    For j = 0 To MaxRows - 1
        aryB(RCtoIndex(j, i - 1)).Color = aryB(RCtoIndex(j, i)).Color
        If aryB(RCtoIndex(j, i - 1)).Color <> 0 Then AllBlack = False
    Next j
Next i
For j = 0 To MaxRows - 1
    aryB(RCtoIndex(j, LastCol - 1)).Color = 0
Next j
If aryB(RCtoIndex(MaxRows - 1, Col)).Color <> 0 Then AllShifted = True
If AllShifted Or AllBlack Then
    LastCol = LastCol - 1
    ShiftLeft = False
Else
    ShiftLeft = True
End If
End Function
Sub SlideDown()
'This triggers the shift-down based on the range of 'triggered' blocks
'Triggered = selected (recursive)
Dim i As Integer
Dim j As Integer
Dim k As Integer
For i = LastCol - 1 To LowCol Step -1
    For j = HiRow To 0 Step -1
        k = RCtoIndex(j, i)
        If aryB(k).Color = 0 Then
            Do While ShiftDown(j, i) = True
            Loop
        End If
    Next j
Next i
End Sub

Sub SlideLeft()
'This initiates the sliding of black-columns to the right
Dim i As Integer
Dim j As Integer
Dim k As Integer
For i = LastCol - 1 To LowCol Step -1
    k = RCtoIndex(MaxRows - 1, i)
    If aryB(k).Color = 0 Then
        Do While ShiftLeft(i) = True
        Loop
    End If
Next i
End Sub
Function UpIndex(RootIndex As Integer) As Integer
'this looks at the block directly above the specified block-index
UpIndex = RCtoIndex(aryB(RootIndex).Row - 1, aryB(RootIndex).Col)
End Function

Function LeftIndex(RootIndex As Integer) As Integer
'this looks at the block directly left of the specified block-index
LeftIndex = RCtoIndex(aryB(RootIndex).Row, aryB(RootIndex).Col - 1)
End Function
Function RightIndex(RootIndex As Integer) As Integer
'this looks at the block directly right of the specified block-index
RightIndex = RCtoIndex(aryB(RootIndex).Row, aryB(RootIndex).Col + 1)
End Function
Function DownIndex(RootIndex As Integer) As Integer
'this looks at the block directly below the specified block-index
DownIndex = RCtoIndex(aryB(RootIndex).Row + 1, aryB(RootIndex).Col)
End Function







Private Sub Form_Load()
On Error Resume Next 'Just in case you set the 'game-board' larger than 32000 squares
MaxRows = GetSetting("SameGame", "Config", "MaxRows", 10)
MaxCols = GetSetting("SameGame", "Config", "MaxCols", 40)
MaxColors = GetSetting("SameGame", "Config", "MaxColors", 3)
BWidth = GetSetting("SameGame", "Config", "BWidth", 200)
BHeight = GetSetting("SameGame", "Config", "BHeight", 200)
Do While (MaxRows * MaxCols) > 32000 'this keeps refresh-time to a reasonable rate
    MaxRows = MaxRows - 1
    MaxCols = MaxCols - 1
Loop
'------------------set up our menu's--------------
'(This is really a lazy tactic to avoid other forms or sliding-panels.)
mnuRows.Caption = "Rows: " & MaxRows
mnuCols.Caption = "Cols: " & MaxCols
mnuColors.Caption = "Colors: " & MaxColors
mnuBWidth.Caption = "Block Width: " & BWidth
mnuBHeight.Caption = "Block Height: " & BHeight
DoEvents 'Jump-start the form load (very poor technique, but effective)
Generate 'This generates our arrays for the playing 'pieces'
End Sub

Private Sub Form_Resize()
picGrid.Height = Me.Height - 660
End Sub


Private Sub mnuBHeight_Click()
Dim response As String
response = InputBox("Enter the height of each block.", "Set Block Height", BHeight)
If response = "" Then Exit Sub
If Not IsNumeric(response) Then Exit Sub
BHeight = response
mnuBHeight.Caption = "Block Height: " & BHeight
SaveSetting "SameGame", "Config", "BHeight", BHeight
Generate
End Sub

Private Sub mnuBWidth_Click()
Dim response As String
response = InputBox("Enter the width of each block.", "Set Block Width", BWidth)
If response = "" Then Exit Sub
If Not IsNumeric(response) Then Exit Sub
BWidth = response
mnuBWidth.Caption = "Block Width: " & BWidth
SaveSetting "SameGame", "Config", "BWidth", BWidth
Generate
End Sub

Private Sub mnuColors_Click()
Dim response As String
response = InputBox("Enter the maximum number of colors.", "Set Colors", MaxColors)
If response = "" Then Exit Sub
If Not IsNumeric(response) Then Exit Sub
If response > 14 Then response = 14
MaxColors = response
mnuColors.Caption = "Colors: " & MaxColors
SaveSetting "SameGame", "Config", "MaxColors", MaxColors
Generate
End Sub

Private Sub mnuCols_Click()
Dim response As String
response = InputBox("Enter the maximum number of columns.", "Set Columns", MaxCols)
If response = "" Then Exit Sub
If Not IsNumeric(response) Then Exit Sub
MaxCols = response
mnuCols.Caption = "Cols: " & MaxCols
SaveSetting "SameGame", "Config", "MaxCols", MaxCols
Generate
End Sub

Private Sub mnuGen_Click()
Generate
End Sub

Private Sub mnuRows_Click()
Dim response As String
response = InputBox("Enter the maximum number of rows.", "Set Rows", MaxRows)
If response = "" Then Exit Sub
If Not IsNumeric(response) Then Exit Sub
MaxRows = response
mnuRows.Caption = "Rows: " & MaxRows
SaveSetting "SameGame", "Config", "MaxRows", MaxRows
Generate
End Sub


Private Sub mnuScores_Click()
On Error Resume Next
Dim i As Integer
For i = 0 To 9
    mnuScore(i).Caption = Format(GetSetting("SameGame", "Scores", "Points" & i), "!@@@@@@@@@@@@") & Format(GetSetting("SameGame", "Scores", "Name" & i), "!@@@@@@@@@@@@@@@@@@@@@@@") & GetSetting("SameGame", "Scores", "Level" & i)
Next i
End Sub

Private Sub picGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF2
        Generate
    Case vbKeyReturn
        DoSelect CurBlock
    Case vbKeyUp
        If KeyRow > 0 Then KeyRow = KeyRow - 1
        CurBlock = RCtoIndex(KeyRow, KeyCol)
        HiliteBlock CurBlock
    Case vbKeyDown
        If KeyRow < MaxRows - 1 Then KeyRow = KeyRow + 1
        CurBlock = RCtoIndex(KeyRow, KeyCol)
        HiliteBlock CurBlock
    Case vbKeyLeft
        If KeyCol > 0 Then KeyCol = KeyCol - 1
        CurBlock = RCtoIndex(KeyRow, KeyCol)
        HiliteBlock CurBlock
    Case vbKeyRight
        If KeyCol < MaxCols - 1 Then KeyCol = KeyCol + 1
        CurBlock = RCtoIndex(KeyRow, KeyCol)
        HiliteBlock CurBlock
End Select
End Sub

Private Sub picGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim BIndex As Integer
BIndex = GetAryIndex(X, Y)
If BIndex < 0 Then Exit Sub
KeyRow = aryB(BIndex).Row
KeyCol = aryB(BIndex).Col
HiliteBlock BIndex
End Sub

Private Sub picGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim BIndex As Integer
BIndex = GetAryIndex(X, Y)
DoSelect BIndex
'txtResult = aryB(MyIndex).Row & "," & aryB(MyIndex).Col
End Sub


