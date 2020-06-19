VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' >> Declare variables <<
#If VBA7 Then ' Excel 2010 or later
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long
#Else ' -> Todo: test if this works
    Private Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long
#End If

'   -> Storage
Dim myExcel As Object
Dim Scene As Object

Dim Art As Variant
Dim Map As Variant


'   -> Game
Dim PosX As Integer
Dim PosY As Integer

Dim LastX As Integer
Dim LastY As Integer

Dim Speed As Double
Dim Score As Integer

Dim PosXRaw As Integer
Dim PosYRaw As Integer

'   -> Scene
Dim CanvasX As Integer
Dim CanvasY As Integer

Dim SceneX As Integer
Dim SceneY As Integer

   

'   -> Input
Dim KeyMap(255) As Boolean
Private Sub Document_Open()
    Dim start As Integer
    start = MsgBox("Click 'Yes' to start the game.", vbYesNo, "Start Game?")
    If start = vbYes Then
        ActiveDocument.ActiveWindow.WindowState = wdWindowStateMaximize
        Main
    Else
        Application.Quit SaveChanges:=wdDoNotSaveChanges
    End If
End Sub
Sub Main()
    CanvasX = 40
    CanvasY = 20
    
    With ActiveDocument.ActiveWindow.View.Zoom
        .PageFit = wdPageFitFullPage
        .Percentage = .Percentage + 5
    End With
    
    Debug.Print "Initializing..."
    Set myExcel = CreateObject("Excel.Application")
    Set Scene = myExcel.Workbooks.Open(ActiveDocument.Path & "\scene.xlsx")
    
    Debug.Print "Initialization complete."
    
    ImportScene
    Update
End Sub
Sub Fire()
    Dim Output As String
    Dim x, y, t As Integer
    
    For y = 1 To CanvasY
        For x = 1 To CanvasX
            Output = Output & GetEmoji(5)
        Next x
        If (y <> CanvasX + 1) Then
            Output = Output & Chr(10)
        End If
    Next y
    Canvas.Caption = Output
End Sub
Sub GameOver()
    Dim continue As Integer
    Fire
    continue = MsgBox("Your score is " & Score & Chr(10) & "Do you want to play again?", vbYesNo, "Game Over")
    If continue = vbYes Then
        Canvas.Caption = ""
        ImportScene
    Else
        Application.Quit SaveChanges:=wdDoNotSaveChanges
    End If
End Sub
'   -> TODO: Error handler
Sub ImportScene()
    Speed = 0.6
    Score = 0
    
    PosXRaw = 0
    PosYRaw = 0
    
    PosX = -CanvasX / 2 + 2
    PosY = CanvasY / 2 - 2
    
    LastX = -1
    LastY = -1
    
    With Scene.Worksheets("Map").Range("A1").CurrentRegion
        SceneX = .Rows.Count
        SceneY = .Columns.Count
    End With

    Art = Scene.Worksheets("Art").Range("B2:B7").Value
    Map = Scene.Worksheets("Map").Range("A1:A1").Resize(SceneX, SceneY).Value
    
    ScoreValue.Caption = "Score: " & Score
    
    Render
End Sub
Public Function GetEmoji(ByVal Index As Integer) As String
    GetEmoji = Art(Index + 1, 1)
End Function
Sub Render()
    Dim Output As String
    Dim x, y, t As Integer
    
    For y = 1 To CanvasY
        For x = 1 To CanvasX + 1
            If (x = CanvasX / 2 And y = CanvasY / 2) Then
                Output = Output & GetEmoji(2)
            Else
                If (x + PosX <= 0 Or x + PosX >= SceneX Or y - PosY <= 0 Or y - PosY >= SceneY) Then
                    Output = Output & GetEmoji(0)
                Else
                    Output = Output & GetEmoji(CInt(Map(x + PosX, y - PosY)))
                End If
            End If
        Next x
        If (y <> CanvasX + 1) Then
            Output = Output & Chr(10)
        End If
    Next y
    Canvas.Caption = Output
End Sub
Sub UpdatePosition()
    If (KeyMap(87) Or KeyMap(38)) Then
        If (CInt(Map(CanvasX / 2 + PosX, CanvasY / 2 - PosY - 1)) <> 0) Then
            PosYRaw = PosYRaw + 1
        End If
    End If
    If (KeyMap(65) Or KeyMap(37)) Then
        If (CInt(Map(CanvasX / 2 + PosX - 1, CanvasY / 2 - PosY)) <> 0) Then
            PosXRaw = PosXRaw - 1
        End If
    End If
    If (KeyMap(83) Or KeyMap(40)) Then
        If (CInt(Map(CanvasX / 2 + PosX, CanvasY / 2 - PosY + 1)) <> 0) Then
            PosYRaw = PosYRaw - 1
        End If
    End If
    If (KeyMap(68) Or KeyMap(39)) Then
        If (CInt(Map(CanvasX / 2 + PosX + 1, CanvasY / 2 - PosY)) <> 0) Then
            PosXRaw = PosXRaw + 1
        End If
    End If
    PosX = PosXRaw * Speed - CanvasX / 2 + 2
    PosY = PosYRaw * Speed + CanvasY / 2 - 2
End Sub
Sub Events()
    If (CInt(Map(CanvasX / 2 + PosX, CanvasY / 2 - PosY)) = 3) Then
        GameOver
    End If
    If (CInt(Map(CanvasX / 2 + PosX, CanvasY / 2 - PosY)) = 4) Then
        Score = Score + 1
        ScoreValue.Caption = "Score: " & Score
        Map(CanvasX / 2 + PosX, CanvasY / 2 - PosY) = 1
    End If
End Sub
Sub Update()
    UpdatePosition
    
    TextInput.Text = ""
    If (PosX <> LastX Or PosY <> LastY) Then
        Events
        Render
        
        LastX = PosX
        LastY = PosY
    End If
    
End Sub
Private Sub TextInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyMap(KeyCode) = True
    Update
End Sub
Private Sub TextInput_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyMap(KeyCode) = False
    Update
End Sub
