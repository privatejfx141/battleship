Attribute VB_Name = "modBattleship"
'Title: BATTLESHIP 2016
'Author: Yun Jie (Jeffrey) Li
'Date: May 19, 2016
'Files: Battleship.bas, About.frm, Exit.frm, Game.frm, GameOver.frm, Help.frm, Return.frm,
'       Splash.frm, Theatres.frm, respective .frx files.
'Purpose: The purpose of this application is to recreate the classical board game
'         'Battleship'. Also for bragging rights.

Global Const MAX = 10
Global Const SUNK = -2
Global Const HIT = -1
Global Const WATER = 0
Global Const CARRIER = 1
Global Const BCRUISER = 2
Global Const SUBMARINE = 3
Global Const CRUISER = 4
Global Const FRIGATE = 5
Global Const MISS = 6
Global Const REHIT = 7

Global Const EASY = 0
Global Const MEDIUM = 1
Global Const HARD = 2

Global Const SHIPCOLOR = 2629140
Global Const CMDBOXCOLOR = 10524260

Global GameTime As Integer
Global Hits As Integer
Global Misses As Integer
Global Accuracy As Single
Global Turn As Integer

Global TheatreFromMenu As Boolean
Global PassiveClose, PassiveCloseGame As Boolean
Global Difficulty As Integer
Global Victorious As Boolean

Global PCellValue(1 To MAX, 1 To MAX) As Integer
Global CCellValue(1 To MAX, 1 To MAX) As Integer

Dim PrevRow, PrevCol As Integer

Option Explicit

Public Sub EndProgram()

    frmExit.Show vbModal

End Sub

Public Function GetRandom(ByVal Low As Integer, ByVal High As Integer) As Integer

    GetRandom = Int(Rnd * (High - Low + 1)) + Low
    
End Function

Public Sub CentreForm(CurrentForm As Form)
    
    With CurrentForm
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
    
End Sub

Public Sub InitializeArrayValues(CellValue() As Integer)

    Dim CRow, CCol As Integer
    
    For CRow = 1 To MAX
        For CCol = 1 To MAX
            CellValue(CCol, CRow) = WATER
        Next CCol
    Next CRow

End Sub

Public Sub InitializeBoard(Grid As Control)

    Dim CRow, CCol As Integer
    
    With Grid
    
    .BackColor = RGB(72, 100, 117)
    .BackColorFixed = RGB(20, 30, 35)
    .BackColorBkg = RGB(62, 87, 101)
    
    For CRow = 0 To .Rows - 1
        .Row = CRow
        For CCol = 0 To .Cols - 1
            .Col = CCol
            .CellAlignment = 4
            If CRow > 0 And CCol > 0 Then
                .CellBackColor = RGB(72, 100, 117)
            End If
            .CellForeColor = vbWhite
            .CellPicture = LoadPicture()
            .Text = ""
        Next CCol
    Next CRow
    
    .Row = 0
    For CCol = 0 To .Cols - 1
        .ColWidth(CCol) = 480
        If CCol > 0 Then
            .Col = CCol
            .Text = CCol
        End If
    Next CCol
    
    .Col = 0
    For CRow = 0 To .Rows - 1
        If CRow > 0 Then
            .Row = CRow
            .Text = Chr$(64 + CRow)
        End If
    Next CRow
    
    .Row = 1
    .Col = 1
    
    End With
    
End Sub

Public Sub SetGameFormControls(Difficulty As Integer, Navy As Variant, PShip As Variant, CShip As Variant, _
    EFlag As Variant, Flag As Variant)

    Select Case Difficulty
        Case EASY
            Navy(0).Caption = "US Fourth Fleet"
            PShip(0).Caption = "NIMITZ"
            PShip(1).Caption = "TICONDEROGA"
            PShip(2).Caption = "LOS ANGELES"
            PShip(3).Caption = "ARLEIGH BURKE"
            PShip(4).Caption = "LCS"
            
            Flag(1) = EFlag(0)
            Navy(1).Caption = "Black Sea Fleet"
            CShip(0).Caption = "KUZNETSOV"
            CShip(1).Caption = "KIROV"
            CShip(2).Caption = "TYPHOON"
            CShip(3).Caption = "SLAVA"
            CShip(4).Caption = "TARANTUL"
        Case MEDIUM
            Navy(0).Caption = "US Seventh Fleet"
            PShip(0).Caption = "NIMITZ"
            PShip(1).Caption = "TICONDEROGA"
            PShip(2).Caption = "VIRGINIA"
            PShip(3).Caption = "ARLEIGH BURKE"
            PShip(4).Caption = "LCS"
            
            Flag(1) = EFlag(1)
            Navy(1).Caption = "South Sea Fleet"
            CShip(0).Caption = "LIAONING"
            CShip(1).Caption = "ZHEJIANG"
            CShip(2).Caption = "HAN"
            CShip(3).Caption = "SHANDONG"
            CShip(4).Caption = "JIANGDAO"
        Case HARD
            Navy(0).Caption = "USNORTHCOMM"
            PShip(0).Caption = "ENTERPRISE"
            PShip(1).Caption = "TICONDEROGA"
            PShip(2).Caption = "LOS ANGELES"
            PShip(3).Caption = "ARLEIGH BURKE"
            PShip(4).Caption = "LCS"
            
            Flag(1) = EFlag(2)
            Navy(1).Caption = "Northern Fleet"
            CShip(0).Caption = "SHTORM"
            CShip(1).Caption = "KIROV"
            CShip(2).Caption = "AKULA"
            CShip(3).Caption = "GORSHKOV"
            CShip(4).Caption = "GREMYASHCHY"
    End Select

End Sub

Public Function GetShipClass(Number As Integer) As String
    
    Dim ShipClass As String
    
    Select Case Number
        Case 0
            ShipClass = "Aircraft carrier"
        Case 1
            ShipClass = "Battlecruiser"
        Case 2
            ShipClass = "Missile submarine"
        Case 3
            ShipClass = "Cruiser"
        Case 4
            ShipClass = "Frigate"
    End Select
    
    GetShipClass = ShipClass
    
End Function

Public Function GetShipLength(ShipType As Integer) As Integer
    
    Dim ShipLength As Integer
    
    Select Case ShipType
        Case 0
            ShipLength = 5
        Case 1
            ShipLength = 4
        Case 2, 3
            ShipLength = 3
        Case 4
            ShipLength = 2
    End Select
    
    GetShipLength = ShipLength

End Function

Public Sub CheckAndPlaceShip(Grid As Control, CellValue() As Integer, ByVal ShipType As Integer, _
    Optional ByVal Direction As Integer = 0, Optional ByRef Valid As Boolean = True, _
    Optional ShowShips As Boolean = True)
        
    Const HORIZONTAL = 0, VERTICAL = 1
    Dim CRow, CCol As Integer
    Dim ORow, OCol As Integer
    Dim ShipLength As Integer
    Dim ShipValue As Integer

    ShipLength = GetShipLength(ShipType)
    ShipValue = ShipType + 1
    Valid = True
    
    With Grid
    
    ORow = .Row
    OCol = .Col
    
    If Direction = HORIZONTAL Then
    
        CCol = 0
        Do
        
            If (OCol + ShipLength - 1) > (.Cols - 1) Then
                .Col = OCol + CCol - (OCol + ShipLength - .Cols)
            Else
                .Col = OCol + CCol
            End If
            If CellValue(.Col, .Row) > WATER Then
                Valid = False
            End If
            
            If Valid Then
                CCol = CCol + 1
            End If
            
        Loop Until CCol = ShipLength Or Not Valid
        
        If Valid Then
            For CCol = 0 To ShipLength - 1
                If (OCol + ShipLength - 1) > (.Cols - 1) Then
                    .Col = OCol + CCol - (OCol + ShipLength - .Cols)
                Else
                    .Col = OCol + CCol
                End If
                CellValue(.Col, .Row) = ShipValue
                If ShowShips Then
                    .CellBackColor = SHIPCOLOR
                    '.Text = ShipValue
                End If
            Next CCol
        End If
        
    ElseIf Direction = VERTICAL Then
    
        CRow = 0
        Do
        
            If (ORow + ShipLength - 1) > (.Rows - 1) Then
                .Row = ORow + CRow - (ORow + ShipLength - .Rows)
            Else
                .Row = ORow + CRow
            End If
            If CellValue(.Col, .Row) > WATER Then
                Valid = False
            End If
            
            If Valid Then
                CRow = CRow + 1
            End If
            
        Loop Until CRow = ShipLength Or Not Valid
        
        If Valid Then
            For CRow = 0 To ShipLength - 1
                If (ORow + ShipLength - 1) > (.Rows - 1) Then
                    .Row = ORow + CRow - (ORow + ShipLength - .Rows)
                Else
                    .Row = ORow + CRow
                End If
                CellValue(.Col, .Row) = ShipValue
                If ShowShips Then
                    .CellBackColor = SHIPCOLOR
                    '.Text = ShipValue
                End If
            Next CRow
        End If
        
    End If
    
    End With

End Sub

Public Function GetDialogMsg(Optional Number As Integer = 0) As String

    Dim Msg As String
    
    Select Case Number
        Case 0
            Msg = "Select a ship and place it in the player's grid, "
            Msg = Msg & "or select 'Random Placement', to begin."
        Case 4
            Msg = "All ships have been deployed on the grid. "
            Msg = Msg & "Select 'Begin Operation' to commence engagement."
        Case 5
            Msg = "Enemy ships detected. All hands man your battle stations."
        Case 6
            Msg = "Hostile forces are firing upon our fleet. Retalitate back."
        Case 7
            Msg = "We have eliminated all hostile vessels. Region is secured. We are victorious."
        Case 8
            Msg = "The enemy fleet has sunken all our ships. Order a retreat now."
    End Select
    
    GetDialogMsg = Msg
    
End Function

Public Sub RandomPlaceShips(Grid As Control, CellValue() As Integer, _
    Optional ShowShips As Boolean = True)
    
    Dim ShipType As Integer
    Dim Direction As Integer
    Dim IsValid As Boolean
    
    With Grid
    
        Randomize
        For ShipType = 0 To 4
            
            Do
                .Row = GetRandom(1, 10)
                .Col = GetRandom(1, 10)
                Direction = GetRandom(0, 1)
                CheckAndPlaceShip Grid, CellValue, ShipType, Direction, IsValid, ShowShips
            Loop While Not IsValid
            
        Next ShipType

    End With

End Sub

Public Function PlayerFireTurnResult(Grid As Control, DMsg As String, _
    Hits As Integer, Misses As Integer, CShip As Variant, CShipIcon As Variant, _
    Colour As Variant) As Integer
    
    Dim K As Integer
    Dim Result As Integer
    Dim RemainingTiles As Integer
    Dim CRow, CCol As Integer
    Dim ShipType As Integer
    Dim ShipValue As Integer
    
    RemainingTiles = 0
    With Grid
        
        .Visible = False
        
        If .CellPicture.Handle = 0 Then
        
            If CCellValue(.Col, .Row) > 0 Then
                .CellPicture = Colour(1).Picture
                ShipType = CCellValue(.Col, .Row) - 1
                ShipValue = ShipType + 1
                CCellValue(.Col, .Row) = CCellValue(.Col, .Row) * HIT
                
                For CRow = 1 To .Rows - 1
                    For CCol = 1 To .Cols - 1
                        If CCellValue(CCol, CRow) = ShipValue Then
                            RemainingTiles = RemainingTiles + 1
                        End If
                    Next CCol
                Next CRow
                
                If RemainingTiles = 0 Then
                    DMsg = "We have sunken the enemy's "
                    DMsg = DMsg & GetShipClass(ShipType) & "."
                    Result = SUNK
                    CShip(ShipType).ForeColor = vbRed
                    CShipIcon(ShipType).Visible = False
                Else
                    DMsg = "We have a confirmed hit on an enemy target."
                    Result = HIT
                End If
                Hits = Hits + 1
                
            ElseIf CCellValue(.Col, .Row) = WATER Then
                .CellPicture = Colour(0).Picture
                DMsg = "Miss. No effect on target."
                Result = MISS
                Misses = Misses + 1
            End If
            
        Else
            DMsg = "That position was already fired upon. Select another."
            Result = REHIT
        End If
        
        .Visible = True
        
    End With
    
    PlayerFireTurnResult = Result
    
End Function

Public Function ComputerFireTurnResult(Grid As Control, DMsg As String, PShip As Variant, _
    PShipIcon As Variant, Colour As Variant, Optional Difficulty As Integer = EASY) As Integer
    
    Dim Result As Integer
    Dim RemainingTiles As Integer
    Dim CRow, CCol As Integer
    Dim ShipType As Integer
    Dim ShipValue As Integer
    Dim REHIT As Boolean
    Dim CoinFlip As Integer
    
    REHIT = False
    
    With Grid
    
    .Visible = False
            
    Do
        REHIT = False
        
        Select Case Difficulty
            Case EASY
                .Row = GetRandom(1, 10)
                .Col = GetRandom(1, 10)
            Case MEDIUM
                CoinFlip = GetRandom(0, 1)
                If CoinFlip = 0 Then
                    .Row = GetRandom(1, 10)
                    .Col = GetRandom(1, 10)
                Else
                    If PrevRow - 1 >= 1 Then
                        .Row = PrevRow - 1
                    End If
                End If
            Case HARD
                .Row = GetRandom(1, 10)
                .Col = GetRandom(1, 10)
        End Select
        
        If .CellPicture.Handle = 0 Then
            If PCellValue(.Col, .Row) > WATER Then
                
                ShipValue = PCellValue(.Col, .Row)
                ShipType = ShipValue - 1
                PCellValue(.Col, .Row) = PCellValue(.Col, .Row) * HIT
                For CRow = 1 To .Rows - 1
                    For CCol = 1 To .Cols - 1
                        If PCellValue(CCol, CRow) = ShipValue Then
                            RemainingTiles = RemainingTiles + 1
                        End If
                    Next CCol
                Next CRow
                
                PrevRow = .Row
                PrevCol = .Col
                .CellPicture = Colour(1).Picture
                If RemainingTiles = 0 Then
                    DMsg = "The enemy has sunken our "
                    DMsg = DMsg & GetShipClass(ShipType) & "."
                    Result = SUNK
                    PShip(ShipType).ForeColor = vbRed
                    PShipIcon(ShipType).Visible = False
                Else
                    PShip(ShipType).ForeColor = RGB(255, 165, 0)
                    DMsg = "Our " & GetShipClass(ShipType)
                    DMsg = DMsg & " has been hit by the enemy."
                    Result = HIT
                End If
            Else
                .CellPicture = Colour(0).Picture
                Result = MISS
            End If
        Else
            REHIT = True
        End If
        
    Loop While REHIT
    
    .Visible = True
    
    End With
    
    ComputerFireTurnResult = Result
    
End Function
