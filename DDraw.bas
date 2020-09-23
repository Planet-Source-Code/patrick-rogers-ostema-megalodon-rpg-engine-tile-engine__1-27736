Attribute VB_Name = "DDraw"
Option Explicit
'DirectX variables
Dim mdd As DirectDraw7
Dim msurfFront As DirectDrawSurface7
Dim msurfBack As DirectDrawSurface7
Dim ElSprite As DirectDrawSurface7
Dim msurfTiles As DirectDrawSurface7
Dim TalkBox As DirectDrawSurface7
Dim Itemz As DirectDrawSurface7
Const SPEECH_WIDTH = 640
Const SPEECH_HEIGHT = 120
Const SPEECH_START_HEIGHT = 32
Const SPEECH_START_WIDTH = 15
Const SPEECH_LINE_SPACING = 20
Dim mlngFrameTime As Long                   'How long since last frame?
Dim mlngTimer As Long                       'How long since last FPS count update?
Dim mintFPSCounter As Integer               'Our FPS counter
Dim mintFPS As Integer                      'Our FPS storage variable
Dim mrectScreen As RECT
Public SaleLoop As Byte
Public PointerX As Integer
Public PointerY As Integer

Public Sub DrawTiles()
Dim i As Integer
Dim j As Integer
Dim rectTile As RECT
Dim bytTileNum As Tilez
Dim bytTileNum2 As Tilez
Dim intx As Integer
Dim inty As Integer
Dim intItem As Byte
    'Draw the tiles according to the map array
    For i = 0 To CInt(SCREEN_WIDTH / TILE_WIDTH)
        For j = 0 To CInt(SCREEN_HEIGHT / TILE_HEIGHT)
            'Calc X,Y coords for this tile's placement
            intx = i * TILE_WIDTH - mintX Mod TILE_WIDTH
            inty = j * TILE_HEIGHT - mintY Mod TILE_HEIGHT
            bytTileNum = GetTileG(intx, inty)
            bytTileNum2 = GetTileF(intx, inty)
            intItem = GetItem(intx, inty)
            GetRect bytTileNum, intx, inty, rectTile
            msurfBack.BltFast intx, inty, msurfTiles, rectTile, DDBLTFAST_WAIT
            If bytTileNum2.X <> 0 Or bytTileNum2.Y <> 0 Then
            intx = i * TILE_WIDTH - mintX Mod TILE_WIDTH
            inty = j * TILE_HEIGHT - mintY Mod TILE_HEIGHT
            GetRect bytTileNum2, intx, inty, rectTile
            msurfBack.BltFast intx, inty, msurfTiles, rectTile, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            End If
            If intItem > 0 Then
            intx = i * TILE_WIDTH - mintX Mod TILE_WIDTH
            inty = j * TILE_HEIGHT - mintY Mod TILE_HEIGHT
            intItem = DaItems(intItem - 1).GrhIndex
            'Get the rectangle
            GetItemRect intItem, intx, inty, rectTile
            'Blit the tile
            msurfBack.BltFast intx, inty, Itemz, rectTile, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            End If
            Next j
    Next i

End Sub

Private Function GetTileG(intTileX As Integer, intTileY As Integer) As Tilez

    'Return the value returned by the map array for the given tile
    GetTileG = mbytMap((intTileX + TILE_WIDTH \ 2 + mintX - SCREEN_WIDTH \ 2) \ TILE_WIDTH, (intTileY + TILE_HEIGHT \ 2 + mintY - SCREEN_HEIGHT \ 2) \ TILE_HEIGHT).TileNumber.Ground

End Function

Private Function GetTileF(intTileX As Integer, intTileY As Integer) As Tilez

    'Return the value returned by the map array for the given tile
    GetTileF = mbytMap((intTileX + TILE_WIDTH \ 2 + mintX - SCREEN_WIDTH \ 2) \ TILE_WIDTH, (intTileY + TILE_HEIGHT \ 2 + mintY - SCREEN_HEIGHT \ 2) \ TILE_HEIGHT).TileNumber.Floor

End Function
Private Function GetTileS(X As Integer, Y As Integer) As Tilez

    'Return the value returned by the map array for the given tile
    GetTileS = mbytMap(X, Y).TileNumber.Sky

End Function
Private Sub GetRect(bytTileNumber As Tilez, ByRef intTileX As Integer, ByRef intTileY As Integer, ByRef rectTile As RECT)

    'Calc rect
    With rectTile
        .Left = bytTileNumber.X * TILE_WIDTH
        .Right = .Left + TILE_WIDTH
        .Top = bytTileNumber.Y * TILE_HEIGHT
        .Bottom = .Top + TILE_HEIGHT
    
    'Clip rect
        
        'If this tile is off the left side of the screen...
        If intTileX < 0 Then
            .Left = .Left - intTileX
            intTileX = 0
        End If
        'If this tile is off the top of the screen...
        If intTileY < 0 Then
            .Top = .Top - intTileY
            intTileY = 0
        End If
        'If this tile is off the right side of the screen...
        If intTileX + TILE_WIDTH > SCREEN_WIDTH Then .Right = .Right + (SCREEN_WIDTH - (intTileX + TILE_WIDTH))
        'If this tile is off the bottom of the screen...
        If intTileY + TILE_HEIGHT > SCREEN_HEIGHT Then .Bottom = .Bottom + (SCREEN_HEIGHT - (intTileY + TILE_HEIGHT))
    End With

End Sub

Public Sub MoveScreen()
msurfBack.BltColorFill mrectScreen, 0
    'Move screen
    If Walking = South Then
        If DudeCoord.Y < 8 Then GoTo MoveIt:
    mintY = mintY + SCROLL_SPEED
    GoTo MoveIt:
    End If
    If Walking = North Then
        If DudeCoord.Y > 41 Then GoTo MoveIt:
    mintY = mintY - SCROLL_SPEED
    GoTo MoveIt:
    End If
    If Walking = West Then
        If DudeCoord.X > 39 Then GoTo MoveIt:
    mintX = mintX - SCROLL_SPEED
    GoTo MoveIt:
    End If
    If Walking = East Then
        If DudeCoord.X < 10 Then GoTo MoveIt:
    mintX = mintX + SCROLL_SPEED
    GoTo MoveIt:
    End If
'Ensure we don't go off the edge, and move NPC Accordingly if so.
MoveIt:
    If mintX < SCREEN_WIDTH \ 2 Then
    mintX = SCREEN_WIDTH \ 2
    End If
    If mintX > UBound(mbytMap, 1) * TILE_WIDTH - SCREEN_WIDTH \ 2 Then
    mintX = UBound(mbytMap, 1) * TILE_WIDTH - SCREEN_WIDTH \ 2
    End If
    If mintY < SCREEN_HEIGHT \ 2 Then
    mintY = SCREEN_HEIGHT \ 2
    End If
    If mintY > UBound(mbytMap, 2) * TILE_HEIGHT - SCREEN_HEIGHT \ 2 Then
    mintY = UBound(mbytMap, 2) * TILE_HEIGHT - SCREEN_HEIGHT \ 2
    End If
End Sub

Public Sub FPS()
    'Count FPS
    If mlngTimer + 1000 <= Gdx.TickCount Then
        mlngTimer = Gdx.TickCount
        mintFPS = mintFPSCounter + 1
        mintFPSCounter = 0
    Else
        mintFPSCounter = mintFPSCounter + 1
    End If
    'Display FPS, text, and possibly NPC Speech

    msurfBack.DrawText 0, 0, "Megalodon's RPG Engine", False
    msurfBack.DrawText 0, 20, "FPS: " & mintFPS, False
    msurfBack.DrawText 0, 40, "X= " & DudeCoord.X & "Y= " & DudeCoord.Y, False
    If TradeNPC = 1 Then DisplayNPCQuery
    If SayNPC Then DisplaySpeech
    If TradeNPC = 2 Then DrawTradeMenu
    If TradeNPC = 3 Then DrawTradeMenu
    If DispInventMenu = True Then DrawInventMenu
    DrawHealth
    msurfFront.Flip Nothing, DDFLIP_WAIT
    End Sub
Public Sub LoadSurfaces()
Dim CKey As DDCOLORKEY
Dim ddsdGeneric As DDSURFACEDESC2
Dim tLocked As Long
Dim TempRct As RECT
    'Set up generic surface description
    ddsdGeneric.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsdGeneric.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    'Load our tileset
    ddsdGeneric.lHeight = 352
    ddsdGeneric.lWidth = 768
    Set msurfTiles = mdd.CreateSurfaceFromFile(App.Path & "\tileset.bmp", ddsdGeneric)
    ddsdGeneric.lHeight = 768
    ddsdGeneric.lWidth = 256
    Set ElSprite = mdd.CreateSurfaceFromFile(App.Path & "\Pat.bmp", ddsdGeneric)
    With TempRct
        .Bottom = 384
        .Top = 0
        .Left = 0
        .Right = 192
    End With
    ElSprite.Lock TempRct, ddsdGeneric, DDLOCK_NOSYSLOCK, frmMain.hWnd
    tLocked = ElSprite.GetLockedPixel(0, 0)
    ElSprite.Unlock TempRct
    ddsdGeneric.lHeight = 885
    ddsdGeneric.lWidth = 640
    Set TalkBox = mdd.CreateSurfaceFromFile(App.Path & "\ChatBox.bmp", ddsdGeneric)
    ddsdGeneric.lHeight = 32
    ddsdGeneric.lWidth = 800
    Set Itemz = mdd.CreateSurfaceFromFile(App.Path & "\Itemz.bmp", ddsdGeneric)
CKey.low = tLocked
CKey.high = tLocked
msurfTiles.SetColorKey DDCKEY_SRCBLT, CKey
ElSprite.SetColorKey DDCKEY_SRCBLT, CKey
TalkBox.SetColorKey DDCKEY_SRCBLT, CKey
Itemz.SetColorKey DDCKEY_SRCBLT, CKey
End Sub

Public Function ExclusiveMode() As Boolean

Dim lngTestExMode As Long
    
    'This function tests if we're still in exclusive mode
    lngTestExMode = mdd.TestCooperativeLevel
    
    If (lngTestExMode = DD_OK) Then
        ExclusiveMode = True
    Else
        ExclusiveMode = False
    End If
    
End Function

Public Function LostSurfaces() As Boolean

    'This function will tell if we should reload our bitmaps or not
    LostSurfaces = False
    Do Until ExclusiveMode
        DoEvents
        LostSurfaces = True
    Loop
    
    'If we did lose our bitmaps, restore the surfaces and return 'true'
    DoEvents
    If LostSurfaces Then
        mdd.RestoreAllSurfaces
    End If
    
End Function

Public Sub Terminate()
    'Terminate the render loop
    mblnRunning = False

    'Restore resolution
    mdd.RestoreDisplayMode
    mdd.SetCooperativeLevel 0, DDSCL_NORMAL

    'Kill the surfaces
    Set Itemz = Nothing
    Set TalkBox = Nothing
    Set msurfTiles = Nothing
    Set ElSprite = Nothing
    Set msurfBack = Nothing
    Set msurfFront = Nothing
    'Kill directdraw
    Set mdd = Nothing
    'Kill DirectX
    Set Gdx = Nothing
End Sub



Public Sub Init()
Dim ddsdMain As DDSURFACEDESC2
Dim ddsdFlip As DDSURFACEDESC2    'Show the main form
    frmMain.Show

    'Initialize DirectDraw
    Set mdd = Gdx.DirectDrawCreate("")
    
    'Set the cooperative level (Fullscreen exclusive)
    mdd.SetCooperativeLevel frmMain.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
    
    'Set the resolution
    mdd.SetDisplayMode SCREEN_WIDTH, SCREEN_HEIGHT, SCREEN_BITDEPTH, 0, DDSDM_DEFAULT

    'Describe the flipping chain architecture we'd like to use
    ddsdMain.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsdMain.lBackBufferCount = 1
    ddsdMain.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_PRIMARYSURFACE
    
    'Create the primary surface
    Set msurfFront = mdd.CreateSurface(ddsdMain)
    
    'Create the backbuffer
    ddsdFlip.ddsCaps.lCaps = DDSCAPS_BACKBUFFER
    Set msurfBack = msurfFront.GetAttachedSurface(ddsdFlip.ddsCaps)
    
    'Set the text colour for the backbuffer
    msurfBack.SetForeColor vbWhite
    msurfBack.SetFontTransparency True

    'Create our screen-sized rectangle
    mrectScreen.Bottom = SCREEN_HEIGHT
    mrectScreen.Right = SCREEN_WIDTH
YCharOffSet = 7
XCharOffSet = 10
PointerX = 275
PointerY = 113
SaleLoop = 0
    LoadSurfaces
End Sub
Private Sub DisplaySpeech()
Dim SrcRect As RECT
Dim strTemp As String
Dim strOne As String
Dim strTwo As String
Dim strThree As String
Dim strFour As String
    With SrcRect
    .Left = 0
    .Right = 640
    .Top = 0
    .Bottom = 100
    End With
    msurfBack.BltFast 0, 380, TalkBox, SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    'Extract NPC Name and Text
    strTemp = NPCTalk
    If NPCFirst Then
        If strTemp <> "" Then TempName = Left(strTemp, InStr(1, strTemp, ":") - 1)
        strTemp = Right(strTemp, Len(strTemp) - InStr(1, strTemp, ":"))
        End If
    If strTemp <> "" Then strOne = Left(strTemp, InStr(1, strTemp, "|") - 1)
    If strTemp <> "" Then strTemp = Right(strTemp, Len(strTemp) - InStr(1, strTemp, "|"))
    If strTemp <> "" Then strTwo = Left(strTemp, InStr(1, strTemp, "|") - 1)
    If strTemp <> "" Then strTemp = Right(strTemp, Len(strTemp) - InStr(1, strTemp, "|"))
    If strTemp <> "" Then strThree = Left(strTemp, InStr(1, strTemp, "|") - 1)
    If strTemp <> "" Then strTemp = Right(strTemp, Len(strTemp) - InStr(1, strTemp, "|"))
    If strTemp <> "" Then strFour = Left(strTemp, InStr(1, strTemp, "|") - 1)
    If strTemp <> "" Then
    strTemp = Right(strTemp, Len(strTemp) - InStr(1, strTemp, "|"))
    Else
    strTemp = ""
    End If
    strOne = TempName & ": " & strOne
    'Ensure that these don't overlap the next "page" of text, and display them
    If InStr(1, strOne, "~") = 0 And strOne <> "" Then
        msurfBack.DrawText SPEECH_START_WIDTH, 480 - SPEECH_HEIGHT + SPEECH_START_HEIGHT, strOne, False
        If InStr(1, strTwo, "~") = 0 And strTwo <> "" Then
            msurfBack.DrawText SPEECH_START_WIDTH + 55, 480 - SPEECH_HEIGHT + SPEECH_START_HEIGHT + SPEECH_LINE_SPACING, strTwo, False
            If InStr(1, strThree, "~") = 0 And strThree <> "" Then
                msurfBack.DrawText SPEECH_START_WIDTH + 55, 480 - SPEECH_HEIGHT + SPEECH_START_HEIGHT + 2 * SPEECH_LINE_SPACING, strThree, False
                If InStr(1, strFour, "~") = 0 And strFour <> "" Then msurfBack.DrawText SPEECH_START_WIDTH + 55, 480 - SPEECH_HEIGHT + SPEECH_START_HEIGHT + 3 * SPEECH_LINE_SPACING, strFour, False
            End If
        End If
    End If
    NPCNext = strTemp
End Sub
Public Sub DrawNPC()
Dim i As Integer
Dim j As Integer
Dim rectTile As RECT
Dim intx As Integer
Dim inty As Integer
Dim blnNPC As Boolean
Dim TempNPCX As Integer
Dim TempNPCY As Integer

    'Draw the tiles according to the map array
    For i = 0 To CInt(SCREEN_WIDTH / TILE_WIDTH)
        For j = 0 To CInt(SCREEN_HEIGHT / TILE_HEIGHT)
            
            'Calc X,Y coords for this tile's placement
            intx = i * TILE_WIDTH - mintX Mod TILE_WIDTH
            inty = j * TILE_HEIGHT - mintY Mod TILE_HEIGHT
            blnNPC = GetNPC(intx, inty)
            If blnNPC Then
            'Get the rectangle
            TempNPCX = (intx + TILE_WIDTH \ 2 + mintX - SCREEN_WIDTH \ 2) \ TILE_WIDTH
            TempNPCY = (inty + TILE_HEIGHT \ 2 + mintY - SCREEN_HEIGHT \ 2) \ TILE_HEIGHT
            intx = intx + NPCz(TempNPCX, TempNPCY).MoveX
            inty = inty + NPCz(TempNPCX, TempNPCY).MoveY
            GetNPCRect intx, inty, rectTile, TempNPCX, TempNPCY
            'Blit the tile
            msurfBack.BltFast intx, inty, ElSprite, rectTile, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            End If

            Next j
    Next i
DrawSky
End Sub
Private Function GetNPC(ByRef intTileX As Integer, ByRef intTileY As Integer) As Boolean
Dim TempNPCX As Integer
Dim TempNPCY As Integer
            TempNPCX = (intTileX + TILE_WIDTH \ 2 + mintX - SCREEN_WIDTH \ 2) \ TILE_WIDTH
            TempNPCY = (intTileY + TILE_HEIGHT \ 2 + mintY - SCREEN_HEIGHT \ 2) \ TILE_HEIGHT
    'Return the value returned by the map array for the given tile
    GetNPC = mbytMap(TempNPCX, TempNPCY).NPC
If TempNPCX > 50 Then Exit Function
If NPCz(TempNPCX, TempNPCY).Bubbz.Damage > 0 Then
If TempNPCX = DudeCoord.X And TempNPCY = DudeCoord.Y Then
msurfBack.SetForeColor vbRed
msurfBack.DrawText intTileX + NPCz(TempNPCX, TempNPCY).MoveX + 10 + NPCz(TempNPCX, TempNPCY).Bubbz.Xaxis, intTileY + NPCz(TempNPCX, TempNPCY).MoveY + NPCz(TempNPCX, TempNPCY).Bubbz.Yaxis, NPCz(TempNPCX, TempNPCY).Bubbz.Damage, False
msurfBack.SetForeColor vbWhite
Else: msurfBack.DrawText intTileX + NPCz(TempNPCX, TempNPCY).MoveX + 10 + NPCz(TempNPCX, TempNPCY).Bubbz.Xaxis, intTileY + NPCz(TempNPCX, TempNPCY).MoveY + NPCz(TempNPCX, TempNPCY).Bubbz.Yaxis, NPCz(TempNPCX, TempNPCY).Bubbz.Damage, False
End If
End If
If GetNPC Then Exit Function
If TempNPCY >= 50 Then Exit Function
If TempNPCY <= 0 Then Exit Function
If TempNPCX >= 50 Then Exit Function
If TempNPCX <= 0 Then Exit Function
If mbytMap(TempNPCX + 1, TempNPCY).NPC And TempNPCX = DudeCoord.X + XCharOffSet And NPCz(TempNPCX + 1, TempNPCY).Walking = East Then
GetNPC = True
intTileX = intTileX + 32
Exit Function
End If
If mbytMap(TempNPCX - 1, TempNPCY).NPC And TempNPCX = DudeCoord.X - XCharOffSet + 1 And NPCz(TempNPCX - 1, TempNPCY).Walking = West Then
GetNPC = True
intTileX = intTileX - 32
Exit Function
End If
If mbytMap(TempNPCX, TempNPCY - 1).NPC And TempNPCY = DudeCoord.Y - YCharOffSet And NPCz(TempNPCX, TempNPCY - 1).Walking = North Then
GetNPC = True
intTileY = intTileY - 32
Exit Function
End If
If mbytMap(TempNPCX, TempNPCY + 1).NPC And TempNPCY = DudeCoord.Y + YCharOffSet + 1 And NPCz(TempNPCX, TempNPCY + 1).Walking = South Then
GetNPC = True
intTileY = intTileY + 32
Exit Function
End If
End Function
Private Sub GetNPCRect(ByRef intTileX As Integer, ByRef intTileY As Integer, ByRef rectTile As RECT, TempNPCX As Integer, TempNPCY As Integer)
If NPCz(TempNPCX, TempNPCY).Walking = 0 Then
    If NPCz(TempNPCX, TempNPCY).State <> Attacking Then
        NPCz(TempNPCX, TempNPCY).Frame = NPCz(TempNPCX, TempNPCY).LastStep
    End If
End If
If NPCz(TempNPCX, TempNPCY).Walking <> 0 Or NPCz(TempNPCX, TempNPCY).State = Attacking Then
NPC.SetNPCFrame TempNPCX, TempNPCY
End If
'Calc rect
    With rectTile
        .Left = mbytMap(TempNPCX, TempNPCY).Sprite * 32
        .Right = mbytMap(TempNPCX, TempNPCY).Sprite * 32 + 32
        .Top = NPCz(TempNPCX, TempNPCY).Frame * 32
        .Bottom = NPCz(TempNPCX, TempNPCY).Frame * 32 + 32
    
    'Clip rect
        
        'If this tile is off the left side of the screen...
        If intTileX < 0 Then
            .Left = .Left - intTileX
            intTileX = 0
        End If
        'If this tile is off the top of the screen...
        If intTileY < 0 Then
            .Top = .Top - intTileY
            intTileY = 0
        End If
        'If this tile is off the right side of the screen...
        If intTileX + TILE_WIDTH > SCREEN_WIDTH Then .Right = .Right + (SCREEN_WIDTH - (intTileX + TILE_WIDTH))
        'If this tile is off the bottom of the screen...
        If intTileY + TILE_HEIGHT > SCREEN_HEIGHT Then .Bottom = .Bottom + (SCREEN_HEIGHT - (intTileY + TILE_HEIGHT))
    End With
End Sub


Private Sub DisplayNPCQuery()
Dim SrcRect As RECT
    With SrcRect
    .Left = 0
    .Right = 379
    .Top = 101
    .Bottom = 162
    End With
    msurfBack.BltFast 260, 1, TalkBox, SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
With SrcRect
    .Left = 380
    .Right = 416
    .Top = 115
    .Bottom = 152
End With
    msurfBack.BltFast PointerX, 10, TalkBox, SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End Sub

Private Sub DrawTradeMenu()
Dim TempY As Integer
Dim SrcRect As RECT
Dim IntCount As Byte
Dim TempNum As Integer
Dim TempName As String
Dim TempCost As Integer
Dim TempAmount As Integer
    With SrcRect
    .Left = 0
    .Right = 392
    .Top = 162
    .Bottom = 540
    End With
msurfBack.BltFast 120, 50, TalkBox, SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    With SrcRect
    .Left = 392
    .Right = 419
    .Top = 472
    .Bottom = 500
    End With
TempY = 118
msurfBack.BltFast PointerX, PointerY, TalkBox, SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
msurfBack.DrawText 210, 384, UserCash, False
If TradeNPC = 2 Then
    If SalePointer >= 1 And SalePointer <= 10 Then
    If Facing = East Then TempCost = NPCz(DudeCoord.X + 1, DudeCoord.Y).Sell(SalePointer - 1).Cost
    If Facing = West Then TempCost = NPCz(DudeCoord.X - 1, DudeCoord.Y).Sell(SalePointer - 1).Cost
    If Facing = South Then TempCost = NPCz(DudeCoord.X, DudeCoord.Y + 1).Sell(SalePointer - 1).Cost
    If Facing = North Then TempCost = NPCz(DudeCoord.X, DudeCoord.Y - 1).Sell(SalePointer - 1).Cost
    msurfBack.DrawText 405, 384, TempCost, False
    End If
For IntCount = 0 To 9
If Facing = East Then TempNum = NPCz(DudeCoord.X + 1, DudeCoord.Y).Sell(IntCount).Index
If Facing = West Then TempNum = NPCz(DudeCoord.X - 1, DudeCoord.Y).Sell(IntCount).Index
If Facing = South Then TempNum = NPCz(DudeCoord.X, DudeCoord.Y + 1).Sell(IntCount).Index
If Facing = North Then TempNum = NPCz(DudeCoord.X, DudeCoord.Y - 1).Sell(IntCount).Index
If TempNum < 0 Then Exit Sub
TempName = DaItems(TempNum).Name
msurfBack.DrawText 170, TempY, "Buy: " & TempName, False
TempY = TempY + 25
Next IntCount
End If
If TradeNPC = 3 Then
If PointerX = 135 Then
    If UserInvent(SalePointer - 1).Index >= 0 Then TempCost = DaItems(UserInvent(SalePointer - 1).Index).Value
    msurfBack.DrawText 405, 384, TempCost, False
End If
For IntCount = SaleLoop To SaleLoop + 9
TempNum = UserInvent(IntCount).Index
TempAmount = UserInvent(IntCount).Amount
If TempNum < 0 Then GoTo SkipIt:
TempName = DaItems(TempNum).Name
msurfBack.DrawText 170, TempY, "Sell: " & TempName & ", Amount: " & TempAmount, False
SkipIt:
TempY = TempY + 25
Next IntCount
End If
End Sub

Private Sub DrawInventMenu()
Dim TempY As Integer
Dim SrcRect As RECT
Dim IntCount As Byte
Dim TempNum As Integer
Dim TempName As String
Dim TempAmount As Integer
Dim TempIndex As Integer
    With SrcRect
    .Left = 0
    .Right = 622
    .Top = 540
    .Bottom = 885
    End With
msurfBack.BltFast 10, 50, TalkBox, SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    With SrcRect
    .Left = 392
    .Right = 419
    .Top = 472
    .Bottom = 500
    End With
msurfBack.BltFast PointerX, PointerY, TalkBox, SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
TempY = 99
For IntCount = 0 To 4
    Select Case IntCount
    Case 0: TempIndex = UserWear.HelmetIndex
    Case 1: TempIndex = UserWear.WeapIndex
    Case 2: TempIndex = UserWear.ArmorIndex
    Case 3: TempIndex = UserWear.ShieldIndex
    Case 4: TempIndex = UserWear.RingIndex
    End Select
If TempIndex > -1 Then
    With SrcRect
    .Left = DaItems(TempIndex).GrhIndex * 32
    .Right = DaItems(TempIndex).GrhIndex * 32 + 32
    .Top = 0
    .Bottom = 32
    End With
msurfBack.BltFast 427, TempY, Itemz, SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If
TempY = TempY + 32
Next
TempY = 118
For IntCount = SaleLoop To SaleLoop + 9
TempNum = UserInvent(IntCount).Index
TempAmount = UserInvent(IntCount).Amount
If TempNum < 0 Then GoTo SkipIt:
TempName = DaItems(TempNum).Name
msurfBack.DrawText 60, TempY, TempName & ", Amount: " & TempAmount, False
SkipIt:
TempY = TempY + 25
Next IntCount
If UserWear.WeapIndex > -1 Then
msurfBack.DrawText 475, 262, DaItems(UserWear.WeapIndex).Name, False
Else: msurfBack.DrawText 475, 262, "None", False
End If
If UserWear.ArmorIndex > -1 Then
msurfBack.DrawText 465, 289, DaItems(UserWear.ArmorIndex).Name, False
Else: msurfBack.DrawText 465, 289, "None", False
End If
If UserWear.ShieldIndex > -1 Then
msurfBack.DrawText 454, 315, DaItems(UserWear.ShieldIndex).Name, False
Else: msurfBack.DrawText 454, 315, "None", False
End If
If UserWear.HelmetIndex > -1 Then
msurfBack.DrawText 455, 341, DaItems(UserWear.HelmetIndex).Name, False
Else: msurfBack.DrawText 455, 341, "None", False
End If
If UserWear.RingIndex > -1 Then
msurfBack.DrawText 441, 367, DaItems(UserWear.RingIndex).Name, False
Else: msurfBack.DrawText 441, 367, "None", False
End If
msurfBack.DrawText 80, 370, UserCash, False
End Sub

Private Function GetItem(intTileX As Integer, intTileY As Integer) As Byte
    GetItem = mbytMap((intTileX + TILE_WIDTH \ 2 + mintX - SCREEN_WIDTH \ 2) \ TILE_WIDTH, (intTileY + TILE_HEIGHT \ 2 + mintY - SCREEN_HEIGHT \ 2) \ TILE_HEIGHT).GItem.Index
End Function

Private Sub GetItemRect(bytItemNumber As Byte, ByRef intTileX As Integer, ByRef intTileY As Integer, ByRef rectTile As RECT)

    'Calc rect
    With rectTile
        .Left = bytItemNumber * TILE_WIDTH
        .Right = bytItemNumber * TILE_WIDTH + 32
        .Top = 0
        .Bottom = TILE_HEIGHT
    
    'Clip rect
        
        'If this tile is off the left side of the screen...
        If intTileX < 0 Then
            .Left = .Left - intTileX
            intTileX = 0
        End If
        'If this tile is off the top of the screen...
        If intTileY < 0 Then
            .Top = .Top - intTileY
            intTileY = 0
        End If
        'If this tile is off the right side of the screen...
        If intTileX + TILE_WIDTH > SCREEN_WIDTH Then .Right = .Right + (SCREEN_WIDTH - (intTileX + TILE_WIDTH))
        'If this tile is off the bottom of the screen...
        If intTileY + TILE_HEIGHT > SCREEN_HEIGHT Then .Bottom = .Bottom + (SCREEN_HEIGHT - (intTileY + TILE_HEIGHT))
    End With
End Sub

Private Sub DrawHealth()
Dim SrcRect As RECT
Dim DestRect As RECT
    With SrcRect
    .Left = 393
    .Right = 501
    .Top = 162
    .Bottom = 194
    End With
    msurfBack.BltFast 90, 17, TalkBox, SrcRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    With SrcRect
    .Left = 392
    .Right = 493
    .Top = 194
    .Bottom = 214
    End With
    With DestRect
    .Left = 94
    .Right = ((UserAtts.HP / 100) * 100) + .Left
    .Top = 23
    .Bottom = 44
    End With
    msurfBack.Blt DestRect, TalkBox, SrcRect, DDBLT_KEYSRC Or DDBLT_WAIT
    If UserAtts.HP >= 100 Then
    msurfBack.DrawText 130, 25, UserAtts.HP, False
    End If
    If UserAtts.HP < 100 And UserAtts.HP > 0 Then
    msurfBack.DrawText 135, 25, UserAtts.HP, False
    End If
    If UserAtts.HP = 0 Then
    msurfBack.DrawText 125, 25, "Dead", False
    End If
End Sub
Private Sub DrawSky()
Dim i As Integer
Dim j As Integer
Dim rectTile As RECT
Dim intx As Integer
Dim inty As Integer
Dim bytTileNum As Tilez
Dim LockBol As Boolean
Dim TempXY As Tilez
    'Draw the tiles according to the map array
    For i = 0 To CInt(SCREEN_WIDTH / TILE_WIDTH)
        For j = 0 To CInt(SCREEN_HEIGHT / TILE_HEIGHT)
            intx = i * TILE_WIDTH - mintX Mod TILE_WIDTH
            inty = j * TILE_HEIGHT - mintY Mod TILE_HEIGHT
            TempXY = GetXY(intx, inty)
            bytTileNum = GetTileS(TempXY.X, TempXY.Y)
            LockBol = GetLock(TempXY.X, TempXY.Y)
            If bytTileNum.X <> 0 Or bytTileNum.Y <> 0 Then
            GetRect bytTileNum, intx, inty, rectTile
            msurfBack.BltFast intx, inty, msurfTiles, rectTile, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            End If
                If LockBol Then
                bytTileNum.X = 7
                bytTileNum.Y = 8
                intx = i * TILE_WIDTH - mintX Mod TILE_WIDTH
                inty = j * TILE_HEIGHT - mintY Mod TILE_HEIGHT
                GetRect bytTileNum, intx, inty, rectTile
                msurfBack.BltFast intx, inty, msurfTiles, rectTile, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                End If
            
        Next j
    Next i
End Sub

Private Function GetLock(X As Integer, Y As Integer) As Boolean
GetLock = False
If mbytMap(X, Y).Walkable = False And mbytMap(X, Y).NPC = False And mbytMap(X, Y).NPCText <> "" Then
GetLock = True
End If
End Function

Private Function GetXY(intTileX As Integer, intTileY As Integer) As Tilez
GetXY.X = (intTileX + TILE_WIDTH \ 2 + mintX - SCREEN_WIDTH \ 2) \ TILE_WIDTH
GetXY.Y = (intTileY + TILE_HEIGHT \ 2 + mintY - SCREEN_HEIGHT \ 2) \ TILE_HEIGHT
End Function
