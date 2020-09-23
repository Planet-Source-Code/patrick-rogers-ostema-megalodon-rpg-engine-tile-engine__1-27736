VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Tile Scrolling Demo"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Megalodon's RPG Engine (Based on Lucky's Simple Tile Engine)
'www.megalodoncode.homestead.com
'Program flow variables
Public mblnRunning As Boolean
Const ConstTick = 1000 / 61 'The number of ticks in a cycle
Dim StartingTick As Long 'What the tick count is when
'your program starts
Dim CurrentTick1 As Long 'What your accumulated ticks
'should be about
Dim CurrentTick2 As Long 'What your accumulated ticks are
Dim DrawnFrame As Boolean 'Tells the program to draw the
'next frame
Dim GetNextFrame As Boolean 'Tells the program to get the
'next cycle if needed

Private Sub Form_Load()
Dim FileName As String
Dim intFreeFile As Integer
Dim fileNum As Integer
Dim xCounter As Integer
Dim yCounter As Integer
On Error GoTo UhOh:
With CommonDialog1
    .Filter = "Map Files(*.map)|*.map"
    .flags = cdlOFNFileMustExist
    .flags = cdlOFNHideReadOnly
    .DialogTitle = "Load A Map:"
    .ShowOpen
End With
FileName = CommonDialog1.FileName
If FileName = "" Then End
fileNum = FreeFile
Open FileName For Binary As fileNum
For xCounter = 0 To 50
    For yCounter = 0 To 50
        Get fileNum, , mbytMap(xCounter, yCounter)
    Next
Next
Get fileNum, , Exits
Close fileNum
If LoadList() = 1 Then
'Start the pre loop!
    PreLoop
End If
UhOh:
Unload Me
End Sub

Private Sub MainLoop()


Do While mblnRunning
If ((Gdx.TickCount - StartingTick) >= CurrentTick1) Or (GetNextFrame = True) Then
        If DDraw.LostSurfaces Then DDraw.LoadSurfaces
        DInput.HandleKeys
CurrentTick1 = CurrentTick1 + ConstTick
CurrentTick2 = Gdx.TickCount - StartingTick
DrawnFrame = False
GetNextFrame = False
End If
If CurrentTick2 > CurrentTick1 Then
GetNextFrame = True
DrawnFrame = True
End If
If DrawnFrame = False Then
        DDraw.MoveScreen
        DDraw.DrawTiles
        Combat.NPCManager
        DDraw.DrawNPC
        NPC.BubbleManager
        DDraw.FPS
DrawnFrame = True
End If
DoEvents
Loop

    'Unload everything
    DInput.Terminate
    DDraw.Terminate
    ShowCursor 1
    Unload frmMain
End Sub

Private Function LoadList() As Byte
Dim strMapName As String
Dim intFreeFile As Integer
Dim intCounter As Byte
LoadList = 0
On Error GoTo ErrorTrap2:
With CommonDialog1
    .Filter = "Megalodon RPG Item Lists((*.mrl)|*.mrl"
    .flags = cdlOFNFileMustExist
    .flags = cdlOFNHideReadOnly
    .DialogTitle = "Load Megalodon RPG Item List"
    .ShowOpen
End With

strMapName = CommonDialog1.FileName

If strMapName = "" Then GoTo ErrorTrap2:

intFreeFile = FreeFile
LoadList = 1
Open strMapName For Binary As intFreeFile
    For intCounter = 0 To 255
        Get intFreeFile, , DaItems(intCounter)
    Next

Close intFreeFile
ErrorTrap2:
End Function

Private Sub PreLoop()
Dim X As Integer
Dim Y As Integer
Dim strData(16) As String
Dim strTemp As String
Dim IntCount As Byte
Dim strSell As String
Dim TempNum As Integer
Dim TempDir As Byte
Dim TempMobile As Boolean
    For X = 0 To 50
        For Y = 0 To 50
        TempMobile = False
        TempDir = 0
        mbytMap(X, Y).Sprite = 0
        NPCz(X, Y).Duty = 0
        If mbytMap(X, Y).NPC Then
        strTemp = mbytMap(X, Y).NPCData
        For IntCount = 0 To 16
            If strTemp <> "" Then strData(IntCount) = Left(strTemp, InStr(1, strTemp, "|") - 1)
            If strTemp <> "" Then strTemp = Right(strTemp, Len(strTemp) - InStr(1, strTemp, "|"))
        Next
        mbytMap(X, Y).Sprite = Val(strData(0))
        TempDir = Val(strData(1)) * 3
        If Val(strData(2)) = 1 Then TempMobile = True
        If Val(strData(2)) = 0 Then TempMobile = False
        NPCz(X, Y).Duty = Val(strData(3))
            If NPCz(X, Y).Duty = 1 Then
            For IntCount = 0 To 9
            If strTemp <> "" Then strSell = Left(strTemp, InStr(1, strTemp, "|") - 1)
            TempNum = Val(strSell)
            NPCz(X, Y).Sell(IntCount).Index = TempNum
            If strTemp <> "" Then strTemp = Right(strTemp, Len(strTemp) - InStr(1, strTemp, "|"))
            If strTemp <> "" Then strSell = Left(strTemp, InStr(1, strTemp, "|") - 1)
            TempNum = Val(strSell)
            NPCz(X, Y).Sell(IntCount).Cost = TempNum
            If strTemp <> "" Then strTemp = Right(strTemp, Len(strTemp) - InStr(1, strTemp, "|"))
            Next
            End If
            If NPCz(X, Y).Duty = 2 Then
            NPCz(X, Y).State = Patrolling
            Else: NPCz(X, Y).State = Wandering
            End If
        NPCz(X, Y).Atts.HP = Val(strData(4))
        NPCz(X, Y).Atts.Strength = Val(strData(5))
        NPCz(X, Y).Atts.Armor = Val(strData(6))
        NPCz(X, Y).Atts.Speed = 60 - Val(strData(7))
        NPCz(X, Y).Atts.DSkill = Val(strData(8))
        NPCz(X, Y).Atts.ASkill = Val(strData(9))
        NPCz(X, Y).Atts.Sight = Val(strData(10))
        NPCz(X, Y).SpeedCounter = NPCz(X, Y).Atts.Speed
        For IntCount = 0 To 2
            NPCz(X, Y).Atts.Dropage(IntCount).Item = Val(strData(IntCount + 11 + IntCount)) + 1
            NPCz(X, Y).Atts.Dropage(IntCount).Amount = Val(strData(IntCount + 12 + IntCount))
        Next
        End If
        NPCz(X, Y).Mobile = TempMobile
        NPCz(X, Y).MoveX = 0
        NPCz(X, Y).MoveY = 0
        NPCz(X, Y).Frame = 0
        NPCz(X, Y).Step = 0
        NPCz(X, Y).StepCounter = 0
        NPCz(X, Y).CanMove = True
        NPCz(X, Y).LastStep = TempDir
        Next Y
    Next X
    For IntCount = 0 To 24
    UserInvent(IntCount).Index = -1
    UserInvent(IntCount).Amount = 0
    Next
    With DudeCoord
        .X = 25
        .Y = 25
    End With
    mbytMap(25, 25).NPC = True
    mbytMap(25, 25).Sprite = 0
    NPCz(25, 25).Mobile = True
    CanMove = True
    Walking = 0
    NPCFirst = True
    TradeNPC = 0
    SalePointer = 1
    UserCash = 0
    With UserWear
        .ArmorIndex = -1
        .HelmetIndex = -1
        .RingIndex = -1
        .ShieldIndex = -1
        .WeapIndex = -1
    End With
    With UserAtts
        .HP = 100
        .Speed = 60
        .Strength = 1
        .ASkill = 4
        .DSkill = 4
        .Armor = 0
    End With
        NPCz(25, 25).SpeedCounter = UserAtts.Speed
    Set Gdx = New DirectX7
    DInput.Initialize
    DDraw.Init
    'Set the initial player X,Y coords to the center
    mintX = 832
    mintY = 816
    'Set the Initial Main Character X,Y Coords
    Facing = South
    ShowCursor 0
'Start the loop running
mblnRunning = True
DrawnFrame = False
StartingTick = Gdx.TickCount
CurrentTick1 = 0
CurrentTick2 = 0
    MainLoop
End Sub

Public Sub LoadMap(MapName As String)
Dim FileName As String
Dim fileNum As Integer
Dim xCounter As Integer
Dim yCounter As Integer
On Error GoTo UhOh:
FileName = MapName & ".map"
fileNum = FreeFile
Open FileName For Binary As fileNum
For xCounter = 0 To 50
    For yCounter = 0 To 50
        Get fileNum, , mbytMap(xCounter, yCounter)
    Next
Next
Get fileNum, , Exits
Close fileNum
MapSetup
UhOh:
Unload Me
End Sub

Private Sub MapSetup()
Dim X As Integer
Dim Y As Integer
Dim strData(16) As String
Dim strTemp As String
Dim IntCount As Byte
Dim strSell As String
Dim TempNum As Integer
Dim TempDir As Byte
Dim TempMobile As Boolean
    For X = 0 To 50
        For Y = 0 To 50
        TempMobile = False
        TempDir = 0
        mbytMap(X, Y).Sprite = 0
        NPCz(X, Y).Duty = 0
        If mbytMap(X, Y).NPC Then
        strTemp = mbytMap(X, Y).NPCData
        For IntCount = 0 To 16
            If strTemp <> "" Then strData(IntCount) = Left(strTemp, InStr(1, strTemp, "|") - 1)
            If strTemp <> "" Then strTemp = Right(strTemp, Len(strTemp) - InStr(1, strTemp, "|"))
        Next
        mbytMap(X, Y).Sprite = Val(strData(0))
        TempDir = Val(strData(1)) * 3
        If Val(strData(2)) = 1 Then TempMobile = True
        If Val(strData(2)) = 0 Then TempMobile = False
        NPCz(X, Y).Duty = Val(strData(3))
            If NPCz(X, Y).Duty = 1 Then
            For IntCount = 0 To 9
            If strTemp <> "" Then strSell = Left(strTemp, InStr(1, strTemp, "|") - 1)
            TempNum = Val(strSell)
            NPCz(X, Y).Sell(IntCount).Index = TempNum
            If strTemp <> "" Then strTemp = Right(strTemp, Len(strTemp) - InStr(1, strTemp, "|"))
            If strTemp <> "" Then strSell = Left(strTemp, InStr(1, strTemp, "|") - 1)
            TempNum = Val(strSell)
            NPCz(X, Y).Sell(IntCount).Cost = TempNum
            If strTemp <> "" Then strTemp = Right(strTemp, Len(strTemp) - InStr(1, strTemp, "|"))
            Next
            End If
            If NPCz(X, Y).Duty = 2 Then
            NPCz(X, Y).State = Patrolling
            Else: NPCz(X, Y).State = Wandering
            End If
        NPCz(X, Y).Atts.HP = Val(strData(4))
        NPCz(X, Y).Atts.Strength = Val(strData(5))
        NPCz(X, Y).Atts.Armor = Val(strData(6))
        NPCz(X, Y).Atts.Speed = 60 - Val(strData(7))
        NPCz(X, Y).Atts.DSkill = Val(strData(8))
        NPCz(X, Y).Atts.ASkill = Val(strData(9))
        NPCz(X, Y).Atts.Sight = Val(strData(10))
        NPCz(X, Y).SpeedCounter = NPCz(X, Y).Atts.Speed
        For IntCount = 0 To 2
            NPCz(X, Y).Atts.Dropage(IntCount).Item = Val(strData(IntCount + 11 + IntCount)) + 1
            NPCz(X, Y).Atts.Dropage(IntCount).Amount = Val(strData(IntCount + 12 + IntCount))
        Next
        End If
        NPCz(X, Y).Mobile = TempMobile
        NPCz(X, Y).MoveX = 0
        NPCz(X, Y).MoveY = 0
        NPCz(X, Y).Frame = 0
        NPCz(X, Y).Step = 0
        NPCz(X, Y).StepCounter = 0
        NPCz(X, Y).CanMove = True
        NPCz(X, Y).LastStep = TempDir
        Next Y
    Next X
    If Facing = West Then
    DudeCoord.X = 50
    mbytMap(DudeCoord.X, DudeCoord.Y).NPC = True
    mbytMap(DudeCoord.X, DudeCoord.Y).Sprite = 0
    NPCz(DudeCoord.X, DudeCoord.Y).Mobile = True
    NPCz(DudeCoord.X, DudeCoord.Y).Facing = West
    NPCz(DudeCoord.X, DudeCoord.Y).LastStep = 6
    mintX = 1312
    Facing = West
    MainLoop
    End If
    If Facing = East Then
    DudeCoord.X = 0
    mbytMap(DudeCoord.X, DudeCoord.Y).NPC = True
    mbytMap(DudeCoord.X, DudeCoord.Y).Sprite = 0
    NPCz(DudeCoord.X, DudeCoord.Y).Mobile = True
    NPCz(DudeCoord.X, DudeCoord.Y).Facing = East
    NPCz(DudeCoord.X, DudeCoord.Y).LastStep = 9
    mintX = 320
    Facing = East
    MainLoop
    End If
    If Facing = North Then
    DudeCoord.Y = 50
    mbytMap(DudeCoord.X, DudeCoord.Y).NPC = True
    mbytMap(DudeCoord.X, DudeCoord.Y).Sprite = 0
    NPCz(DudeCoord.X, DudeCoord.Y).Mobile = True
    NPCz(DudeCoord.X, DudeCoord.Y).Facing = North
    NPCz(DudeCoord.X, DudeCoord.Y).LastStep = 3
    mintY = 1360
    Facing = North
    MainLoop
    End If
    If Facing = South Then
    DudeCoord.Y = 0
    mbytMap(DudeCoord.X, DudeCoord.Y).NPC = True
    mbytMap(DudeCoord.X, DudeCoord.Y).Sprite = 0
    NPCz(DudeCoord.X, DudeCoord.Y).Mobile = True
    NPCz(DudeCoord.X, DudeCoord.Y).Facing = South
    NPCz(DudeCoord.X, DudeCoord.Y).LastStep = 0
    mintY = 240
    Facing = South
    MainLoop
    End If
End Sub
