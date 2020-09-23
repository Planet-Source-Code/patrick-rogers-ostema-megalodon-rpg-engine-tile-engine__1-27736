Attribute VB_Name = "MainChar"


Public Sub CheckCol()
If Walking = North Then
Facing = North
If DudeCoord.Y = 0 Then
Walking = 0
frmMain.LoadMap Exits.North
End If
        If mbytMap(DudeCoord.X, DudeCoord.Y - 1).Walkable = False Then
        Walking = 0
        NPCz(DudeCoord.X, DudeCoord.Y).Walking = 0
        NPCz(DudeCoord.X, DudeCoord.Y).LastStep = 3
        NPCz(DudeCoord.X, DudeCoord.Y).Facing = North
        Else
        Walking = North
        NPCz(DudeCoord.X, DudeCoord.Y - 1).Walking = North
        NPCz(DudeCoord.X, DudeCoord.Y - 1).Step = 0
        NPCz(DudeCoord.X, DudeCoord.Y - 1).Atts = UserAtts
        NPCz(DudeCoord.X, DudeCoord.Y - 1).Facing = North
        NPCz(DudeCoord.X, DudeCoord.Y - 1).Mobile = True
        NPCz(DudeCoord.X, DudeCoord.Y - 1).State = Wandering
        NPCz(DudeCoord.X, DudeCoord.Y - 1).LastStep = 3
        NPCz(DudeCoord.X, DudeCoord.Y - 1).MoveY = 32
        NPCz(DudeCoord.X, DudeCoord.Y - 1).StepCounter = 0

        NPCz(DudeCoord.X, DudeCoord.Y - 1).Bubbz = NPCz(DudeCoord.X, DudeCoord.Y).Bubbz
        NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Damage = 0
        mbytMap(DudeCoord.X, DudeCoord.Y).NPC = False
        mbytMap(DudeCoord.X, DudeCoord.Y - 1).NPC = True
        mbytMap(DudeCoord.X, DudeCoord.Y - 1).Walkable = False
        mbytMap(DudeCoord.X, DudeCoord.Y).Walkable = True
        mbytMap(DudeCoord.X, DudeCoord.Y - 1).Sprite = mbytMap(DudeCoord.X, DudeCoord.Y).Sprite
        DudeCoord.Y = DudeCoord.Y - 1
        If DudeCoord.Y <= 6 Then YCharOffSet = YCharOffSet + 1
        If DudeCoord.Y > 41 Then YCharOffSet = YCharOffSet - 1
        End If
    GoTo Skipper:
End If
If Walking = South Then
Facing = South
If DudeCoord.Y = 50 Then
Walking = 0
frmMain.LoadMap Exits.South
End If
        If mbytMap(DudeCoord.X, DudeCoord.Y + 1).Walkable = False Then
        Walking = 0
        NPCz(DudeCoord.X, DudeCoord.Y).Walking = 0
        NPCz(DudeCoord.X, DudeCoord.Y).LastStep = 0
        NPCz(DudeCoord.X, DudeCoord.Y).Facing = South
        Else
        Walking = South
        NPCz(DudeCoord.X, DudeCoord.Y + 1).Walking = South
        NPCz(DudeCoord.X, DudeCoord.Y + 1).Step = 0
        NPCz(DudeCoord.X, DudeCoord.Y + 1).Atts = UserAtts
        NPCz(DudeCoord.X, DudeCoord.Y + 1).Facing = South
        NPCz(DudeCoord.X, DudeCoord.Y + 1).Mobile = True
        NPCz(DudeCoord.X, DudeCoord.Y + 1).State = Wandering
        NPCz(DudeCoord.X, DudeCoord.Y + 1).LastStep = 0
        NPCz(DudeCoord.X, DudeCoord.Y + 1).MoveY = -32
        NPCz(DudeCoord.X, DudeCoord.Y + 1).StepCounter = 0

        NPCz(DudeCoord.X, DudeCoord.Y + 1).Bubbz = NPCz(DudeCoord.X, DudeCoord.Y).Bubbz
        NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Damage = 0
        mbytMap(DudeCoord.X, DudeCoord.Y + 1).Walkable = False
        mbytMap(DudeCoord.X, DudeCoord.Y + 1).NPC = True
        mbytMap(DudeCoord.X, DudeCoord.Y).NPC = False
        mbytMap(DudeCoord.X, DudeCoord.Y).Walkable = True
        mbytMap(DudeCoord.X, DudeCoord.Y + 1).Sprite = mbytMap(DudeCoord.X, DudeCoord.Y).Sprite
        DudeCoord.Y = DudeCoord.Y + 1
        If DudeCoord.Y < 8 Then YCharOffSet = YCharOffSet - 1
        If DudeCoord.Y >= 43 Then YCharOffSet = YCharOffSet + 1
        End If
    GoTo Skipper:
End If
If Walking = West Then
Facing = West
If DudeCoord.X = 0 Then
Walking = 0
frmMain.LoadMap Exits.West
End If
        If mbytMap(DudeCoord.X - 1, DudeCoord.Y).Walkable = False Then
        Walking = 0
        NPCz(DudeCoord.X, DudeCoord.Y).Walking = 0
        NPCz(DudeCoord.X, DudeCoord.Y).LastStep = 6
        NPCz(DudeCoord.X, DudeCoord.Y).Facing = West
        Else
        Walking = West
        NPCz(DudeCoord.X - 1, DudeCoord.Y).Walking = West
        NPCz(DudeCoord.X - 1, DudeCoord.Y).Step = 0
        NPCz(DudeCoord.X - 1, DudeCoord.Y).Atts = UserAtts
        NPCz(DudeCoord.X - 1, DudeCoord.Y).Facing = West
        NPCz(DudeCoord.X - 1, DudeCoord.Y).Mobile = True
        NPCz(DudeCoord.X - 1, DudeCoord.Y).State = Wandering
        NPCz(DudeCoord.X - 1, DudeCoord.Y).LastStep = 6
        NPCz(DudeCoord.X - 1, DudeCoord.Y).MoveX = 32
        NPCz(DudeCoord.X - 1, DudeCoord.Y).StepCounter = 0

        NPCz(DudeCoord.X - 1, DudeCoord.Y).Bubbz = NPCz(DudeCoord.X, DudeCoord.Y).Bubbz
        NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Damage = 0
        mbytMap(DudeCoord.X - 1, DudeCoord.Y).Walkable = False
        mbytMap(DudeCoord.X, DudeCoord.Y).NPC = False
        mbytMap(DudeCoord.X - 1, DudeCoord.Y).NPC = True
        mbytMap(DudeCoord.X, DudeCoord.Y).Walkable = True
        mbytMap(DudeCoord.X - 1, DudeCoord.Y).Sprite = mbytMap(DudeCoord.X, DudeCoord.Y).Sprite
        DudeCoord.X = DudeCoord.X - 1
        If DudeCoord.X < 9 Then XCharOffSet = XCharOffSet + 1
        If DudeCoord.X >= 40 Then XCharOffSet = XCharOffSet - 1
        End If
    GoTo Skipper:
End If
If Walking = East Then
Facing = East
If DudeCoord.X = 50 Then
Walking = 0
frmMain.LoadMap Exits.East
End If
        If mbytMap(DudeCoord.X + 1, DudeCoord.Y).Walkable = False Then
        Walking = 0
        NPCz(DudeCoord.X, DudeCoord.Y).Walking = 0
        NPCz(DudeCoord.X, DudeCoord.Y).LastStep = 9
        NPCz(DudeCoord.X, DudeCoord.Y).Facing = East
        Else
        Walking = East
        NPCz(DudeCoord.X + 1, DudeCoord.Y).Walking = East
        NPCz(DudeCoord.X + 1, DudeCoord.Y).Step = 0
        NPCz(DudeCoord.X + 1, DudeCoord.Y).Atts = UserAtts
        NPCz(DudeCoord.X + 1, DudeCoord.Y).Facing = East
        NPCz(DudeCoord.X + 1, DudeCoord.Y).Mobile = True
        NPCz(DudeCoord.X + 1, DudeCoord.Y).State = Wandering
        NPCz(DudeCoord.X + 1, DudeCoord.Y).LastStep = 9
        NPCz(DudeCoord.X + 1, DudeCoord.Y).MoveX = -32
        NPCz(DudeCoord.X + 1, DudeCoord.Y).StepCounter = 0

        NPCz(DudeCoord.X + 1, DudeCoord.Y).Bubbz = NPCz(DudeCoord.X, DudeCoord.Y).Bubbz
        NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Damage = 0
        mbytMap(DudeCoord.X + 1, DudeCoord.Y).Walkable = False
        mbytMap(DudeCoord.X, DudeCoord.Y).NPC = False
        mbytMap(DudeCoord.X + 1, DudeCoord.Y).NPC = True
        mbytMap(DudeCoord.X, DudeCoord.Y).Walkable = True
        mbytMap(DudeCoord.X + 1, DudeCoord.Y).Sprite = mbytMap(DudeCoord.X, DudeCoord.Y).Sprite
        DudeCoord.X = DudeCoord.X + 1
        If DudeCoord.X <= 9 Then XCharOffSet = XCharOffSet - 1
        If DudeCoord.X > 40 Then XCharOffSet = XCharOffSet + 1
        End If
End If
Skipper:
End Sub

Public Sub KeyOP(KeyID As Integer)
Dim LockB As Integer
Dim LockS As String
Select Case Facing
Case North:
If mbytMap(DudeCoord.X, DudeCoord.Y - 1).NPCText <> "" Then LockS = mbytMap(DudeCoord.X, DudeCoord.Y - 1).NPCText
LockB = Val(LockS)
If KeyID = LockB Then mbytMap(DudeCoord.X, DudeCoord.Y - 1).Walkable = True
Case South:
If mbytMap(DudeCoord.X, DudeCoord.Y + 1).NPCText <> "" Then LockS = mbytMap(DudeCoord.X, DudeCoord.Y + 1).NPCText
LockB = Val(LockS)
If KeyID = LockB Then mbytMap(DudeCoord.X, DudeCoord.Y + 1).Walkable = True
Case East:
If mbytMap(DudeCoord.X + 1, DudeCoord.Y).NPCText <> "" Then LockS = mbytMap(DudeCoord.X + 1, DudeCoord.Y).NPCText
LockB = Val(LockS)
If KeyID = LockB Then mbytMap(DudeCoord.X + 1, DudeCoord.Y).Walkable = True
Case West:
If mbytMap(DudeCoord.X - 1, DudeCoord.Y).NPCText <> "" Then LockS = mbytMap(DudeCoord.X - 1, DudeCoord.Y).NPCText
LockB = Val(LockS)
If KeyID = LockB Then mbytMap(DudeCoord.X - 1, DudeCoord.Y).Walkable = True
End Select
End Sub
