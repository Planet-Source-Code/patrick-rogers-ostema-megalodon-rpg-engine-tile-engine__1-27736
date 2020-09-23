Attribute VB_Name = "Combat"

Public Sub Attack()
Dim TempHP As Integer
Dim TempAttack As Integer
Dim TempRnd As Integer
Dim TempRndA As Integer
Dim TempRndB As Integer
If NPCz(DudeCoord.X, DudeCoord.Y).SpeedCounter <> 0 Then NPCz(DudeCoord.X, DudeCoord.Y).SpeedCounter = NPCz(DudeCoord.X, DudeCoord.Y).SpeedCounter - 1
If NPCz(DudeCoord.X, DudeCoord.Y).SpeedCounter <> 0 Then GoTo SkipIt:
Randomize
If UserWear.WeapIndex = -1 Then Exit Sub
Select Case Facing
Case North:
If mbytMap(DudeCoord.X, DudeCoord.Y - 1).NPC = False Then Exit Sub
If NPCz(DudeCoord.X, DudeCoord.Y - 1).Duty <> 2 Then Exit Sub
If NPCz(DudeCoord.X, DudeCoord.Y - 1).Walking <> 0 Then Exit Sub
NPCz(DudeCoord.X, DudeCoord.Y).State = Attacking
NPCz(DudeCoord.X, DudeCoord.Y - 1).State = Attacking
NPCz(DudeCoord.X, DudeCoord.Y - 1).LastStep = 0
NPCz(DudeCoord.X, DudeCoord.Y - 1).Facing = South
If UserAtts.ASkill < NPCz(DudeCoord.X, DudeCoord.Y - 1).Atts.DSkill Then
    TempRnd = NPCz(DudeCoord.X, DudeCoord.Y - 1).Atts.DSkill - UserAtts.ASkill
    TempRndA = Int(((TempRnd + 1) * Rnd) + 0)
    Else
    TempRnd = TempRndA
End If
If TempRnd = TempRndA Then
TempRndB = Int(((UserAtts.ASkill + 1) * Rnd) + 0)
If TempRndB = UserAtts.ASkill Then Exit Sub
TempAttack = (Int(((UserAtts.Strength + 1) * Rnd) + 0) + DaItems(UserWear.WeapIndex).Modifier) - NPCz(DudeCoord.X, DudeCoord.Y - 1).Atts.Armor
If TempAttack < 1 Then Exit Sub
TempHP = NPCz(DudeCoord.X, DudeCoord.Y - 1).Atts.HP - TempAttack
    If TempHP > 1 Then
    NPCz(DudeCoord.X, DudeCoord.Y - 1).Atts.HP = TempHP
    NPCz(DudeCoord.X, DudeCoord.Y - 1).Bubbz.Yaxis = 0
    NPCz(DudeCoord.X, DudeCoord.Y - 1).Bubbz.Xaxis = 0
    NPCz(DudeCoord.X, DudeCoord.Y - 1).Bubbz.Sway = 0
    NPCz(DudeCoord.X, DudeCoord.Y - 1).Bubbz.Damage = TempAttack
    Else:
        NPCz(DudeCoord.X, DudeCoord.Y - 1).Atts.HP = 0
        NPCz(DudeCoord.X, DudeCoord.Y - 1).Mobile = False
        NPCz(DudeCoord.X, DudeCoord.Y - 1).Step = 0
        NPCz(DudeCoord.X, DudeCoord.Y - 1).Bubbz.Yaxis = 0
        NPCz(DudeCoord.X, DudeCoord.Y - 1).Bubbz.Xaxis = 0
        NPCz(DudeCoord.X, DudeCoord.Y - 1).Bubbz.Sway = 0
        NPCz(DudeCoord.X, DudeCoord.Y - 1).Bubbz.Damage = TempAttack
        mbytMap(DudeCoord.X, DudeCoord.Y - 1).NPC = False
        mbytMap(DudeCoord.X, DudeCoord.Y - 1).Walkable = True
        NPCDeath DudeCoord.X, DudeCoord.Y - 1
        End If
End If
Case South:
If NPCz(DudeCoord.X, DudeCoord.Y + 1).Walking <> 0 Then Exit Sub
If mbytMap(DudeCoord.X, DudeCoord.Y + 1).NPC = False Then Exit Sub
If NPCz(DudeCoord.X, DudeCoord.Y + 1).Duty <> 2 Then Exit Sub
NPCz(DudeCoord.X, DudeCoord.Y + 1).State = Attacking
NPCz(DudeCoord.X, DudeCoord.Y + 1).LastStep = 3
NPCz(DudeCoord.X, DudeCoord.Y + 1).Facing = North
NPCz(DudeCoord.X, DudeCoord.Y).State = Attacking
If UserAtts.ASkill < NPCz(DudeCoord.X, DudeCoord.Y + 1).Atts.DSkill Then
    TempRnd = NPCz(DudeCoord.X, DudeCoord.Y + 1).Atts.DSkill - UserAtts.ASkill
    TempRndA = Int(((TempRnd + 1) * Rnd) + 0)
    Else
    TempRnd = TempRndA
End If
If TempRnd = TempRndA Then
TempRndB = Int(((UserAtts.ASkill + 1) * Rnd) + 0)
If TempRndB = UserAtts.ASkill Then Exit Sub
TempAttack = (Int(((UserAtts.Strength + 1) * Rnd) + 0) + DaItems(UserWear.WeapIndex).Modifier) - NPCz(DudeCoord.X, DudeCoord.Y + 1).Atts.Armor
If TempAttack < 1 Then Exit Sub
TempHP = NPCz(DudeCoord.X, DudeCoord.Y + 1).Atts.HP - TempAttack
    If TempHP > 1 Then
    NPCz(DudeCoord.X, DudeCoord.Y + 1).Atts.HP = TempHP
    NPCz(DudeCoord.X, DudeCoord.Y + 1).Bubbz.Yaxis = 0
    NPCz(DudeCoord.X, DudeCoord.Y + 1).Bubbz.Xaxis = 0
    NPCz(DudeCoord.X, DudeCoord.Y + 1).Bubbz.Sway = 0
    NPCz(DudeCoord.X, DudeCoord.Y + 1).Bubbz.Damage = TempAttack
    Else:
        NPCz(DudeCoord.X, DudeCoord.Y + 1).Atts.HP = 0
        NPCz(DudeCoord.X, DudeCoord.Y + 1).Mobile = False
        NPCz(DudeCoord.X, DudeCoord.Y + 1).Step = 0
        NPCz(DudeCoord.X, DudeCoord.Y + 1).Bubbz.Yaxis = 0
        NPCz(DudeCoord.X, DudeCoord.Y + 1).Bubbz.Xaxis = 0
        NPCz(DudeCoord.X, DudeCoord.Y + 1).Bubbz.Sway = 0
        NPCz(DudeCoord.X, DudeCoord.Y + 1).Bubbz.Damage = TempAttack
        mbytMap(DudeCoord.X, DudeCoord.Y + 1).NPC = False
        mbytMap(DudeCoord.X, DudeCoord.Y + 1).Walkable = True
        NPCDeath DudeCoord.X, DudeCoord.Y + 1
    End If
End If
Case West:
If NPCz(DudeCoord.X - 1, DudeCoord.Y).Walking <> 0 Then Exit Sub
If mbytMap(DudeCoord.X - 1, DudeCoord.Y).NPC = False Then Exit Sub
If NPCz(DudeCoord.X - 1, DudeCoord.Y).Duty <> 2 Then Exit Sub
NPCz(DudeCoord.X - 1, DudeCoord.Y).State = Attacking
NPCz(DudeCoord.X - 1, DudeCoord.Y).LastStep = 9
NPCz(DudeCoord.X - 1, DudeCoord.Y).Facing = East
NPCz(DudeCoord.X, DudeCoord.Y).State = Attacking
If UserAtts.ASkill < NPCz(DudeCoord.X - 1, DudeCoord.Y).Atts.DSkill Then
    TempRnd = NPCz(DudeCoord.X - 1, DudeCoord.Y).Atts.DSkill - UserAtts.ASkill
    TempRndA = Int(((TempRnd + 1) * Rnd) + 0)
    Else
    TempRnd = TempRndA
End If
If TempRnd = TempRndA Then
TempRndB = Int(((UserAtts.ASkill + 1) * Rnd) + 0)
If TempRndB = UserAtts.ASkill Then Exit Sub
TempAttack = (Int(((UserAtts.Strength + 1) * Rnd) + 0) + DaItems(UserWear.WeapIndex).Modifier) - NPCz(DudeCoord.X - 1, DudeCoord.Y).Atts.Armor
If TempAttack < 1 Then Exit Sub
TempHP = NPCz(DudeCoord.X - 1, DudeCoord.Y).Atts.HP - TempAttack
    If TempHP > 1 Then
    NPCz(DudeCoord.X - 1, DudeCoord.Y).Atts.HP = TempHP
    NPCz(DudeCoord.X - 1, DudeCoord.Y).Bubbz.Yaxis = 0
    NPCz(DudeCoord.X - 1, DudeCoord.Y).Bubbz.Xaxis = 0
    NPCz(DudeCoord.X - 1, DudeCoord.Y).Bubbz.Sway = 0
    NPCz(DudeCoord.X - 1, DudeCoord.Y).Bubbz.Damage = TempAttack
    Else:
        NPCz(DudeCoord.X - 1, DudeCoord.Y).Atts.HP = 0
        NPCz(DudeCoord.X - 1, DudeCoord.Y).Mobile = False
        NPCz(DudeCoord.X - 1, DudeCoord.Y).Step = 0
        NPCz(DudeCoord.X - 1, DudeCoord.Y).Bubbz.Yaxis = 0
        NPCz(DudeCoord.X - 1, DudeCoord.Y).Bubbz.Xaxis = 0
        NPCz(DudeCoord.X - 1, DudeCoord.Y).Bubbz.Sway = 0
        NPCz(DudeCoord.X - 1, DudeCoord.Y).Bubbz.Damage = TempAttack
        mbytMap(DudeCoord.X - 1, DudeCoord.Y).NPC = False
        mbytMap(DudeCoord.X - 1, DudeCoord.Y).Walkable = True
        NPCDeath DudeCoord.X - 1, DudeCoord.Y
    End If
End If
Case East:
If NPCz(DudeCoord.X + 1, DudeCoord.Y).Walking <> 0 Then Exit Sub
If mbytMap(DudeCoord.X + 1, DudeCoord.Y).NPC = False Then Exit Sub
If NPCz(DudeCoord.X + 1, DudeCoord.Y).Duty <> 2 Then Exit Sub
NPCz(DudeCoord.X + 1, DudeCoord.Y).State = Attacking
NPCz(DudeCoord.X + 1, DudeCoord.Y).LastStep = 6
NPCz(DudeCoord.X + 1, DudeCoord.Y).Facing = West
NPCz(DudeCoord.X, DudeCoord.Y).State = Attacking
If UserAtts.ASkill < NPCz(DudeCoord.X + 1, DudeCoord.Y).Atts.DSkill Then
    TempRnd = NPCz(DudeCoord.X + 1, DudeCoord.Y).Atts.DSkill - UserAtts.ASkill
    TempRndA = Int(((TempRnd + 1) * Rnd) + 0)
    Else
    TempRnd = TempRndA
End If
If TempRnd = TempRndA Then
TempRndB = Int(((UserAtts.ASkill + 1) * Rnd) + 0)
If TempRndB = UserAtts.ASkill Then Exit Sub
TempAttack = (Int(((UserAtts.Strength + 1) * Rnd) + 0) + DaItems(UserWear.WeapIndex).Modifier) - NPCz(DudeCoord.X + 1, DudeCoord.Y).Atts.Armor
If TempAttack < 1 Then Exit Sub
TempHP = NPCz(DudeCoord.X + 1, DudeCoord.Y).Atts.HP - TempAttack
    If TempHP > 1 Then
    NPCz(DudeCoord.X + 1, DudeCoord.Y).Atts.HP = TempHP
    NPCz(DudeCoord.X + 1, DudeCoord.Y).Bubbz.Yaxis = 0
    NPCz(DudeCoord.X + 1, DudeCoord.Y).Bubbz.Xaxis = 0
    NPCz(DudeCoord.X + 1, DudeCoord.Y).Bubbz.Sway = 0
    NPCz(DudeCoord.X + 1, DudeCoord.Y).Bubbz.Damage = TempAttack
    Else:
        NPCz(DudeCoord.X + 1, DudeCoord.Y).Atts.HP = 0
        NPCz(DudeCoord.X + 1, DudeCoord.Y).Mobile = False
        NPCz(DudeCoord.X + 1, DudeCoord.Y).Step = 0
        NPCz(DudeCoord.X + 1, DudeCoord.Y).Atts.HP = TempHP
        NPCz(DudeCoord.X + 1, DudeCoord.Y).Bubbz.Yaxis = 0
        NPCz(DudeCoord.X + 1, DudeCoord.Y).Bubbz.Xaxis = 0
        NPCz(DudeCoord.X + 1, DudeCoord.Y).Bubbz.Sway = 0
        NPCz(DudeCoord.X + 1, DudeCoord.Y).Bubbz.Damage = TempAttack
        mbytMap(DudeCoord.X + 1, DudeCoord.Y).NPC = False
        mbytMap(DudeCoord.X + 1, DudeCoord.Y).Walkable = True
        NPCDeath DudeCoord.X + 1, DudeCoord.Y
    End If
End If
End Select
NPCz(DudeCoord.X, DudeCoord.Y).SpeedCounter = UserAtts.Speed
SkipIt:
End Sub

Public Sub NPCManager()
Dim X As Integer
Dim Y As Integer
For X = 0 To 50
    For Y = 0 To 50
    If mbytMap(X, Y).NPC Then
    Select Case NPCz(X, Y).State
    Case Wandering:
        MoveNPC X, Y
    Case Attacking:
        AIAttack X, Y
    Case Patrolling:
        If Combat.GetDist(X, Y, DudeCoord.X, DudeCoord.Y) > NPCz(X, Y).Atts.Sight Then
            NPC.MoveNPC X, Y
            Else
                If NPCz(X, Y).Walking = 0 Then
                    AISeek X, Y
                Else
                    NPC.MoveNPC X, Y
                End If
        End If
    End Select
    End If
    Next
Next
End Sub

Private Sub AIAttack(X As Integer, Y As Integer)
Dim TempHP As Integer
Dim TempAttack As Integer
Dim TempRnd As Integer
Dim TempRndA As Integer
Dim TempRndB As Integer
Dim TempBub As Byte
If NPCz(X, Y).CanMove = False Then Exit Sub
If UserAtts.HP = 0 Then
    NPCz(X, Y).State = Wandering
    NPCz(X, Y).StepCounter = 0
    NPCz(X, Y).Step = 0
End If
If (X <> DudeCoord.X) And (Y <> DudeCoord.Y) Or GetDist(X, Y, DudeCoord.X, DudeCoord.Y) > 1 Then
    NPCz(X, Y).State = Patrolling
    NPCz(X, Y).StepCounter = 0
    NPCz(X, Y).Step = 0
    Exit Sub
End If
        If NPCz(X, Y).SpeedCounter <> 0 Then NPCz(X, Y).SpeedCounter = NPCz(X, Y).SpeedCounter - 1
        If NPCz(X, Y).SpeedCounter <> 0 Then GoTo SkipIt:
TempBub = 0
Randomize
        Select Case NPCz(X, Y).Facing
        Case North:
        If DudeCoord.X <> X Then GoTo SkipIt:
        If DudeCoord.Y <> Y - 1 Then GoTo SkipIt:
            If UserAtts.ASkill < NPCz(X, Y).Atts.DSkill Then
                TempRnd = NPCz(X, Y).Atts.DSkill - UserAtts.ASkill
                TempRndA = Int(((TempRnd + 1) * Rnd) + 0)
                Else
                TempRnd = TempRndA
            End If
            If TempRnd = TempRndA Then
            TempRndB = Int(((NPCz(X, Y).Atts.ASkill + 1) * Rnd) + 0)
            If TempRndB = NPCz(X, Y).Atts.ASkill Then Exit Sub
            TempAttack = (Int((NPCz(X, Y).Atts.Strength + 1) * Rnd) + 0) - UserAtts.Armor
            If TempAttack < 1 Then GoTo SkipIt:
            TempHP = UserAtts.HP - TempAttack
                If TempHP > 1 Then
                    UserAtts.HP = TempHP
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Yaxis = 0
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Xaxis = 0
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Sway = 0
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Damage = TempAttack
                
                Else:
                    UserAtts.HP = 0
                    
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Damage = TempAttack
                    CanMove = False
                End If
            End If
        Case South:
        If DudeCoord.X <> X Then GoTo SkipIt:
        If DudeCoord.Y <> Y + 1 Then GoTo SkipIt:
            If UserAtts.ASkill < NPCz(X, Y).Atts.DSkill Then
                TempRnd = NPCz(X, Y).Atts.DSkill - UserAtts.ASkill
                TempRndA = Int(((TempRnd + 1) * Rnd) + 0)
                Else
                TempRnd = TempRndA
            End If
            If TempRnd = TempRndA Then
            TempRndB = Int(((NPCz(X, Y).Atts.ASkill + 1) * Rnd) + 0)
            If TempRndB = NPCz(X, Y).Atts.ASkill Then Exit Sub
            TempAttack = (Int((NPCz(X, Y).Atts.Strength + 1) * Rnd) + 0) - UserAtts.Armor
            If TempAttack < 1 Then GoTo SkipIt:
            TempHP = UserAtts.HP - TempAttack
                If TempHP > 1 Then
                    UserAtts.HP = TempHP
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Yaxis = 0
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Xaxis = 0
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Sway = 0
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Damage = TempAttack
                Else:
                    UserAtts.HP = 0
                    
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Damage = TempAttack
                    CanMove = False
                End If
            End If
        Case East:
        If DudeCoord.X <> X + 1 Then GoTo SkipIt:
        If DudeCoord.Y <> Y Then GoTo SkipIt:
            If UserAtts.ASkill < NPCz(X, Y).Atts.DSkill Then
                TempRnd = NPCz(X, Y).Atts.DSkill - UserAtts.ASkill
                TempRndA = Int(((TempRnd + 1) * Rnd) + 0)
                Else
                TempRnd = TempRndA
            End If
            If TempRnd = TempRndA Then
            TempRndB = Int(((NPCz(X, Y).Atts.ASkill + 1) * Rnd) + 0)
            If TempRndB = NPCz(X, Y).Atts.ASkill Then Exit Sub
            TempAttack = (Int((NPCz(X, Y).Atts.Strength + 1) * Rnd) + 0) - UserAtts.Armor
            If TempAttack < 1 Then GoTo SkipIt:
            TempHP = UserAtts.HP - TempAttack
                If TempHP > 1 Then
                    UserAtts.HP = TempHP
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Yaxis = 0
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Xaxis = 0
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Sway = 0
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Damage = TempAttack
                Else:
                    UserAtts.HP = 0
                    
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Damage = TempAttack
                    CanMove = False
                End If
            End If
        Case West:
        If DudeCoord.X <> X - 1 Then GoTo SkipIt:
        If DudeCoord.Y <> Y Then GoTo SkipIt:
            If UserAtts.ASkill < NPCz(X, Y).Atts.DSkill Then
                TempRnd = NPCz(X, Y).Atts.DSkill - UserAtts.ASkill
                TempRndA = Int(((TempRnd + 1) * Rnd) + 0)
                Else
                TempRnd = TempRndA
            End If
            If TempRnd = TempRndA Then
            TempRndB = Int(((NPCz(X, Y).Atts.ASkill + 1) * Rnd) + 0)
            If TempRndB = NPCz(X, Y).Atts.ASkill Then Exit Sub
            TempAttack = (Int((NPCz(X, Y).Atts.Strength + 1) * Rnd) + 0) - UserAtts.Armor
            If TempAttack < 1 Then GoTo SkipIt:
            TempHP = UserAtts.HP - TempAttack
                If TempHP > 1 Then
                    UserAtts.HP = TempHP
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Yaxis = 0
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Xaxis = 0
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Sway = 0
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Damage = TempAttack
                Else:
                    UserAtts.HP = 0
                    
                    NPCz(DudeCoord.X, DudeCoord.Y).Bubbz.Damage = TempAttack
                    CanMove = False
                End If
            End If
        End Select
            NPCz(X, Y).SpeedCounter = NPCz(X, Y).Atts.Speed
SkipIt:
End Sub

Private Sub AISeek(X As Integer, Y As Integer)
Dim TempDist As Byte
If NPCz(X, Y).Mobile = False Then Exit Sub
If NPCz(X, Y).Walking <> 0 Then Exit Sub
TempDist = GetDist(X, Y, DudeCoord.X, DudeCoord.Y)
If TempDist <= NPCz(X, Y).Atts.Sight Then
    If DudeCoord.X < X Then
    MoveNPCProp X, Y, West
    If NPCz(X - 1, Y).MoveX = 32 Then NPCz(X - 1, Y).MoveX = 30
    End If
    If DudeCoord.X > X Then
    MoveNPCProp X, Y, East
    End If
    If NPCz(X - 1, Y).Walking = 0 Then
        If NPCz(X + 1, Y).Walking = 0 Then
                If DudeCoord.Y > Y Then
                MoveNPCProp X, Y, South
                End If
                If DudeCoord.Y < Y Then
                MoveNPCProp X, Y, North
                If NPCz(X, Y - 1).MoveY = 32 Then NPCz(X, Y - 1).MoveY = 30
                End If
        End If
    End If
End If
If TempDist = 1 Then
    If DudeCoord.X > X Then
        NPCz(X, Y).LastStep = 9
        NPCz(X, Y).Facing = East
        NPCz(X, Y).State = Attacking
    End If
    If DudeCoord.X < X Then
        NPCz(X, Y).LastStep = 6
        NPCz(X, Y).Facing = West
        NPCz(X, Y).State = Attacking
    End If
    If DudeCoord.Y < Y Then
        NPCz(X, Y).LastStep = 3
        NPCz(X, Y).Facing = North
        NPCz(X, Y).State = Attacking
    End If
    If DudeCoord.Y > Y Then
        NPCz(X, Y).LastStep = 0
        NPCz(X, Y).Facing = South
        NPCz(X, Y).State = Attacking
    End If
End If
NPC.MoveNPC X, Y
End Sub
Public Function GetDist(intX1 As Integer, intY1 As Integer, intX2 As Integer, intY2 As Integer) As Byte
    GetDist = Sqr((intX1 - intX2) ^ 2 + (intY1 - intY2) ^ 2)
End Function

Private Sub NPCDeath(X As Integer, Y As Integer)
Dim RndNum As Byte
Randomize
RndNum = Int((3 * Rnd) + 0)
If NPCz(X, Y).Atts.Dropage(RndNum).Item <> -1 Then
mbytMap(X, Y).GItem.Index = NPCz(X, Y).Atts.Dropage(RndNum).Item
mbytMap(X, Y).GItem.Amount = NPCz(X, Y).Atts.Dropage(RndNum).Amount
End If
End Sub
