Attribute VB_Name = "NPC"
Public Sub CheckNPC()
SayNPC = False
TradeNPC = False
    If mbytMap(DudeCoord.X + 1, DudeCoord.Y).NPC And NPCz(DudeCoord.X + 1, DudeCoord.Y).Walking = 0 And Facing = East Then
    SayNPC = True
        If NPCz(DudeCoord.X + 1, DudeCoord.Y).Duty = 1 Then
        TradeNPC = 1
        End If
    CanMove = False
    NPCz(DudeCoord.X + 1, DudeCoord.Y).LastStep = 6
    NPCz(DudeCoord.X + 1, DudeCoord.Y).Facing = West
    NPCTalk = mbytMap(DudeCoord.X + 1, DudeCoord.Y).NPCText
    NPCz(DudeCoord.X + 1, DudeCoord.Y).CanMove = False
    End If
    If mbytMap(DudeCoord.X - 1, DudeCoord.Y).NPC And NPCz(DudeCoord.X - 1, DudeCoord.Y).Walking = 0 And Facing = West Then
    SayNPC = True
        If NPCz(DudeCoord.X - 1, DudeCoord.Y).Duty = 1 Then
        TradeNPC = 1
        End If
    CanMove = False
    NPCz(DudeCoord.X - 1, DudeCoord.Y).LastStep = 9
    NPCz(DudeCoord.X - 1, DudeCoord.Y).Facing = East
    NPCTalk = mbytMap(DudeCoord.X - 1, DudeCoord.Y).NPCText
    NPCz(DudeCoord.X - 1, DudeCoord.Y).CanMove = False
    End If
    If mbytMap(DudeCoord.X, DudeCoord.Y + 1).NPC And NPCz(DudeCoord.X, DudeCoord.Y + 1).Walking = 0 And Facing = South Then
    SayNPC = True
        If NPCz(DudeCoord.X, DudeCoord.Y + 1).Duty = 1 Then
        TradeNPC = 1
        End If
    CanMove = False
    NPCz(DudeCoord.X, DudeCoord.Y + 1).LastStep = 3
    NPCz(DudeCoord.X, DudeCoord.Y + 1).Facing = North
    NPCTalk = mbytMap(DudeCoord.X, DudeCoord.Y + 1).NPCText
    NPCz(DudeCoord.X, DudeCoord.Y + 1).CanMove = False
    End If
    If mbytMap(DudeCoord.X, DudeCoord.Y - 1).NPC And NPCz(DudeCoord.X, DudeCoord.Y - 1).Walking = 0 And Facing = North Then
    SayNPC = True
        If NPCz(DudeCoord.X, DudeCoord.Y - 1).Duty = 1 Then
        TradeNPC = 1
        End If
    CanMove = False
    NPCz(DudeCoord.X, DudeCoord.Y - 1).LastStep = 0
    NPCz(DudeCoord.X, DudeCoord.Y - 1).Facing = South
    NPCTalk = mbytMap(DudeCoord.X, DudeCoord.Y - 1).NPCText
    NPCz(DudeCoord.X, DudeCoord.Y - 1).CanMove = False
    End If
End Sub
Public Sub MoveNPC(X As Integer, Y As Integer)
Dim IntRandom As Integer
Dim TempDist As Byte
Randomize
'Move them badboys!
        If X = DudeCoord.X Then
            If Y = DudeCoord.Y Then
            GoTo Skipper:
            End If
        End If
        If NPCz(X, Y).Mobile = False Then GoTo Skipper:
        If NPCz(X, Y).State = Attacking Then Exit Sub
        If NPCz(X, Y).State = Patrolling Then
            If Combat.GetDist(X, Y, DudeCoord.X, DudeCoord.Y) <= NPCz(X, Y).Atts.Sight Then GoTo Skipper:
        End If
        If mbytMap(X, Y).NPC = True Then
        IntRandom = Int((800 * Rnd) + 1)
            If IntRandom = 1 Then
            If X = 50 Then Exit Sub
            MoveNPCProp X, Y, East
            End If
            If IntRandom = 2 Then
            If X = 0 Then Exit Sub
            MoveNPCProp X, Y, West
            End If
            If IntRandom = 3 Then
            If Y = 50 Then Exit Sub
            MoveNPCProp X, Y, South
            End If
            If IntRandom = 4 Then
            If Y = 0 Then Exit Sub
            MoveNPCProp X, Y, North
            End If
        End If
Skipper:
    If NPCz(X, Y).Walking <> 0 Then
            If NPCz(X, Y).Walking = East Then
            NPCz(X, Y).MoveX = NPCz(X, Y).MoveX + SCROLL_SPEED
                If NPCz(X, Y).MoveX >= 0 Then
                    If X = DudeCoord.X Then
                        If Y = DudeCoord.Y Then
                            Walking = 0
                            If DInput.aKeys(DIK_RIGHT) Then
                            Walking = East
                            MainChar.CheckCol
                            NPCz(DudeCoord.X, DudeCoord.Y).MoveX = -33
                            End If
                        End If
                    End If
                NPCz(X, Y).MoveX = 0
                NPCz(X, Y).Walking = 0
                GoTo Skipper2:
                End If
            End If
            If NPCz(X, Y).Walking = West Then
            NPCz(X, Y).MoveX = NPCz(X, Y).MoveX - SCROLL_SPEED
                If NPCz(X, Y).MoveX <= 0 Then
                    If X = DudeCoord.X Then
                        If Y = DudeCoord.Y Then
                            Walking = 0
                            If DInput.aKeys(DIK_LEFT) Then
                            Walking = West
                            MainChar.CheckCol
                            End If
                        End If
                    End If
                NPCz(X, Y).MoveX = 0
                NPCz(X, Y).Walking = 0
                GoTo Skipper2:
                End If
            End If
            If NPCz(X, Y).Walking = South Then
            NPCz(X, Y).MoveY = NPCz(X, Y).MoveY + SCROLL_SPEED
                If NPCz(X, Y).MoveY >= 0 Then
                     If X = DudeCoord.X Then
                        If Y = DudeCoord.Y Then
                        Walking = 0
                            If DInput.aKeys(DIK_DOWN) Then
                            Walking = South
                            MainChar.CheckCol
                            NPCz(DudeCoord.X, DudeCoord.Y).MoveY = -33
                            End If
                        End If
                    End If
                NPCz(X, Y).MoveY = 0
                NPCz(X, Y).Walking = 0
                GoTo Skipper2:
                End If
            End If
            If NPCz(X, Y).Walking = North Then
            NPCz(X, Y).MoveY = NPCz(X, Y).MoveY - SCROLL_SPEED
                If NPCz(X, Y).MoveY <= 0 Then
                    If X = DudeCoord.X Then
                        If Y = DudeCoord.Y Then
                        Walking = 0
                            If DInput.aKeys(DIK_UP) Then
                            Walking = North
                            MainChar.CheckCol
                            End If
                        End If
                    End If
                NPCz(X, Y).MoveY = 0
                NPCz(X, Y).Walking = 0
                GoTo Skipper2:
                End If
            End If
    
    End If
Skipper2:
End Sub

Public Sub SetNPCFrame(TempNPCX As Integer, TempNPCY As Integer)
If NPCz(TempNPCX, TempNPCY).Walking <> 0 Then
If NPCz(TempNPCX, TempNPCY).StepCounter = 0 Then
    If NPCz(TempNPCX, TempNPCY).Walking = South Then
        Select Case NPCz(TempNPCX, TempNPCY).Step
        Case 0:
                NPCz(TempNPCX, TempNPCY).Frame = 1
                NPCz(TempNPCX, TempNPCY).Step = 1
                
        Case 1:
                NPCz(TempNPCX, TempNPCY).Frame = 0
                NPCz(TempNPCX, TempNPCY).Step = 2
                
        Case 2:
                NPCz(TempNPCX, TempNPCY).Frame = 2
                NPCz(TempNPCX, TempNPCY).Step = 3
                
        Case 3:
                NPCz(TempNPCX, TempNPCY).Frame = 0
                NPCz(TempNPCX, TempNPCY).Step = 0
                
        End Select
    End If
    If NPCz(TempNPCX, TempNPCY).Walking = North Then
        Select Case NPCz(TempNPCX, TempNPCY).Step
        Case 0:
                NPCz(TempNPCX, TempNPCY).Frame = 4
                NPCz(TempNPCX, TempNPCY).Step = 1
        Case 1:
                NPCz(TempNPCX, TempNPCY).Frame = 3
                NPCz(TempNPCX, TempNPCY).Step = 2
        Case 2:
                NPCz(TempNPCX, TempNPCY).Frame = 5
                NPCz(TempNPCX, TempNPCY).Step = 3
        Case 3:
                NPCz(TempNPCX, TempNPCY).Frame = 3
                NPCz(TempNPCX, TempNPCY).Step = 0
        End Select
    End If
    If NPCz(TempNPCX, TempNPCY).Walking = West Then
        Select Case NPCz(TempNPCX, TempNPCY).Step
        Case 0:
                NPCz(TempNPCX, TempNPCY).Frame = 7
                NPCz(TempNPCX, TempNPCY).Step = 1
        Case 1:
                NPCz(TempNPCX, TempNPCY).Frame = 6
                NPCz(TempNPCX, TempNPCY).Step = 2
        Case 2:
                NPCz(TempNPCX, TempNPCY).Frame = 8
                NPCz(TempNPCX, TempNPCY).Step = 3
        Case 3:
                NPCz(TempNPCX, TempNPCY).Frame = 6
                NPCz(TempNPCX, TempNPCY).Step = 0
        End Select
    End If
    If NPCz(TempNPCX, TempNPCY).Walking = East Then
        Select Case NPCz(TempNPCX, TempNPCY).Step
        Case 0:
                NPCz(TempNPCX, TempNPCY).Frame = 10
                NPCz(TempNPCX, TempNPCY).Step = 1
        Case 1:
                NPCz(TempNPCX, TempNPCY).Frame = 9
                NPCz(TempNPCX, TempNPCY).Step = 2
        Case 2:
                NPCz(TempNPCX, TempNPCY).Frame = 11
                NPCz(TempNPCX, TempNPCY).Step = 3
        Case 3:
                NPCz(TempNPCX, TempNPCY).Frame = 9
                NPCz(TempNPCX, TempNPCY).Step = 0
        End Select
    End If
NPCz(TempNPCX, TempNPCY).StepCounter = 8 / SCROLL_SPEED
End If
NPCz(TempNPCX, TempNPCY).StepCounter = NPCz(TempNPCX, TempNPCY).StepCounter - 1
End If
If NPCz(TempNPCX, TempNPCY).Walking <> 0 Then Exit Sub
If NPCz(TempNPCX, TempNPCY).State = Attacking Then
If NPCz(TempNPCX, TempNPCY).StepCounter = 0 Then
    If NPCz(TempNPCX, TempNPCY).Facing = South Then
        Select Case NPCz(TempNPCX, TempNPCY).Step
        Case 0:
                NPCz(TempNPCX, TempNPCY).Frame = 12
                NPCz(TempNPCX, TempNPCY).Step = 1
                
        Case 1:
                NPCz(TempNPCX, TempNPCY).Frame = 13
                NPCz(TempNPCX, TempNPCY).Step = 2
                
        Case 2:
                NPCz(TempNPCX, TempNPCY).Frame = 14
                NPCz(TempNPCX, TempNPCY).Step = 3
                
        Case 3:
                NPCz(TempNPCX, TempNPCY).Frame = 13
                NPCz(TempNPCX, TempNPCY).Step = 4
        
        Case 4:
                NPCz(TempNPCX, TempNPCY).Frame = 12
                NPCz(TempNPCX, TempNPCY).Step = 0
        End Select
    End If
    If NPCz(TempNPCX, TempNPCY).Facing = North Then
    Select Case NPCz(TempNPCX, TempNPCY).Step
        Case 0:
                NPCz(TempNPCX, TempNPCY).Frame = 15
                NPCz(TempNPCX, TempNPCY).Step = 1
                
        Case 1:
                NPCz(TempNPCX, TempNPCY).Frame = 16
                NPCz(TempNPCX, TempNPCY).Step = 2
                
        Case 2:
                NPCz(TempNPCX, TempNPCY).Frame = 17
                NPCz(TempNPCX, TempNPCY).Step = 3
                
        Case 3:
                NPCz(TempNPCX, TempNPCY).Frame = 16
                NPCz(TempNPCX, TempNPCY).Step = 4
        
        Case 4:
                NPCz(TempNPCX, TempNPCY).Frame = 15
                NPCz(TempNPCX, TempNPCY).Step = 0
        End Select
    End If
    If NPCz(TempNPCX, TempNPCY).Facing = West Then
    Select Case NPCz(TempNPCX, TempNPCY).Step
        Case 0:
                NPCz(TempNPCX, TempNPCY).Frame = 18
                NPCz(TempNPCX, TempNPCY).Step = 1
                
        Case 1:
                NPCz(TempNPCX, TempNPCY).Frame = 19
                NPCz(TempNPCX, TempNPCY).Step = 2
                
        Case 2:
                NPCz(TempNPCX, TempNPCY).Frame = 20
                NPCz(TempNPCX, TempNPCY).Step = 3
                
        Case 3:
                NPCz(TempNPCX, TempNPCY).Frame = 19
                NPCz(TempNPCX, TempNPCY).Step = 4
        
        Case 4:
                NPCz(TempNPCX, TempNPCY).Frame = 18
                NPCz(TempNPCX, TempNPCY).Step = 0
        End Select
    End If
    If NPCz(TempNPCX, TempNPCY).Facing = East Then
    Select Case NPCz(TempNPCX, TempNPCY).Step
        Case 0:
                NPCz(TempNPCX, TempNPCY).Frame = 21
                NPCz(TempNPCX, TempNPCY).Step = 1
                
        Case 1:
                NPCz(TempNPCX, TempNPCY).Frame = 22
                NPCz(TempNPCX, TempNPCY).Step = 2
                
        Case 2:
                NPCz(TempNPCX, TempNPCY).Frame = 23
                NPCz(TempNPCX, TempNPCY).Step = 3
                
        Case 3:
                NPCz(TempNPCX, TempNPCY).Frame = 22
                NPCz(TempNPCX, TempNPCY).Step = 4
        
        Case 4:
                NPCz(TempNPCX, TempNPCY).Frame = 21
                NPCz(TempNPCX, TempNPCY).Step = 0
        End Select
    End If
NPCz(TempNPCX, TempNPCY).StepCounter = 16 / SCROLL_SPEED
End If
NPCz(TempNPCX, TempNPCY).StepCounter = NPCz(TempNPCX, TempNPCY).StepCounter - 1
End If
End Sub



Public Sub MoveNPCProp(X As Integer, Y As Integer, Direction As Byte)
Dim IntCount As Byte
Select Case Direction
    Case East:
            If mbytMap(X + 1, Y).Walkable = False Then GoTo Skipper:
            If NPCz(X, Y).Walking <> 0 Then GoTo Skipper:
            If NPCz(X, Y).CanMove = False Then GoTo Skipper:
            mbytMap(X + 1, Y).NPCText = mbytMap(X, Y).NPCText
            mbytMap(X + 1, Y).Sprite = mbytMap(X, Y).Sprite
            NPCz(X + 1, Y).Duty = NPCz(X, Y).Duty
            NPCz(X + 1, Y).Mobile = NPCz(X, Y).Mobile
            NPCz(X + 1, Y).State = NPCz(X, Y).State

            NPCz(X + 1, Y).Bubbz = NPCz(X, Y).Bubbz
            NPCz(X, Y).Bubbz.Damage = 0
            NPCz(X + 1, Y).Walking = East
            NPCz(X + 1, Y).Step = 0
            NPCz(X + 1, Y).Facing = East
            NPCz(X + 1, Y).MoveX = -32
            NPCz(X + 1, Y).StepCounter = 0
            NPCz(X + 1, Y).Atts = NPCz(X, Y).Atts
            mbytMap(X, Y).NPC = False
            mbytMap(X + 1, Y).Walkable = False
            mbytMap(X + 1, Y).NPC = True
            mbytMap(X, Y).Walkable = True
            NPCz(X + 1, Y).LastStep = 9
                If NPCz(X, Y).Duty = 1 Then
                For IntCount = 0 To 9
                NPCz(X + 1, Y).Sell(IntCount) = NPCz(X, Y).Sell(IntCount)
                Next
                End If
    Case West:
            If mbytMap(X - 1, Y).Walkable = False Then GoTo Skipper:
            If NPCz(X, Y).Walking <> 0 Then GoTo Skipper:
            If NPCz(X, Y).CanMove = False Then GoTo Skipper:
            mbytMap(X - 1, Y).NPCText = mbytMap(X, Y).NPCText
            mbytMap(X - 1, Y).Sprite = mbytMap(X, Y).Sprite
            NPCz(X - 1, Y).Duty = NPCz(X, Y).Duty
            NPCz(X - 1, Y).Mobile = NPCz(X, Y).Mobile
            NPCz(X - 1, Y).State = NPCz(X, Y).State

            NPCz(X - 1, Y).Bubbz = NPCz(X, Y).Bubbz
            NPCz(X, Y).Bubbz.Damage = 0
            NPCz(X - 1, Y).Walking = West
            NPCz(X - 1, Y).Step = 0
            NPCz(X - 1, Y).Facing = West
            NPCz(X - 1, Y).MoveX = 32
            NPCz(X - 1, Y).StepCounter = 0
            NPCz(X - 1, Y).Atts = NPCz(X, Y).Atts
            mbytMap(X, Y).NPC = False
            mbytMap(X - 1, Y).NPC = True
            mbytMap(X - 1, Y).Walkable = False
            mbytMap(X, Y).Walkable = True
            NPCz(X - 1, Y).LastStep = 6
                If NPCz(X, Y).Duty = 1 Then
                For IntCount = 0 To 9
                NPCz(X - 1, Y).Sell(IntCount) = NPCz(X, Y).Sell(IntCount)
                Next
                End If
    Case South:
            If mbytMap(X, Y + 1).Walkable = False Then GoTo Skipper:
            If NPCz(X, Y).Walking <> 0 Then GoTo Skipper:
            If NPCz(X, Y).CanMove = False Then GoTo Skipper:
            mbytMap(X, Y + 1).NPCText = mbytMap(X, Y).NPCText
            mbytMap(X, Y + 1).Sprite = mbytMap(X, Y).Sprite
            NPCz(X, Y + 1).Duty = NPCz(X, Y).Duty
            NPCz(X, Y + 1).Mobile = NPCz(X, Y).Mobile
            NPCz(X, Y + 1).State = NPCz(X, Y).State

            NPCz(X, Y + 1).Bubbz = NPCz(X, Y).Bubbz
            NPCz(X, Y).Bubbz.Damage = 0
            NPCz(X, Y + 1).Walking = South
            NPCz(X, Y + 1).Step = 0
            NPCz(X, Y + 1).Facing = South
            NPCz(X, Y + 1).MoveY = -32
            NPCz(X, Y + 1).StepCounter = 0
            NPCz(X, Y + 1).Atts = NPCz(X, Y).Atts
            mbytMap(X, Y).NPC = False
            mbytMap(X, Y + 1).Walkable = False
            mbytMap(X, Y + 1).NPC = True
            mbytMap(X, Y).Walkable = True
            NPCz(X, Y + 1).LastStep = 0
                If NPCz(X, Y).Duty = 1 Then
                For IntCount = 0 To 9
                NPCz(X, Y + 1).Sell(IntCount) = NPCz(X, Y).Sell(IntCount)
                Next
                End If
    Case North:
            If mbytMap(X, Y - 1).Walkable = False Then GoTo Skipper:
            If NPCz(X, Y).Walking <> 0 Then GoTo Skipper:
            If NPCz(X, Y).CanMove = False Then GoTo Skipper:
            mbytMap(X, Y - 1).NPCText = mbytMap(X, Y).NPCText
            mbytMap(X, Y - 1).Sprite = mbytMap(X, Y).Sprite
            NPCz(X, Y - 1).Duty = NPCz(X, Y).Duty
            NPCz(X, Y - 1).Mobile = NPCz(X, Y).Mobile
            NPCz(X, Y - 1).State = NPCz(X, Y).State

            NPCz(X, Y - 1).Bubbz = NPCz(X, Y).Bubbz
            NPCz(X, Y).Bubbz.Damage = 0
            NPCz(X, Y - 1).Walking = North
            NPCz(X, Y - 1).Step = 0
            NPCz(X, Y - 1).Facing = North
            NPCz(X, Y - 1).MoveY = 32
            NPCz(X, Y - 1).StepCounter = 0
            NPCz(X, Y - 1).Atts = NPCz(X, Y).Atts
            mbytMap(X, Y).NPC = False
            mbytMap(X, Y - 1).Walkable = False
            mbytMap(X, Y - 1).NPC = True
            mbytMap(X, Y).Walkable = True
            NPCz(X, Y - 1).LastStep = 3
                If NPCz(X, Y).Duty = 1 Then
                For IntCount = 0 To 9
                NPCz(X, Y - 1).Sell(IntCount) = NPCz(X, Y).Sell(IntCount)
                Next
                End If
End Select
Skipper:
End Sub

Public Sub BubbleManager()
Dim intCounterX As Byte
Dim intCounterY As Byte

    For intCounterX = 0 To 50
        For intCounterY = 0 To 50

            If NPCz(intCounterX, intCounterY).Bubbz.Damage > 0 Then
                NPCz(intCounterX, intCounterY).Bubbz.Yaxis = NPCz(intCounterX, intCounterY).Bubbz.Yaxis - 1
                Select Case NPCz(intCounterX, intCounterY).Bubbz.Sway
                Case 0 To 8:
                            NPCz(intCounterX, intCounterY).Bubbz.Xaxis = NPCz(intCounterX, intCounterY).Bubbz.Xaxis - 1
                            NPCz(intCounterX, intCounterY).Bubbz.Sway = NPCz(intCounterX, intCounterY).Bubbz.Sway + 1
                Case 9 To 24:
                            NPCz(intCounterX, intCounterY).Bubbz.Xaxis = NPCz(intCounterX, intCounterY).Bubbz.Xaxis + 1
                            NPCz(intCounterX, intCounterY).Bubbz.Sway = NPCz(intCounterX, intCounterY).Bubbz.Sway + 1
                Case 25 To 33:
                            NPCz(intCounterX, intCounterY).Bubbz.Xaxis = NPCz(intCounterX, intCounterY).Bubbz.Xaxis - 1
                            NPCz(intCounterX, intCounterY).Bubbz.Sway = NPCz(intCounterX, intCounterY).Bubbz.Sway + 1
                End Select
                If NPCz(intCounterX, intCounterY).Bubbz.Yaxis = -33 Then
                    NPCz(intCounterX, intCounterY).Bubbz.Damage = 0
                    NPCz(intCounterX, intCounterY).Bubbz.Yaxis = 0
                    NPCz(intCounterX, intCounterY).Bubbz.Xaxis = 0
                    NPCz(intCounterX, intCounterY).Bubbz.Sway = 0
                End If
            End If

        Next
    Next
End Sub
