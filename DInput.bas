Attribute VB_Name = "DInput"
'dX Variables
Dim di As DirectInput
Dim diDEV As DirectInputDevice
Dim diState As DIKEYBOARDSTATE
'Loop counter
Dim i As Integer
'Public array showing which keys are active
Public aKeys(211) As Boolean
'Action Flow variables
Dim StopTalk As Boolean
Dim StopTradeR As Boolean
Dim StopTradeL As Boolean
Dim StopTradeD As Boolean
Dim StopTradeU As Boolean
Dim StopP As Boolean
'Keycode constants
Global Const DIK_ESCAPE = 1
Global Const DIK_1 = 2
Global Const DIK_2 = 3
Global Const DIK_3 = 4
Global Const DIK_4 = 5
Global Const DIK_5 = 6
Global Const DIK_6 = 7
Global Const DIK_7 = 8
Global Const DIK_8 = 9
Global Const DIK_9 = 10
Global Const DIK_0 = 11
Global Const DIK_MINUS = 12
Global Const DIK_EQUALS = 13
Global Const DIK_BACKSPACE = 14
Global Const DIK_TAB = 15
Global Const DIK_Q = 16
Global Const DIK_W = 17
Global Const DIK_E = 18
Global Const DIK_R = 19
Global Const DIK_T = 20
Global Const DIK_Y = 21
Global Const DIK_U = 22
Global Const DIK_I = 23
Global Const DIK_O = 24
Global Const DIK_P = 25
Global Const DIK_LBRACKET = 26
Global Const DIK_RBRACKET = 27
Global Const DIK_RETURN = 28
Global Const DIK_LCONTROL = 29
Global Const DIK_A = 30
Global Const DIK_S = 31
Global Const DIK_D = 32
Global Const DIK_F = 33
Global Const DIK_G = 34
Global Const DIK_H = 35
Global Const DIK_J = 36
Global Const DIK_K = 37
Global Const DIK_L = 38
Global Const DIK_SEMICOLON = 39
Global Const DIK_APOSTROPHE = 40
Global Const DIK_GRAVE = 41
Global Const DIK_LSHIFT = 42
Global Const DIK_BACKSLASH = 43
Global Const DIK_Z = 44
Global Const DIK_X = 45
Global Const DIK_C = 46
Global Const DIK_V = 47
Global Const DIK_B = 48
Global Const DIK_N = 49
Global Const DIK_M = 50
Global Const DIK_COMMA = 51
Global Const DIK_PERIOD = 52
Global Const DIK_SLASH = 53
Global Const DIK_RSHIFT = 54
Global Const DIK_MULTIPLY = 55
Global Const DIK_LALT = 56
Global Const DIK_SPACE = 57
Global Const DIK_CAPSLOCK = 58
Global Const DIK_F1 = 59
Global Const DIK_F2 = 60
Global Const DIK_F3 = 61
Global Const DIK_F4 = 62
Global Const DIK_F5 = 63
Global Const DIK_F6 = 64
Global Const DIK_F7 = 65
Global Const DIK_F8 = 66
Global Const DIK_F9 = 67
Global Const DIK_F10 = 68
Global Const DIK_NUMLOCK = 69
Global Const DIK_SCROLL = 70
Global Const DIK_NUMPAD7 = 71
Global Const DIK_NUMPAD8 = 72
Global Const DIK_NUMPAD9 = 73
Global Const DIK_SUBTRACT = 74
Global Const DIK_NUMPAD4 = 75
Global Const DIK_NUMPAD5 = 76
Global Const DIK_NUMPAD6 = 77
Global Const DIK_ADD = 78
Global Const DIK_NUMPAD1 = 79
Global Const DIK_NUMPAD2 = 80
Global Const DIK_NUMPAD3 = 81
Global Const DIK_NUMPAD0 = 82
Global Const DIK_DECIMAL = 83
Global Const DIK_F11 = 87
Global Const DIK_F12 = 88
Global Const DIK_NUMPADENTER = 156
Global Const DIK_RCONTROL = 157
Global Const DIK_DIVIDE = 181
Global Const DIK_RALT = 184
Global Const DIK_HOME = 199
Global Const DIK_UP = 200
Global Const DIK_PAGEUP = 201
Global Const DIK_LEFT = 203
Global Const DIK_RIGHT = 205
Global Const DIK_END = 207
Global Const DIK_DOWN = 208
Global Const DIK_PAGEDOWN = 209
Global Const DIK_INSERT = 210
Global Const DIK_DELETE = 211

Public Sub Initialize()
    Set Gdx = New DirectX7
    'Create the direct input object
    Set di = Gdx.DirectInputCreate()
        
    'Aquire the keyboard as the device
    Set diDEV = di.CreateDevice("GUID_SysKeyboard")
    
    'Get input nonexclusively, only when in foreground mode
    diDEV.SetCommonDataFormat DIFORMAT_KEYBOARD
    diDEV.SetCooperativeLevel frmMain.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    diDEV.Acquire
    
End Sub

Public Sub CheckKeys()
    
    'Get the current state of the keyboard
    diDEV.GetDeviceStateKeyboard diState
    
    'Scan through all the keys to check which are depressed
    For i = 1 To 211
        If diState.Key(i) <> 0 Then
            aKeys(i) = True             'If the key is pressed, set the appropriate array index to true
        Else
            aKeys(i) = False            'If the key is not pressed, set the appropriate array index to false
        End If
    Next
    
End Sub

Public Sub Terminate()
    
    'Unaquire the keyboard when we quit
    diDEV.Unacquire
    
End Sub

Public Sub HandleKeys()
Dim TempIndex As Integer
Dim TempIndex2 As Integer
Dim TempCost As Integer
Dim IntCount As Byte
Dim TempAmount As Integer
    CheckKeys                                            'Get the current state of the keyboard
    If aKeys(DIK_ESCAPE) Then
    frmMain.mblnRunning = False             'Has the escape key been pressed?
    TradeNPC = 0
    SayNPC = False
    DispInventMenu = False
    End If
If UserAtts.HP > 0 Then
    If aKeys(DIK_D) Then
        If DispInventMenu And StopP And DDraw.PointerX = 25 Then
        StopP = False
        TempIndex = SalePointer - 1
        TempIndex2 = UserInvent(TempIndex).Index
        If UserInvent(TempIndex).Equipped Then Exit Sub
        If mbytMap(DudeCoord.X, DudeCoord.Y).GItem.Index > 0 Then Exit Sub
        If UserInvent(TempIndex).Amount = 0 Then Exit Sub
        UserInvent(TempIndex).Amount = UserInvent(TempIndex).Amount - 1
        If UserInvent(TempIndex).Amount = 0 Then UserInvent(TempIndex).Index = -1
        mbytMap(DudeCoord.X, DudeCoord.Y).GItem.Index = TempIndex2 + 1
        mbytMap(DudeCoord.X, DudeCoord.Y).GItem.Amount = 1
        End If
    Else: StopP = True
    End If
    If aKeys(DIK_RCONTROL) Then
        If DispInventMenu And StopTalk And DDraw.PointerX = 25 Then
        StopTalk = False
        TempIndex = SalePointer - 1
        TempIndex2 = UserInvent(TempIndex).Index
        If TempIndex2 = -1 Then Exit Sub
        If UserInvent(TempIndex).Equipped Then
            UserInvent(TempIndex).Equipped = False
            Select Case DaItems(TempIndex2).Type
            Case 1: UserWear.WeapIndex = -1
            Case 2:
                    UserWear.HelmetIndex = -1
                    UserAtts.Armor = UserAtts.Armor - DaItems(TempIndex2).Modifier
            Case 3:
                    UserWear.ShieldIndex = -1
                    UserAtts.Armor = UserAtts.Armor - DaItems(TempIndex2).Modifier
            Case 4:
                    UserWear.ArmorIndex = -1
                    UserAtts.Armor = UserAtts.Armor - DaItems(TempIndex2).Modifier
            Case 5:
                    UserWear.RingIndex = -1
                    If DaItems(TempIndex2).Gives = 0 Then UserAtts.Armor = UserAtts.Armor - DaItems(TempIndex2).Modifier
                    If DaItems(TempIndex2).Gives = 1 Then UserAtts.Strength = UserAtts.Strength - DaItems(TempIndex2).Modifier
            End Select
            Exit Sub
        End If
        
        If UserInvent(TempIndex).Amount = 0 Then UserInvent(TempIndex).Index = -1
        If DaItems(TempIndex2).Type >= 1 And DaItems(TempIndex2).Type <= 5 Then
            For IntCount = 0 To 24
                If UserInvent(IntCount).Index > -1 Then
                    If DaItems(UserInvent(IntCount).Index).Type = DaItems(TempIndex2).Type Then UserInvent(IntCount).Equipped = False
                End If
            Next
        End If
        Select Case DaItems(TempIndex2).Type
        Case 1:
        UserInvent(TempIndex).Equipped = True
        UserWear.WeapIndex = TempIndex2
        Case 2:
        UserInvent(TempIndex).Equipped = True
        UserWear.HelmetIndex = TempIndex2
        UserAtts.Armor = UserAtts.Armor + DaItems(TempIndex2).Modifier
        Case 3:
        UserInvent(TempIndex).Equipped = True
        UserWear.ShieldIndex = TempIndex2
        UserAtts.Armor = UserAtts.Armor + DaItems(TempIndex2).Modifier
        Case 4:
        UserInvent(TempIndex).Equipped = True
        UserWear.ArmorIndex = TempIndex2
        UserAtts.Armor = UserAtts.Armor + DaItems(TempIndex2).Modifier
        Case 5:
        UserInvent(TempIndex).Equipped = True
        UserWear.RingIndex = TempIndex2
            If DaItems(TempIndex2).Gives = 0 Then UserAtts.Armor = UserAtts.Armor + DaItems(TempIndex2).Modifier
            If DaItems(TempIndex2).Gives = 1 Then UserAtts.Strength = UserAtts.Strength + DaItems(TempIndex2).Modifier
        Case 6:
        UserInvent(TempIndex).Amount = UserInvent(TempIndex).Amount - 1
        If UserInvent(TempIndex).Amount = 0 Then UserInvent(TempIndex).Index = -1
        If (DaItems(TempIndex2).Modifier + UserAtts.HP) <= 100 Then
        UserAtts.HP = UserAtts.HP + DaItems(TempIndex2).Modifier
        Else: UserAtts.HP = 100
        End If
        Case 7:
        MainChar.KeyOP UserInvent(TempIndex).Index
        End Select
        End If
        If DispInventMenu And StopTalk And DDraw.PointerY = 70 Then
        DispInventMenu = False
        StopTalk = False
        CanMove = True
        End If
        If TradeNPC = 3 And StopTalk And SalePointer >= 1 And DDraw.PointerX = 135 Then
        StopTalk = False
        TempIndex = SalePointer - 1
        TempIndex2 = UserInvent(TempIndex).Index
        If TempIndex2 = -1 Then Exit Sub
        If UserInvent(TempIndex).Equipped Then Exit Sub
        UserInvent(TempIndex).Amount = UserInvent(TempIndex).Amount - 1
        If UserInvent(TempIndex).Amount = 0 Then UserInvent(TempIndex).Index = -1
        TempCost = DaItems(TempIndex2).Value
        UserCash = UserCash + TempCost
        End If
        If TradeNPC = 2 And StopTalk And SalePointer >= 1 And SalePointer <= 10 Then
        StopTalk = False
        If Facing = East Then
        TempIndex = NPCz(DudeCoord.X + 1, DudeCoord.Y).Sell(SalePointer - 1).Index
        TempCost = NPCz(DudeCoord.X + 1, DudeCoord.Y).Sell(SalePointer - 1).Cost
        End If
        If Facing = West Then
        TempIndex = NPCz(DudeCoord.X - 1, DudeCoord.Y).Sell(SalePointer - 1).Index
        TempCost = NPCz(DudeCoord.X - 1, DudeCoord.Y).Sell(SalePointer - 1).Cost
        End If
        If Facing = South Then
        TempIndex = NPCz(DudeCoord.X, DudeCoord.Y + 1).Sell(SalePointer - 1).Index
        TempCost = NPCz(DudeCoord.X, DudeCoord.Y + 1).Sell(SalePointer - 1).Cost
        End If
        If Facing = North Then
        TempIndex = NPCz(DudeCoord.X, DudeCoord.Y - 1).Sell(SalePointer - 1).Index
        TempCost = NPCz(DudeCoord.X, DudeCoord.Y - 1).Sell(SalePointer - 1).Cost
        End If
            If TempCost <= UserCash Then
            For IntCount = 0 To 24
                If UserInvent(IntCount).Index = TempIndex Or UserInvent(IntCount).Index = -1 Then
                UserInvent(IntCount).Index = TempIndex
                UserInvent(IntCount).GrhIndex = DaItems(TempIndex).GrhIndex
                UserInvent(IntCount).Amount = UserInvent(IntCount).Amount + 1
                UserCash = UserCash - TempCost
                Exit Sub
                End If
            Next
            End If
        End If
        If TradeNPC = 2 And StopTalk And DDraw.PointerY = 70 Then
        TradeNPC = 0
        SayNPC = False
        StopTalk = False
        CanMove = True
        If Facing = East Then NPCz(DudeCoord.X + 1, DudeCoord.Y).CanMove = True
        If Facing = West Then NPCz(DudeCoord.X - 1, DudeCoord.Y).CanMove = True
        If Facing = South Then NPCz(DudeCoord.X, DudeCoord.Y + 1).CanMove = True
        If Facing = North Then NPCz(DudeCoord.X, DudeCoord.Y - 1).CanMove = True
        End If
        If TradeNPC = 3 And StopTalk And DDraw.PointerY = 70 Then
        TradeNPC = 0
        SayNPC = False
        StopTalk = False
        CanMove = True
        If Facing = East Then NPCz(DudeCoord.X + 1, DudeCoord.Y).CanMove = True
        If Facing = West Then NPCz(DudeCoord.X - 1, DudeCoord.Y).CanMove = True
        If Facing = South Then NPCz(DudeCoord.X, DudeCoord.Y + 1).CanMove = True
        If Facing = North Then NPCz(DudeCoord.X, DudeCoord.Y - 1).CanMove = True
        End If
        If TradeNPC = 1 And StopTalk And DDraw.PointerX = 515 Then
        DDraw.PointerX = 275
        SayNPC = False
        StopTalk = False
        TradeNPC = 0
        CanMove = True
        If Facing = East Then NPCz(DudeCoord.X + 1, DudeCoord.Y).CanMove = True
        If Facing = West Then NPCz(DudeCoord.X - 1, DudeCoord.Y).CanMove = True
        If Facing = South Then NPCz(DudeCoord.X, DudeCoord.Y + 1).CanMove = True
        If Facing = North Then NPCz(DudeCoord.X, DudeCoord.Y - 1).CanMove = True
        End If
        If TradeNPC = 1 And StopTalk And DDraw.PointerX = 275 Then
        SayNPC = False
        StopTalk = False
        SalePointer = 1
        TradeNPC = 2
        DDraw.PointerX = 135
        DDraw.PointerY = 113
        End If
        If TradeNPC = 1 And StopTalk And DDraw.PointerX = 395 Then
        SayNPC = False
        StopTalk = False
        SalePointer = 1
        SaleLoop = 0
        TradeNPC = 3
        DDraw.PointerX = 135
        DDraw.PointerY = 113
        End If
        If SayNPC = False And StopTalk And Walking = 0 And TradeNPC = 0 And DispInventMenu = False Then
            StopTalk = False
            NPC.CheckNPC
        If DDraw.PointerY = 70 Then DDraw.PointerX = 275
        End If
        If SayNPC And StopTalk And TradeNPC = 0 Then
            StopTalk = False
            If NPCNext <> "" Then
            NPCTalk = NPCNext
            NPCFirst = False
            Else
            CanMove = True
            NPCFirst = True
            SayNPC = False
            Walking = 0
            If Facing = East Then NPCz(DudeCoord.X + 1, DudeCoord.Y).CanMove = True
            If Facing = West Then NPCz(DudeCoord.X - 1, DudeCoord.Y).CanMove = True
            If Facing = South Then NPCz(DudeCoord.X, DudeCoord.Y + 1).CanMove = True
            If Facing = North Then NPCz(DudeCoord.X, DudeCoord.Y - 1).CanMove = True
            Exit Sub
            End If
        End If
    Else: StopTalk = True
    End If

    If aKeys(DIK_RIGHT) Then
        If TradeNPC = 1 And StopTradeR And Walking = 0 Then
            StopTradeR = False
            If DDraw.PointerX = 515 Then DDraw.PointerX = 155
            DDraw.PointerX = DDraw.PointerX + 120
        End If
    Else: StopTradeR = True
    End If
    If aKeys(DIK_LEFT) Then
        If TradeNPC = 1 And StopTradeL And Walking = 0 Then
            StopTradeL = False
            If DDraw.PointerX = 275 Then DDraw.PointerX = 635
            DDraw.PointerX = DDraw.PointerX - 120
        End If
    Else: StopTradeL = True
    End If
    If aKeys(DIK_DOWN) Then
        If TradeNPC = 2 And StopTradeD And Walking = 0 Then
            StopTradeD = False
            If DDraw.PointerY = 70 Then
            DDraw.PointerY = 88
            SalePointer = 0
            End If
            If DDraw.PointerY = 338 Then
            DDraw.PointerX = 360
            DDraw.PointerY = 45
            Else:
            DDraw.PointerX = 135
            End If
            DDraw.PointerY = DDraw.PointerY + 25
            SalePointer = SalePointer + 1
        End If
        If TradeNPC = 3 And StopTradeD And Walking = 0 Then
            StopTradeD = False
            If DDraw.PointerY = 70 Then
            DDraw.PointerY = 88
            SalePointer = DDraw.SaleLoop
            End If
            If DDraw.PointerY = 338 Then
                If SaleLoop < 15 Then
                DDraw.SaleLoop = DDraw.SaleLoop + 1
                SalePointer = SalePointer + 1
                Exit Sub
                End If
            End If
            If DDraw.PointerY = 338 Then
            DDraw.PointerX = 360
            DDraw.PointerY = 45
            Else:
            DDraw.PointerX = 135
            End If
            DDraw.PointerY = DDraw.PointerY + 25
            SalePointer = SalePointer + 1
        End If
        If DispInventMenu = True And StopTradeD And Walking = 0 Then
        StopTradeD = False
            If DDraw.PointerY = 70 Then
            DDraw.PointerY = 88
            SalePointer = DDraw.SaleLoop
            End If
            If DDraw.PointerY = 338 Then
                If SaleLoop < 15 Then
                DDraw.SaleLoop = DDraw.SaleLoop + 1
                SalePointer = SalePointer + 1
                Exit Sub
                End If
            End If
            If DDraw.PointerY = 338 Then
            DDraw.PointerX = 245
            DDraw.PointerY = 45
            Else:
            DDraw.PointerX = 25
            End If
            DDraw.PointerY = DDraw.PointerY + 25
            SalePointer = SalePointer + 1
        End If
    Else: StopTradeD = True
    End If
    If aKeys(DIK_UP) Then
        If TradeNPC = 2 And StopTradeU And Walking = 0 Then
            StopTradeU = False
            If DDraw.PointerY = 70 Then
            DDraw.PointerY = 363
            SalePointer = 11
            End If
            If DDraw.PointerY = 113 Then
            DDraw.PointerX = 360
            DDraw.PointerY = 95
            Else:
            DDraw.PointerX = 135
            End If
            DDraw.PointerY = DDraw.PointerY - 25
            SalePointer = SalePointer - 1
        End If
        If TradeNPC = 3 And StopTradeU And Walking = 0 Then
            StopTradeU = False
            If DDraw.PointerY = 70 Then
            DDraw.PointerY = 363
            SalePointer = DDraw.SaleLoop + 11
            End If
            If DDraw.PointerY = 113 Then
                If DDraw.SaleLoop > 0 Then
                DDraw.SaleLoop = DDraw.SaleLoop - 1
                SalePointer = SalePointer - 1
                Exit Sub
                End If
            End If
            If DDraw.PointerY = 113 Then
            DDraw.PointerX = 360
            DDraw.PointerY = 95
            Else:
            DDraw.PointerX = 135
            End If
            DDraw.PointerY = DDraw.PointerY - 25
            SalePointer = SalePointer - 1
        End If
        If DispInventMenu = True And StopTradeU And Walking = 0 Then
            StopTradeU = False
            If DDraw.PointerY = 70 Then
            DDraw.PointerY = 363
            SalePointer = DDraw.SaleLoop + 11
            End If
            If DDraw.PointerY = 113 Then
                If DDraw.SaleLoop > 0 Then
                DDraw.SaleLoop = DDraw.SaleLoop - 1
                SalePointer = SalePointer - 1
                Exit Sub
                End If
            End If
            If DDraw.PointerY = 113 Then
            DDraw.PointerX = 245
            DDraw.PointerY = 95
            Else:
            DDraw.PointerX = 25
            End If
            DDraw.PointerY = DDraw.PointerY - 25
            SalePointer = SalePointer - 1
        End If
    Else: StopTradeU = True
    End If
    If aKeys(DIK_I) Then
        If TradeNPC = 0 And SayNPC = False Then
        SaleLoop = 0
        SalePointer = 1
        CanMove = False
        DispInventMenu = True
        DDraw.PointerX = 25
        DDraw.PointerY = 113
        End If
    End If
End If
If Walking = 0 And CanMove Then
    If aKeys(DIK_UP) Then
        Walking = North
        MainChar.CheckCol
    GoTo Skipper:
    End If
    If aKeys(DIK_DOWN) Then
        Walking = South
        MainChar.CheckCol
    GoTo Skipper:
    End If
    If aKeys(DIK_LEFT) Then
        Walking = West
        MainChar.CheckCol
    GoTo Skipper:
    End If
    If aKeys(DIK_RIGHT) Then
        Walking = East
        MainChar.CheckCol
    End If
    If aKeys(DIK_LCONTROL) Then
        Combat.Attack
    Else: NPCz(DudeCoord.X, DudeCoord.Y).State = Wandering
    End If
    If aKeys(DIK_P) Then
        If mbytMap(DudeCoord.X, DudeCoord.Y).GItem.Index > 0 Then
            TempIndex = mbytMap(DudeCoord.X, DudeCoord.Y).GItem.Index - 1
            TempAmount = mbytMap(DudeCoord.X, DudeCoord.Y).GItem.Amount
            If DaItems(TempIndex).Type = 0 Then
                UserCash = UserCash + (DaItems(TempIndex).Value * TempAmount)
                mbytMap(DudeCoord.X, DudeCoord.Y).GItem.Index = 0
                mbytMap(DudeCoord.X, DudeCoord.Y).GItem.Amount = 0
                Exit Sub
            End If
            For IntCount = 0 To 24
                If UserInvent(IntCount).Index = TempIndex Or UserInvent(IntCount).Index = -1 Then
                UserInvent(IntCount).Index = TempIndex
                UserInvent(IntCount).GrhIndex = DaItems(TempIndex).GrhIndex
                UserInvent(IntCount).Amount = UserInvent(IntCount).Amount + TempAmount
                mbytMap(DudeCoord.X, DudeCoord.Y).GItem.Index = 0
                mbytMap(DudeCoord.X, DudeCoord.Y).GItem.Amount = 0
                Exit Sub
                End If
            Next
        End If
    End If
End If
Skipper:

End Sub
