Attribute VB_Name = "Globals"
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Global Gdx As DirectX7
Type Exitz
    North As String
    South As String
    East As String
    West As String
End Type
Type DeadDropz
    Item As Integer
    Amount As Integer
End Type
Type Tilez
    X As Integer
    Y As Integer
End Type
Type TileLayerz
    Ground As Tilez
    Floor As Tilez
    Sky As Tilez
End Type
Type DamBubbles
    Damage As Byte
    Yaxis As Integer
    Xaxis As Integer
    Sway As Integer
End Type
Type Attys
    HP As Integer
    Strength As Integer
    Armor As Integer
    Speed As Integer
    Sight As Integer
    DSkill As Integer
    ASkill As Integer
    Dropage(2) As DeadDropz
End Type
Type GroundItem
    Index As Byte
    Amount As Integer
End Type
Type ViewEquip
    WeapIndex As Integer
    HelmetIndex As Integer
    ArmorIndex As Integer
    RingIndex As Integer
    ShieldIndex As Integer
End Type
Type SalesList
    Cost As Integer
    Index As Integer
End Type
Type MegaItemList
    Name As String
    Type As Byte
    Modifier As Byte
    Gives As Byte
    Value As Integer
    GrhIndex As Byte
End Type
Type Inventory
    Index As Integer
    Amount As Integer
    Equipped As Boolean
    GrhIndex As Byte
End Type
Type CharCoord
    X As Integer
    Y As Integer
End Type
Type Map
    TileNumber As TileLayerz
    Walkable As Boolean
    NPC As Boolean
    NPCText As String
    NPCData As String
    Sprite As Byte
    GItem As GroundItem
End Type
Type NPC
    Mobile As Boolean
    MoveX As Integer
    MoveY As Integer
    Walking As Byte
    Frame As Byte
    Step As Byte
    StepCounter As Byte
    LastStep As Byte
    CanMove As Boolean
    Duty As Byte
    Sell(9) As SalesList
    Atts As Attys
    State As Byte
    Facing As Byte
    SpeedCounter As Byte
    Bubbz As DamBubbles
End Type
'Some constants
Global Const SCREEN_WIDTH = 640
Global Const SCREEN_HEIGHT = 480
Global Const SCREEN_BITDEPTH = 16
Global Const TILE_WIDTH = 32
Global Const TILE_HEIGHT = 32
Global Const SCROLL_SPEED = 2
Global Const North = 1
Global Const South = 2
Global Const West = 3
Global Const East = 4
Global Const Wandering = 1
Global Const Attacking = 2
Global Const Patrolling = 3
Global mbytMap(51, 50) As Map               'Our map array
'Main Character Stuff
Global Walking As Byte
Global DudeCoord As CharCoord
Global CanMove As Boolean
Global Facing As Byte
'NPC Stuff
Global NPC As CharCoord
Global NPCz(50, 50) As NPC
Global SayNPC As Boolean
Global TradeNPC As Byte
Global DispInventMenu As Boolean
Global NPCTalk As String
Global NPCNext As String
Global NPCFirst As Boolean
Global TempName As String
Global mblnRunning As Boolean                  'Is the main loop still running?
Global mintX As Integer                        '"Player" X coordinate
Global mintY As Integer                        '"Player" Y coordinate
Global XCharOffSet As Byte
Global YCharOffSet As Byte
Global DaItems(255) As MegaItemList
Global UserInvent(24) As Inventory
Global UserWear As ViewEquip
Global SalePointer As Byte
Global UserCash As Integer
Global UserAtts As Attys
Global Exits As Exitz
