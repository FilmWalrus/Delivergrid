Module GlobalVars

#Region " Global Variables "
    '--Gameplay

    '--GameTypes

    '--Board
    Public Const GridWidth As Integer = 11 '-- 0 to 12
    Public Const GridHeight As Integer = 11 '-- 0 to 12
    Public TheBoxes(GridWidth, GridHeight) As Label
    Public BoxInfo(GridWidth, GridHeight) As TerrainSquare

    Public Const MaxDistance = 10000

    '--Legend
    Public TheLegend(8) As Label
    Public LegendInfo(8) As TerrainSquare

    '-Inventory
    Public Const InventoryWidth As Integer = 3 '-- 0 to 4
    Public Const InventoryHeight As Integer = 2 '-- 0 to 3
    Public TheInventory(InventoryWidth, InventoryHeight) As Label
    Public InventoryInfo(InventoryWidth, InventoryHeight) As TerrainSquare

    '--Terrain
    Public Const TerrainNull As Integer = -1
    Public Const TerrainRegular As Integer = 0
    Public Const TerrainWall As Integer = 1
    Public Const TerrainUnbreakable As Integer = 2
    Public Const TerrainGoal As Integer = 3
    Public Const TerrainJump As Integer = 4
    Public Const TerrainFast As Integer = 5
    Public Const TerrainSink As Integer = 6
    Public Const TerrainTeleport As Integer = 7

    '--Colors
    Public Const ColorCount As Integer = 8
    Public ColorRegular As Color = Color.Tan
    Public ColorWall As Color = Color.SaddleBrown
    Public ColorUnbreakable As Color = Color.Black
    Public ColorGoal As Color = Color.Gold
    Public ColorJump As Color = Color.DarkBlue 'Color.ForestGreen
    Public ColorFast As Color = Color.Turquoise
    Public ColorSink As Color = Color.DarkOrchid
    Public ColorTeleport As Color = Color.DarkSeaGreen

    '--Direction
    Public Const DirectionCount As Integer = 8
    Public DirectionNull As Integer = 0
    Public DirectionN As Integer = 1
    Public DirectionE As Integer = 2
    Public DirectionS As Integer = 3
    Public DirectionW As Integer = 4
    Public DirectionV As Integer = 5
    Public DirectionH As Integer = 6
    Public DirectionA As Integer = 7
    Public LetterNull As String = ""
    Public LetterN As String = "^"
    Public LetterE As String = ">"
    Public LetterS As String = "v"
    Public LetterW As String = "<"
    Public LetterV As String = "|"
    Public LetterH As String = "--"
    Public LetterA As String = "+"

    Public ColorNull As Color = Color.White
    Public ColorSelected As Color = Color.Red

#End Region

#Region " Global Functions "
    Function GetRandom(ByVal Min As Integer, ByVal Max As Integer) As Integer
        Randomize(Date.Now.Millisecond)
        Max += 1
        Dim TheResult As Integer = Int(Min + (Rnd() * (Max - Min)))
        Return TheResult
    End Function

    Function SafeDivide(ByVal A As Double, ByVal B As Double) As Double
        If A = 0 Or B = 0 Then
            Return 0
        Else
            Return A / B
        End If
    End Function


#End Region

End Module
