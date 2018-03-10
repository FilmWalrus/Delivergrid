Public Class TerrainSquare

#Region " Variables "
    Public X As Integer = 0
    Public Y As Integer = 0

    Public TerrainType As Integer = -1
    Public Value As Double = 0.0

    Public MinDistance As Double = MaxDistance
    'Public ParentRow As Integer = -1
    'Public ParentCol As Integer = -1
    Public Parent As New Pair(-1, -1)

    Public VisitedKey As Integer = -1

#End Region

#Region " New "
    Public Sub New()

    End Sub

    Public Sub New(ByVal theX As Integer, ByVal theY As Integer)
        X = theX
        Y = theY
        TerrainType = TerrainRegular
    End Sub

    Public Sub New(ByVal theX As Integer, ByVal theY As Integer, ByVal theTerrain As Integer, ByVal theValue As Integer)
        X = theX
        Y = theY
        TerrainType = theTerrain
        Value = theValue
    End Sub

    Public Sub New(ByVal other As TerrainSquare)
        X = other.X
        Y = other.Y
        TerrainType = other.TerrainType
        Value = other.Value
    End Sub

    Function BuildSaveString() As String
        Dim TheString As String
        TheString &= TerrainType & ","
        TheString &= Value
        Return TheString
    End Function

    Sub ReadLoadString(ByVal TheString As String)
        Dim TheSections() As String = TheString.Split(",")
        If TheSections.Length >= 2 Then
            TerrainType = TheSections(0)
            Value = TheSections(1)
        End If
    End Sub
#End Region

End Class
