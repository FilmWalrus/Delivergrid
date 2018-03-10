Public Class Pair

#Region " Variables "
    Public X As Integer
    Public Y As Integer
#End Region

#Region " New "
    Public Sub New()
        X = -1
        Y = -1
    End Sub

    Public Sub New(ByVal theX As Integer, ByVal theY As Integer)
        X = theX
        Y = theY
    End Sub

#End Region
End Class
