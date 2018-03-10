Option Compare Text

Public Class Parser
    '-----------------------------------------------
    '-- This class is for the parsing of filenames
    '-----------------------------------------------
    Function RemoveSlash(ByVal ThePath As String, Optional ByVal BlankPath As String = "") As String
        If ThePath.Trim = "" Then
            '-- The original path is blank so we can return the 'BlankPath'
            Return BlankPath
            Exit Function
        End If
        If Microsoft.VisualBasic.Right(ThePath, 1) = "\" Then
            ThePath = Microsoft.VisualBasic.Left(ThePath, ThePath.Length - 1)
        Else
            '-- Do Nothing; There is no slash to remove
        End If
        Return ThePath

    End Function
    Function AddSlash(ByVal ThePath As String, Optional ByVal BlankPath As String = "") As String
        If ThePath.Trim = "" Then
            '-- The original path is blank so we can return the 'BlankPath'
            Return BlankPath
            Exit Function
        End If
        If Microsoft.VisualBasic.Right(ThePath, 1) = "\" Then
            '-- Do Nothing; There is already a slash there
        Else
            ThePath &= "\"
        End If
        Return ThePath

    End Function
    Function ParseRightMost(ByVal Path As String) As String
        Dim ShortPath As String
        Dim i As Integer
        Dim CurrentChar As String
        Dim Length As Integer = Microsoft.VisualBasic.Len(Path)
        For i = 1 To Length
            Try
                CurrentChar = Microsoft.VisualBasic.Mid(Path, Length - i, 1)
            Catch ex As Exception
                CurrentChar = ""
            End Try

            If CurrentChar = "\" Then
                ShortPath = Microsoft.VisualBasic.Right(Path, i)
                Exit For
            End If
        Next
        Return ShortPath

    End Function


    '-----------------------------------------------------
    '-- This section is for reforming strings
    '-- Ex: [Name] automatically puts in a certain value
    '-----------------------------------------------------
#Region " Variables "
    Dim OriginalValues(0) As String
    Dim ReplacementValues(0) As String
    '--Values(OriginalString,Replacement)

#End Region

    Sub AddValue(ByVal TheOriginalString As String, ByVal TheReplacementString As String)
        ReDim Preserve OriginalValues(OriginalValues.GetLength(0))
        OriginalValues(OriginalValues.GetLength(0) - 1) = TheOriginalString
        '--
        ReDim Preserve ReplacementValues(ReplacementValues.GetLength(0))
        ReplacementValues(ReplacementValues.GetLength(0) - 1) = TheReplacementString

    End Sub
    Sub ClearValues()
        ReDim Preserve OriginalValues(0)
        ReDim Preserve ReplacementValues(0)

    End Sub
    Function ReformString(ByRef TheStringToReform As String) As String
        Dim TheReturnString As String = TheStringToReform
        Dim i As Integer
        Dim j As Integer
StartAgain:
        Dim TotalLength As Integer = TheReturnString.Length
        For i = 1 To OriginalValues.GetLength(0) - 1
            Dim TheLength As Integer = OriginalValues(i).Length
            Dim TheOrigString As String = OriginalValues(i)
            Dim TheReplaceString As String = ReplacementValues(i)
            Dim CurrentString As String
            For j = 0 To TotalLength - TheLength
                CurrentString = Microsoft.VisualBasic.Mid(TheReturnString, j + 1, TheLength)
                If CurrentString = TheOrigString Then
                    Dim BeginString As String
                    Dim EndString As String
                    '--
                    BeginString = Microsoft.VisualBasic.Left(TheReturnString, j)
                    EndString = Microsoft.VisualBasic.Right(TheReturnString, TotalLength - TheLength - BeginString.Length)
                    '--
                    TheReturnString = BeginString & TheReplaceString & EndString
                    '--
                    GoTo StartAgain
                End If
            Next
        Next
        Return TheReturnString

    End Function

    Function GetLevelName(ByVal FileName As String) As String
        Dim i, j As Integer
        'Last Slash
        i = FileName.LastIndexOf("\")
        If i < 0 Then
            i = 0
        End If
        'First Period
        j = FileName.IndexOf(".")
        If j < 0 Then
            j = FileName.Length - 1
        End If

        Dim ReturnString As String = ""
        ReturnString = FileName.Substring(i + 1, j - i - 1)
        Return ReturnString
    End Function

End Class
