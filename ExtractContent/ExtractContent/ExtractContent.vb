Public Class ExtractContent
    Public Function ExtractContentWithKeyWords(ByVal strContent As String, strLookFor As String, strLookFor2 As String, Optional ByVal s1 As Long = -1) As String
        Dim s, t As Long
        If s1 <> -1 Then
            s = InStr(s1, strContent, strLookFor, vbTextCompare)
        Else
            s = InStr(1, strContent, strLookFor, vbTextCompare)
        End If
        If s > 0 Then
            t = InStr(s + 1, strContent, strLookFor2, vbTextCompare)
            If t > 0 Then
                t = t + Len(strLookFor2)
                strContent = Mid(strContent, s, (t - s))
            End If
        End If
        ExtractContentWithKeyWords = strContent
    End Function

    Public Function ExtractContentWith1stKeyWord(ByVal strContent As String, strLookFor As String, strLookFor2 As String, Optional ByVal s1 As Long = -1) As String
        Dim s, t As Long
        If s1 <> -1 Then
            s = InStr(s1, strContent, strLookFor, vbTextCompare)
        Else
            s = InStr(1, strContent, strLookFor, vbTextCompare)
        End If
        If s > 0 Then
            t = InStr(s, strContent, strLookFor2, vbTextCompare)
            If t > 0 Then
                strContent = Mid(strContent, s, (t - s))
            End If
        End If
        ExtractContentWith1stKeyWord = strContent
    End Function

    Public Function ExtractContentWith2ndKeyWord(ByVal strContent As String, strLookFor As String, strLookFor2 As String, Optional ByVal s1 As Long = -1) As String
        Dim s, t As Long
        If s1 <> -1 Then
            s = InStr(s1, strContent, strLookFor, vbTextCompare)
        Else
            s = InStr(1, strContent, strLookFor, vbTextCompare)
        End If
        If s > 0 Then
            s = s + Len(strLookFor)
            t = InStr(s, strContent, strLookFor2, vbTextCompare)
            If t > 0 Then
                t = t + Len(strLookFor2)
                strContent = Mid(strContent, s, (t - s))
            End If
        End If
        ExtractContentWith2ndKeyWord = strContent
    End Function

    Public Function ExtractContentWithoutKeyWords(ByVal strContent As String, strLookFor As String, strLookFor2 As String, Optional ByVal s1 As Long = -1, Optional ByVal strLookFor3 As String = "XYZ") As String
        Dim s, t As Long
        If s1 <> -1 Then
            If strLookFor3 = "XYZ" Then
                s = s1
            Else
                s = s1 + Len(strLookFor3)
            End If
            s = InStr(s, strContent, strLookFor, vbTextCompare)
        Else
            s = InStr(1, strContent, strLookFor, vbTextCompare)
        End If
        If s > 0 Then
            s = s + Len(strLookFor)
            t = InStr(s, strContent, strLookFor2, vbTextCompare)
            If t > 0 Then
                strContent = Mid(strContent, s, (t - s))
            End If
        End If
        ExtractContentWithoutKeyWords = strContent
    End Function

    Public Function ExtractContentWithoutKeyWordsWithLoop(ByVal strText As String, strLookFor As String, strLookFor2 As String, Optional ByVal s1 As Long = -1) As String
        Dim s, t As Long
        Dim strContent As String = ""
        If s1 <> -1 Then
            s = InStr(s1, strText, strLookFor, vbTextCompare)
        Else
            s = InStr(1, strText, strLookFor, vbTextCompare)
        End If
        While s > 0
            s = s + Len(strLookFor)
            t = InStr(s, strText, strLookFor2, vbTextCompare)
            If t > 0 Then
                strContent = strContent & Mid(strText, s, (t - s)) & Environment.NewLine
            End If
            s = InStr(s + 1, strText, strLookFor, vbTextCompare)
        End While
        ExtractContentWithoutKeyWordsWithLoop = strContent
    End Function

    Public Function ExtractContentWithKeyWordsWithLoop(ByVal strText As String, strLookFor As String, strLookFor2 As String, Optional ByVal s1 As Long = -1) As String
        Dim s, t As Long
        Dim strContent As String = ""
        If s1 <> -1 Then
            s = InStr(s1, strText, strLookFor, vbTextCompare)
        Else
            s = InStr(1, strText, strLookFor, vbTextCompare)
        End If
        While s > 0
            t = InStr(s, strText, strLookFor2, vbTextCompare)
            If t > 0 Then
                t = t + Len(strLookFor2)
                strContent = strContent & Mid(strText, s, (t - s)) & Environment.NewLine
            End If
            s = InStr(s + 1, strText, strLookFor, vbTextCompare)
        End While
        ExtractContentWithKeyWordsWithLoop = strContent
    End Function

End Class
