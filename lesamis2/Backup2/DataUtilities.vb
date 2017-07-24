Imports System
Imports System.Data
Imports System.IO

Namespace PogoloUtilities

    Public Class DataUtilities

        Public Shared Function convertBooleanToInteger(ByVal boolObject As Boolean) As Integer
            If boolObject = True Then
                Return 1
            Else
                Return 0
            End If
        End Function

        Public Shared Function stripSQLFromString(ByVal stringValue As String) As String
            Dim badList As New ArrayList
            Dim newDic As New DictionaryEntry
            newDic.Key = "'"
            newDic.Value = "''"
            badList.Add(newDic)
            For Each dicEntry As DictionaryEntry In badList
                stringValue = stringValue.Replace(CType(dicEntry.Key, String), CType(dicEntry.Value, String))
            Next
            Return stringValue
        End Function

        Public Shared Function stripGoodSQLFromString(ByVal stringValue As String) As String
            Dim badList As New ArrayList
            Dim newDic As New DictionaryEntry
            newDic.Key = "''"
            newDic.Value = "'"
            badList.Add(newDic)
            For Each dicEntry As DictionaryEntry In badList
                stringValue = stringValue.Replace(CType(dicEntry.Key, String), CType(dicEntry.Value, String))
            Next
            Return stringValue
        End Function

    End Class

End Namespace