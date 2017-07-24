Imports System.Text.RegularExpressions

Namespace PogoloUtilities

    Public Class PogoloValidator

        Public Shared Function IsValidDate(ByVal DateString As String) As Boolean
            Try
                If IsDate(DateString) Then
                    Dim pattern As String = "^[0,1]?\d{1}\/(([0-2]?\d{1})|([3][0,1]{1}))\/(([1]{1}[9]{1}[9]{1}\d{1})|([2-9]{1}\d{3}))$"
                    Dim matcher As New Regex(pattern)
                    Return matcher.IsMatch(DateString)
                End If
                Return False
            Catch
                Return False
            End Try
        End Function

        Public Shared Function IsValidTime(ByVal TimeString As String) As Boolean
            Try
                If IsDate(Now.ToShortDateString & " " & TimeString) Then
                    Return True
                End If
                Return False
            Catch
                Return False
            End Try
        End Function

        Public Shared Function IsValidTimeRange(ByVal StartTime As DateTime, ByVal EndTime As DateTime) As Boolean
            If DateTime.Compare(StartTime, EndTime) >= 0 Then
                Return False
            End If
            Return True
        End Function

        Public Shared Function IsValidTimeRange(ByVal StartTime As String, ByVal EndTime As String) As Boolean
            If Not IsValidTime(StartTime) Or Not IsValidTime(EndTime) Then
                Return False
            End If
            If Not IsValidTimeRange(CType(Now.ToShortDateString & " " & StartTime, DateTime), CType(Now.ToShortDateString & " " & EndTime, DateTime)) Then
                Return False
            End If
            Return True
        End Function

        Public Shared Function IsValidDateRange(ByVal StartDate As DateTime, ByVal EndDate As DateTime) As Boolean
            If DateTime.Compare(StartDate.Date, EndDate.Date) >= 0 Then
                Return False
            End If
            Return True
        End Function

        Public Shared Function IsValidDateRange(ByVal StartDate As String, ByVal EndDate As String) As Boolean
            '''' <Summary>StartDate must be < EndDate - no Equals</Summary>
            If Not IsValidDate(StartDate) Or Not IsValidDate(EndDate) Then
                Return False
            End If
            If IsValidDateRange(CType(StartDate, DateTime), CType(EndDate, DateTime)) Then
                Return True
            End If
            Return False
        End Function

        Public Shared Function IsValidEmail(ByVal InputEmail As String) As Boolean
            Dim strRegex As String = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}" & "\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\" & ".)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
            Dim re As New Regex(strRegex)
            If re.IsMatch(InputEmail) Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Shared Function ContainsInvalidCharacters(ByVal InputString As String) As Boolean
            Try
                Dim characters As String = "~^`^!^@^#^$^%^&^*^(^)^-^_^=^+^,^<^>^.^;^:^'^'^[^]^{^}^\^|^/^?"
                Dim characterList() As String = characters.Split(CType("^", Char))
                For Each strValue As String In characterList
                    If InputString.IndexOf(strValue) > 0 Then
                        Return True
                        Exit Function
                    End If
                Next
                Return False
            Catch
                Return False
            End Try
        End Function

        Public Shared Function IsValidURL(ByVal URL As String) As Boolean
            Dim strRegex As String = "http://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?"
            Dim re As New Regex(strRegex)
            If re.IsMatch(URL) Then
                Return True
            Else
                Return False
            End If
        End Function

    End Class

End Namespace