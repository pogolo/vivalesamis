Imports System
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Text.RegularExpressions
Imports PogoloUtilities

Namespace PogoloUtilities

    Public Class PogoloCookieUtilities

        Private uCurrentPage As Page

        Public Sub New(ByVal currentPage As Page)
            uCurrentPage = currentPage
        End Sub

        Public Function GetCookie(ByVal CookieName As String) As HttpCookie
            Try
                Return uCurrentPage.Request.Cookies(CookieName)
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Public Function GetCookieValue(ByVal CookieName As String, Optional ByVal UseDecryption As Boolean = False) As String
            Try
                Dim cookie As HttpCookie = GetCookie(CookieName)
                If Not cookie Is Nothing Then
                    If UseDecryption Then
                        Return PogoloUtilities.EncryptionHelper.Decrypt(Regex.Unescape(cookie.Value))
                    Else
                        Return Regex.Unescape(cookie.Value)
                    End If
                End If
                Return ""
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Public Sub SetCookieValue(ByVal CookieName As String, ByVal CookieValue As String, ByVal ExpirationDate As DateTime, Optional ByVal Domain As String = "", Optional ByVal Encrypt As Boolean = False)
            Try
                Dim newCookie As New HttpCookie(CookieName)
                'newCookie.Path = "/"
                If Domain <> "" And Domain.ToUpper <> "LOCALHOST" Then
                    newCookie.Domain = "." & Domain
                End If
                newCookie.Expires = ExpirationDate
                If Encrypt Then
                    newCookie.Value = Regex.Escape(PogoloUtilities.EncryptionHelper.Encrypt(CookieValue))
                Else
                    newCookie.Value = Regex.Escape(CookieValue)
                End If
                uCurrentPage.Response.Cookies.Add(newCookie)
            Catch ex As Exception

            End Try
        End Sub

        Public Sub RemoveCookie(ByVal cookieName As String, Optional ByVal Domain As String = "")
            Try
                Dim expirationDate As DateTime = Now
                Dim newCookie As New HttpCookie(cookieName)
                'newCookie.Path = "/"
                If Domain <> "" And Domain.ToUpper <> "LOCALHOST" Then
                    newCookie.Domain = "." & Domain
                End If
                newCookie.Expires = expirationDate.AddDays(-1)
                uCurrentPage.Response.Cookies.Add(newCookie)
            Catch ex As Exception

            End Try
        End Sub

    End Class

End Namespace