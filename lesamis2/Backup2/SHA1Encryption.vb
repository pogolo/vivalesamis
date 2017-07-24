Imports Microsoft.VisualBasic
Imports System.Web.Security
Imports System.Security.Cryptography

Namespace PogoloUtilities

    Public Class SHA1Encryption

        ''' <summary>
        ''' Compare user's password against retrieved salt and password hash
        ''' </summary>
        ''' <param name="Password"></param>
        ''' <param name="Salt"></param>
        ''' <param name="PasswordHash"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsValidPassword(ByVal Password As String, ByVal Salt As String, ByVal PasswordHash As String) As Boolean
            Dim pwordHash As String = CreatePasswordHash(Password, Salt)
            If pwordHash = PasswordHash Then
                Return True
            End If
            Return False
        End Function

        Public Function CreateSalt(ByVal Size As Integer) As String
            ' Generate a cryptographic random number.
            Dim rng As New RNGCryptoServiceProvider()
            Dim buff(Size) As Byte
            rng.GetBytes(buff)
            ' Return a Base64 string representation of the random number
            Return Convert.ToBase64String(buff)
        End Function

        Public Function CreatePasswordHash(ByVal Password As String, ByVal Salt As String) As String
            Dim saltAndPwd As String = String.Concat(Password, Salt)
            Dim hashedPwd As String = FormsAuthentication.HashPasswordForStoringInConfigFile(saltAndPwd, "sha1")
            Return hashedPwd
        End Function

    End Class

End Namespace