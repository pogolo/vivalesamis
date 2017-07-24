Imports System
Imports System.Web
Imports System.Web.HttpApplicationState

Namespace PogoloUtilities

    Public Class PogoloApplicationState

        Private uMyApp As System.Web.HttpApplicationState

        Public Sub New()
            uMyApp = HttpContext.Current.Application
        End Sub

        Public Sub SetVariable(ByVal VariableName As String, ByVal Obj As Object)
            RemoveVariable(VariableName)
            uMyApp.Add(VariableName, Obj)
        End Sub

        Public Sub RemoveVariable(ByVal VariableName As String)
            Dim curVar As Object = uMyApp.Get(VariableName)
            uMyApp.Remove(VariableName)
        End Sub

        Public Function GetVariable(ByVal VariableName As String) As Object
            Try
                Dim curVar As Object = uMyApp.Get(VariableName)
                If Not curVar Is Nothing Then
                    Return curVar
                Else
                    Return Nothing
                End If
            Catch
                Return Nothing
            End Try
        End Function

    End Class

    Public Class PogoloSessionVariable

        Private uMySession As System.Web.SessionState.HttpSessionState

        Public Sub New()
            uMySession = HttpContext.Current.Session
        End Sub

        Public Sub SetVariable(ByVal VariableName As String, ByVal Obj As Object)
            Try
                RemoveVariable(VariableName)
                uMySession.Add(VariableName, Obj)
            Catch ex As Exception

            End Try
        End Sub

        Public Sub RemoveVariable(ByVal VariableName As String)
            Dim curVar As Object = uMySession.Item(VariableName)
            uMySession.Remove(VariableName)
        End Sub

        Public Function GetVariable(ByVal VariableName As String) As Object
            Try
                Dim curVar As Object = uMySession.Item(VariableName)
                If Not curVar Is Nothing Then
                    Return curVar
                Else
                    Return Nothing
                End If
            Catch
                Return Nothing
            End Try
        End Function

    End Class

End Namespace