Imports System
Imports System.Web
Imports System.Web.HttpApplicationState

Public Class AppVariable

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
        If Not curVar Is Nothing Then
            uMyApp.Remove(VariableName)
        End If
    End Sub

    Public Function GetVariable(ByVal VariableName As String) As Object
        Dim curVar As Object = uMyApp.Get(VariableName)
        If Not curVar Is Nothing Then
            Return curVar
        Else
            Return Nothing
        End If
    End Function

End Class
