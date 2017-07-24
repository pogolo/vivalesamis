Imports PogoloUtilities

Namespace PogoloDataProvider.Data

    Public Class DataSortingHelper
        ' Tracks sorting calls

        Private uSortObj As SortHelperObject

        Public Sub New()
            Dim var As New PogoloUtilities.PogoloSessionVariable
            Dim obj As Object = var.GetVariable("SortOrderObj")
            If Not obj Is Nothing Then
                uSortObj = CType(obj, SortHelperObject)
            Else
                uSortObj = New SortHelperObject
                uSortObj.SortOrder = SortOrderEnum.Ascending
                uSortObj.SortCommand = ""
            End If
        End Sub

        Public Sub ClearSortObject()
            Dim var As New PogoloUtilities.PogoloSessionVariable
            var.RemoveVariable("SortOrderObj")
        End Sub

        Private Function getSortCommand() As String
            Return uSortObj.SortCommand()
        End Function

        Private Sub setSortCommand(ByVal sortCommand As String)
            ' Allows for reverse sort
            If uSortObj.SortCommand.ToUpper = sortCommand.ToUpper Then
                If uSortObj.SortOrder = SortOrderEnum.Ascending Then
                    uSortObj.SortOrder = SortOrderEnum.Descending
                Else
                    uSortObj.SortOrder = SortOrderEnum.Ascending
                End If
            End If
            uSortObj.SortCommand = sortCommand
            setSortOrderVariable(uSortObj)
        End Sub

        Private Function getSortOrder() As Integer
            Return uSortObj.SortOrder
        End Function

        Private Sub setSortOrder(ByVal sortOrder As Integer)
            uSortObj.SortOrder = sortOrder
            setSortOrderVariable(uSortObj)
        End Sub

        Private Sub setSortOrderVariable(ByVal sortObj As SortHelperObject)
            Dim var As New PogoloSessionVariable
            var.SetVariable("SortOrderObj", sortObj)
        End Sub

#Region "Properties"

        Public Enum SortOrderEnum
            Ascending
            Descending
        End Enum

        Public Property SortCommand() As String
            Get
                Return getSortCommand()
            End Get
            Set(ByVal Value As String)
                setSortCommand(Value)
            End Set
        End Property

        Public Property SortOrder() As Integer
            Get
                getSortOrder()
            End Get
            Set(ByVal Value As Integer)
                setSortOrder(Value)
            End Set
        End Property

#End Region

#Region "SortHelperObject Class"

        Private Class SortHelperObject

            Private uSortOrder As Integer
            Private uSortCommand As String

            Public Property SortOrder() As Integer
                Get
                    Return uSortOrder
                End Get
                Set(ByVal Value As Integer)
                    uSortOrder = Value
                End Set
            End Property

            Public Property SortCommand() As String
                Get
                    Return uSortCommand
                End Get
                Set(ByVal Value As String)
                    uSortCommand = Value
                End Set
            End Property

        End Class

#End Region

    End Class

End Namespace