
Public Class LesAmisUtilities

    Public Shared Function prependZeros(ByVal numZeros As Integer, ByVal strValue As Integer) As String
        Dim strDocID As String = CType(strValue, String)
        Dim length As Integer = strDocID.Length
        If length >= 7 Then
            Return strDocID
        Else
            numZeros = numZeros - length
            For ct As Integer = 1 To numZeros
                strDocID = "0" & strDocID
            Next
        End If
        Return strDocID
    End Function

    Public Shared Function defaultListItem() As ListItem
        Dim tmpItem As New ListItem
        tmpItem.Text = "< Choose One >"
        tmpItem.Value = "0"
        Return tmpItem
    End Function

    Public Shared Sub selectDropDownValue(ByVal DropDown As DropDownList, ByVal StrVal As String)
        DropDown.ClearSelection()
        Dim item As ListItem
        For Each item In DropDown.Items
            If item.Value = StrVal Then
                item.Selected = True
                Exit Sub
            End If
        Next
    End Sub

    Public Shared Sub selectDropDownText(ByVal DropDown As DropDownList, ByVal StrVal As String)
        DropDown.ClearSelection()
        Dim item As ListItem
        For Each item In DropDown.Items
            If item.Text = StrVal Then
                item.Selected = True
                Exit Sub
            End If
        Next
    End Sub

End Class
