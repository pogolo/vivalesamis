
Partial Class admin_Login
    Inherits System.Web.UI.Page

    Protected Sub cmdLogin_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdLogin.Click
        Try
            Dim userCtrl As New BLL.UserNamesController
            Dim userList As ArrayList = userCtrl.GetByUserName(Me.Username.Text)
            If userList.Count > 0 Then
                Dim userInfo As BLL.UserNamesInfo = CType(userList(0), BLL.UserNamesInfo)
                If userInfo.Password.Equals(Me.Password.Text) Then
                    Dim app As New AppVariable
                    app.SetVariable("UserID", userInfo.UserID)
                    Response.Redirect("~/admin/EmailAdmin.aspx")
                End If
            End If
        Catch

        End Try
    End Sub

    Protected Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Response.Redirect("~/PageOne.aspx")
    End Sub
End Class
