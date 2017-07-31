
Partial Class HTTPErrors_AllErrors
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
    End Sub

    Private Sub cmdHome_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHome.Click
        Response.Redirect("../pageoneFS.htm")
    End Sub

End Class
