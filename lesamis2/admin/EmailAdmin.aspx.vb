
Partial Class admin_EmailAdmin
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim app As New AppVariable
        Dim UserID As Integer = app.GetVariable("UserID")
        If UserID > 0 Then
            loadGrid()
        Else
            Response.Redirect("Login.aspx")
        End If
    End Sub

#Region "Methods & Functions"

    Private Sub loadGrid()
        Try
            Dim emailCtrl As New BLL.EmailsController
            Dim emailList As ArrayList = emailCtrl.List
            Me.EmailDataGrid.DataSource = emailList
            Me.EmailDataGrid.DataBind()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub togglePanels()
        If Me.ClearPanel.Visible = True Then
            Me.ClearPanel.Visible = False
            Me.PromptPanel.Visible = True
        Else
            Me.ClearPanel.Visible = True
            Me.PromptPanel.Visible = False
        End If
    End Sub

#End Region

#Region "Events"

    Protected Sub cmdYes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdYes.Click
        Try
            Dim emailCtrl As New BLL.EmailsController
            Dim emailList As ArrayList = emailCtrl.List
            For Each emailInfo As BLL.EmailsInfo In emailList
                emailCtrl.Delete(emailInfo.EmailID)
            Next
            loadGrid()
            togglePanels()
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub cmdNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdNo.Click
        togglePanels()
    End Sub

    Protected Sub cmdClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClear.Click
        togglePanels()
    End Sub

    Protected Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Response.Redirect("~/pageone.aspx")
    End Sub

#End Region

End Class
