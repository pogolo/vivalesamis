
Partial Class Contact
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Me.Name.Focus()
        End If
    End Sub

#Region "Methods & Functions"

    Function isFormValid() As Boolean

        If Name.Text.Trim = "" Then
            PogoloUtilities.PogoloDocumentUtilities.AddJavascriptAlertToPage(Me.Page, "Please fill out your name.")
            Me.Name.Focus()
            Return False
        End If

        If Email.Text = "" Then
            PogoloUtilities.PogoloDocumentUtilities.AddJavascriptAlertToPage(Me.Page, "Please fill out the email address.")
            Me.Email.Focus()
            Return False
        Else
            Dim comparison As New System.Text.RegularExpressions.Regex("\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*")
            If Not comparison.IsMatch(Me.Email.Text) Then
                PogoloUtilities.PogoloDocumentUtilities.AddJavascriptAlertToPage(Me.Page, "Please check the format of your email address.")
                Me.Email.Focus()
                Return False
            End If
        End If

        Return True
    End Function

#End Region

#Region "Events"

    Protected Sub cmdSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSubmit.Click
        Try
            If isFormValid() Then
                Dim emailCtrl As New BLL.EmailsController
                Dim emailInfo As New BLL.EmailsInfo
                emailInfo.EmailAddress = Me.Email.Text
                emailInfo.Comments = Me.Comments.Text
                emailInfo.FullName = Me.Name.Text
                emailInfo.DateAdded = DateTime.Now
                emailCtrl.Add(emailInfo)
                Me.pnlDone.Visible = True
                Me.MainPanel.Visible = False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
