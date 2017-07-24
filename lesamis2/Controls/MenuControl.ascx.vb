
Partial Class Controls_MenuControl
    Inherits System.Web.UI.UserControl



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Try
                Dim fileList() As String = HttpContext.Current.Request.FilePath.Split(CType("/", Char))
                Dim fileName As String = fileList(fileList.Length - 1)
                Select Case fileName.ToUpper
                    Case "PAGEONE.ASPX"
                        Me.hlHome.CssClass = "Disabled"
                    Case "MERCHANDISE.ASPX"
                        Me.hlMerch.CssClass = "Disabled"
                    Case "SYNOPSIS.ASPX"
                        Me.hlSynopsis.CssClass = "Disabled"
                    Case "HISTORY.ASPX"
                        Me.hlHistory.CssClass = "Disabled"
                    Case "GALLERY.ASPX"
                        Me.hlGallery.CssClass = "Disabled"
                    Case "TRAILER.ASPX"
                        Me.hlTrailer.CssClass = "Disabled"
                    Case "CONTACT.ASPX"
                        Me.hlContact.CssClass = "Disabled"
                    Case "CREDIT.ASPX"
                        Me.hlCredits.CssClass = "Disabled"
                    Case "SPONSORS.ASPX"
                        Me.hlSponsors.CssClass = "Disabled"
                    Case "PRESS.ASPX"
                        Me.hlPress.CssClass = "Disabled"
                End Select
            Catch ex As Exception

            End Try
        End If
    End Sub

End Class
