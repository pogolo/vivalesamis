Imports System
Imports System.Data
Imports System.Configuration
Imports System.IO
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls

Namespace PogoloDataProvider.Data

    Public Class GridViewExcelExport

        ''' <summary>
        ''' Parse a GridView control and render it as an Excel file.
        ''' </summary>
        ''' <param name="FileName">File name of Excel file</param>
        ''' <param name="CurGridView">GridView to re-render</param>
        ''' <remarks></remarks>
        Public Sub Export(ByVal FileName As String, ByVal CurGridView As GridView)
            HttpContext.Current.Response.Clear()
            HttpContext.Current.Response.AddHeader("content-disposition", String.Format("attachment; filename={0}", FileName))
            HttpContext.Current.Response.ContentType = "application/ms-excel"

            Dim table As New Table()
            If Not CurGridView.HeaderRow Is Nothing Then
                table.Rows.Add(CurGridView.HeaderRow)
            End If
            For Each row As GridViewRow In CurGridView.Rows
                row = replaceControls(row)
                table.Rows.Add(row)
            Next
            If Not CurGridView.FooterRow Is Nothing Then
                table.Rows.Add(CurGridView.FooterRow)
            End If

            ' Render the table into the htmlwriter
            Dim sw As New StringWriter()
            Dim htw As New HtmlTextWriter(sw)
            table.RenderControl(htw)

            ' Render the htmlwriter into the response
            HttpContext.Current.Response.Write(sw.ToString())
            HttpContext.Current.Response.End()
        End Sub

        Private Function replaceControls(ByVal row As GridViewRow) As GridViewRow
            ' Remove auto-created checkboxes - won't render
            For Each cell As TableCell In row.Cells
                For Each ctrl As Control In cell.Controls
                    If TypeOf (ctrl) Is CheckBox Then
                        Dim boolVal As Boolean = CType(ctrl, CheckBox).Checked
                        cell.Controls.Remove(ctrl)
                        If boolVal Then
                            cell.Text = "TRUE"
                        Else
                            cell.Text = "FALSE"
                        End If
                    End If
                Next
            Next
            Return row
        End Function

    End Class

End Namespace