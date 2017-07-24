Imports System
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary

Namespace PogoloUtilities

    Public Class ImageFileObject

        ''' <summary>
        ''' Writes a byte() file to a web page
        ''' </summary>
        ''' <param name="CurPage">Target page</param>
        ''' <param name="FileData">Byte() data</param>
        ''' <param name="FileType">Type of file</param>
        ''' <param name="FileName">File Name</param>
        ''' <remarks></remarks>
        Public Sub DeliverFile(ByVal CurPage As System.Web.UI.Page, ByVal FileData() As Byte, ByVal FileType As String, ByVal FileName As String)
            CurPage.Response.ContentType = FileType
            CurPage.Response.AppendHeader("Content-Disposition", "attachment; filename=" + FileName)
            CurPage.Response.BinaryWrite(FileData)
        End Sub

        Public Function ConvertByteToBitMap(ByVal imageByte As Byte()) As Bitmap
            Dim ms As New System.IO.MemoryStream
            ms.Write(imageByte, 0, imageByte.Length)
            Dim bmp As New Bitmap(ms)
            Return bmp
        End Function

        Public Function GetFileName(ByVal fpath As String) As String
            Return Mid(fpath, fpath.LastIndexOf("\") + 2)
        End Function

        Public Function CreateThumbnail(ByVal ImageStream As Stream, ByVal tWidth As Double, ByVal tHeight As Double) As Byte()
            ' Re-added 9/9/2007 - Deprecated!
            Dim g As System.Drawing.Image = System.Drawing.Image.FromStream(ImageStream)
            Dim thumbSize As New Size
            thumbSize = NewThumbSize(g.Width, g.Height, tWidth, tHeight)
            Dim imgOutput As New Bitmap(g, thumbSize.Width, thumbSize.Height)
            Dim imgStream As New MemoryStream
            Dim thisFormat As System.Drawing.Imaging.ImageFormat = g.RawFormat
            imgOutput.Save(imgStream, thisFormat)
            Dim imgbin(CType(imgStream.Length, Integer)) As Byte
            imgStream.Position = 0
            Dim n As Int32 = imgStream.Read(imgbin, 0, imgbin.Length)
            g.Dispose()
            imgOutput.Dispose()
            Return imgbin
        End Function

        Public Function GetByteArrayFromFileField(ByVal FileField As System.Web.UI.HtmlControls.HtmlInputFile) As Byte()
            Try
                ' Returns a byte array from the passed file field controls file
                Dim intFileLength As Integer
                Dim bytData() As Byte
                Dim objStream As System.IO.Stream
                If FileFieldSelected(FileField) Then
                    intFileLength = FileField.PostedFile.ContentLength
                    ReDim bytData(intFileLength)
                    objStream = FileField.PostedFile.InputStream
                    objStream.Read(bytData, 0, intFileLength)
                    Return bytData
                End If
                Return Nothing
            Catch
                Return Nothing
            End Try
        End Function

        Public Function GetByteArrayFromFilePath(ByVal FilePath As String) As Byte()
            Try
                ' Send filepath and returns byte file
                Dim bytData() As Byte
                Dim objStream As System.IO.Stream

                Dim objBr As New BinaryReader(File.OpenRead(FilePath))
                objStream = objBr.BaseStream
                Dim fileLength As Integer = CType(objStream.Length, Integer)
                ReDim bytData(fileLength)
                objStream.Read(bytData, 0, fileLength)
                Return bytData
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Public Function GetStreamFromFilePath(ByVal FilePath As String) As System.IO.Stream
            Try
                Dim objBr As New BinaryReader(File.OpenRead(FilePath))
                Return objBr.BaseStream
            Catch
                Return Nothing
            End Try
        End Function

        Public Function FileFieldType(ByVal FileField As System.Web.UI.HtmlControls.HtmlInputFile) As String
            ' Returns the type of the posted file
            If Not FileField.PostedFile Is Nothing Then
                Return FileField.PostedFile.ContentType
            End If
            Return ""
        End Function

        Public Function FileFieldLength(ByVal FileField As System.Web.UI.HtmlControls.HtmlInputFile) As Integer
            ' Returns the length of the posted file
            If Not FileField.PostedFile Is Nothing Then
                Return FileField.PostedFile.ContentLength
            End If
        End Function

        Public Function FileFieldFilename(ByVal FileField As System.Web.UI.HtmlControls.HtmlInputFile) As String
            ' Returns the core filename of the posted file
            If Not FileField.PostedFile Is Nothing Then
                Return System.IO.Path.GetFileName(FileField.PostedFile.FileName)
            End If
            Return ""
        End Function

        Public Function FileNameFromPath(ByVal FileName As String) As String
            ' Returns the core filename of the posted file
            If Trim(FileName) <> "" Then
                Return System.IO.Path.GetFileName(FileName)
            End If
            Return ""
        End Function

        Public Function FileFieldSelected(ByVal FileField As System.Web.UI.HtmlControls.HtmlInputFile) As Boolean
            ' Returns a True if the passed
            ' FileField has had a user post a file
            If FileField.PostedFile Is Nothing Then Return False
            If FileField.PostedFile.ContentLength = 0 Then Return False
            Return True
        End Function

        Public Function FileIsImage(ByVal FileField As System.Web.UI.HtmlControls.HtmlInputFile) As Boolean
            If FileField.PostedFile.FileName = "" Then Return False
            If LCase(FileFieldType(FileField)).StartsWith("image") Then
                Return True
            Else
                Return False
            End If
        End Function

        Private Const uImageFileTypes As String = "JPG,TIFF,TIF,ICO,JPEG,PNG,PING,BMP"

        Public Function FileIsImage(ByVal FileObj As System.IO.FileInfo) As Boolean
            Dim fileTypeList() As String = uImageFileTypes.Split(CType(",", Char))
            For Each extension As String In fileTypeList
                If FileObj.Extension.ToUpper = "." & extension.ToUpper Then
                    Return True
                End If
            Next
            Return False
        End Function

        Public Function FileIsImage(ByVal FileName As String) As Boolean
            Dim fileTypeList() As String = uImageFileTypes.Split(CType(",", Char))
            For Each extension As String In fileTypeList
                If FileName.ToUpper.EndsWith("." & extension.ToUpper) Then
                    Return True
                End If
            Next
            Return False
        End Function

        Public Function GetImageDimensionsFromByte(ByVal ImageByte As Byte()) As Size
            Try
                Dim imageStream As New System.IO.MemoryStream(ImageByte)
                Dim originalBMP As New Bitmap(imageStream)
                Return New Size(originalBMP.Width, originalBMP.Height)
            Catch ex As Exception
                Return New Size(0, 0)
            End Try
        End Function

        Public Function GetNewImageDimensions(ByVal OriginalWidth As Integer, ByVal OriginalHeight As Integer, ByVal NewWidth As Integer, ByVal NewHeight As Integer) As Size
            Try
                Return NewThumbSize(OriginalWidth, OriginalHeight, NewWidth, NewHeight)
            Catch ex As Exception

            End Try
        End Function

        Public Function GetNewImageDimensions(ByVal ImageByte As Byte(), ByVal Width As Integer, ByVal Height As Integer) As Size
            Try
                Dim imageStream As New System.IO.MemoryStream(ImageByte)
                Dim originalBMP As New Bitmap(imageStream)
                Dim originalImageFormat As Imaging.ImageFormat = originalBMP.RawFormat

                ' Get size while preserving aspect-ratio
                Dim originalImage As Drawing.Image = Drawing.Image.FromStream(imageStream)
                Return NewThumbSize(originalImage.Width, originalImage.Height, Width, Height)
            Catch ex As Exception
                Return New Size(Width, Height)
            End Try
        End Function

        ''' <summary>
        ''' Create high-resolution jpg image from a Byte() array
        ''' </summary>
        ''' <param name="CurrentPage"></param>
        ''' <param name="ImageByte"></param>
        ''' <param name="Width"></param>
        ''' <param name="Height"></param>
        ''' <remarks></remarks>
        Public Sub CreateThumbnailImage(ByVal CurrentPage As Page, ByVal ImageByte As Byte(), ByVal Width As Integer, ByVal Height As Integer)
            Try
                Dim imageStream As New System.IO.MemoryStream(ImageByte)
                Dim originalBMP As New Bitmap(imageStream)
                Dim originalImageFormat As Imaging.ImageFormat = originalBMP.RawFormat

                ' Get size while preserving aspect-ratio
                Dim originalImage As Drawing.Image = Drawing.Image.FromStream(imageStream)
                Dim thumbSize As Drawing.Size = NewThumbSize(originalImage.Width, originalImage.Height, Width, Height)
                Dim newHeight As Integer = thumbSize.Height
                Dim newWidth As Integer = thumbSize.Width

                ' Build output Bitmap
                Dim outputBMP As Bitmap = New Bitmap(thumbSize.Width, thumbSize.Height, PixelFormat.Format32bppArgb)
                Dim outPutGraphic As Graphics = Graphics.FromImage(outputBMP)
                outPutGraphic.CompositingMode = CompositingMode.SourceCopy
                outPutGraphic.InterpolationMode = InterpolationMode.HighQualityBicubic
                outPutGraphic.CompositingQuality = CompositingQuality.HighQuality
                outPutGraphic.SmoothingMode = SmoothingMode.HighQuality
                outPutGraphic.PixelOffsetMode = PixelOffsetMode.HighQuality

                Dim destinationRect As New Rectangle(0, 0, newWidth, newHeight)
                outPutGraphic.DrawImage(originalBMP, destinationRect, 0, 0, originalImage.Width, originalImage.Height, GraphicsUnit.Pixel)
                originalBMP.Dispose()

                ' Set page format encoding
                Dim Info() As System.Drawing.Imaging.ImageCodecInfo = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                Dim Params As New System.Drawing.Imaging.EncoderParameters(1)
                Params.Param(0) = New EncoderParameter(Encoder.Quality, 100L)
                CurrentPage.Response.ContentType = Info(0).MimeType
                outputBMP.Save(CurrentPage.Response.OutputStream, Info(1), Params)
                outputBMP.Dispose()

                Dim converter As New System.Drawing.ImageConverter()
                converter.ConvertTo(outputBMP, GetType(WebControls.Image))
            Catch

            End Try
        End Sub

        Public Function NewThumbSize(ByVal currentwidth As Double, ByVal currentheight As Double, ByVal newWidth As Double, ByVal newHeight As Double) As Size
            ' Calculate the Size of the New image
            Dim tempMultiplier As Double
            If currentheight > currentwidth Then ' portrait
                tempMultiplier = newHeight / currentheight
            Else
                tempMultiplier = newWidth / currentwidth
            End If
            Dim NewSize As New Size(CInt(currentwidth * tempMultiplier), CInt(currentheight * tempMultiplier))
            Return NewSize
        End Function

        Public Sub SaveImageByte(ByVal ImageFile As Byte(), ByVal TargetPath As String)
            Dim oFileStream As New System.IO.FileStream(TargetPath, System.IO.FileMode.Create, FileAccess.Write)
            oFileStream.Write(ImageFile, 0, ImageFile.Length)
            oFileStream.Close()
        End Sub

    End Class

End Namespace