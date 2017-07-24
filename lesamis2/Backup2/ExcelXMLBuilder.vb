Imports Microsoft.VisualBasic
Imports System.Xml

Namespace PogoloUtilities

    ''' <summary>
    ''' Builds an XML document that can be opened in Excel
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ExcelXMLBuilder

        Private uExcelDoc As XmlDocument
        Private Const XMLNS_URN As String = "urn:schemas-microsoft-com:office:spreadsheet"
        Private Const XLMNS_O As String = "urn:schemas-microsoft-com:office:office"
        Private Const XLMNS_X As String = "urn:schemas-microsoft-com:office:excel"
        Private Const XLMNS_SS As String = "urn:schemas-microsoft-com:office:spreadsheet"
        Private Const XLMNS_HTML As String = "http://www.w3.org/TR/REC-html40"

        Public Sub New()
            uExcelDoc = initializeExcelDocument("", DateTime.Now)
        End Sub

        Public Sub New(ByVal Author As String, ByVal CreatedDate As DateTime)
            uExcelDoc = initializeExcelDocument(Author, CreatedDate)
        End Sub

        Public Function AddWorkSheet(ByVal WorkSheetName As String, ByVal NumberOfColumns As Integer) As XmlElement
            Dim ws As XmlElement = uExcelDoc.CreateElement("Worksheet")
            Dim att As XmlAttribute = uExcelDoc.CreateAttribute("ss", "Name", XLMNS_SS)
            att.Value = WorkSheetName
            ws.Attributes.Append(att)

            ' Add table element
            Dim tb As XmlElement = uExcelDoc.CreateElement("Table")
            att = uExcelDoc.CreateAttribute("ss", "ExpandedColumnCount", XLMNS_SS)
            att.Value = NumberOfColumns.ToString
            tb.Attributes.Append(att)

            att = uExcelDoc.CreateAttribute("x", "FullColumns", XLMNS_X)
            att.Value = "1"
            tb.Attributes.Append(att)

            att = uExcelDoc.CreateAttribute("x", "FullRows", XLMNS_X)
            att.Value = "1"
            tb.Attributes.Append(att)

            ' Add columns to table
            For i As Integer = 1 To NumberOfColumns
                Dim col As XmlElement = uExcelDoc.CreateElement("Column")
                att = uExcelDoc.CreateAttribute("ss", "AutoFitWidth", XLMNS_SS)
                att.Value = "0"
                col.Attributes.Append(att)

                att = uExcelDoc.CreateAttribute("ss", "Width", XLMNS_SS)
                att.Value = "200"
                col.Attributes.Append(att)

                tb.AppendChild(col)
            Next
            ws.AppendChild(tb)

            ' WorksheetOptions element
            Dim options As XmlElement = uExcelDoc.CreateElement("WorksheetOptions")
            att = uExcelDoc.CreateAttribute("xmnlns")
            att.Value = XLMNS_X
            options.Attributes.Append(att)

            Dim child As XmlElement = uExcelDoc.CreateElement("PageSetup")
            options.AppendChild(child)

            child = uExcelDoc.CreateElement("Print")
            options.AppendChild(child)

            child = uExcelDoc.CreateElement("ProtectObjects")
            child.InnerText = "False"
            options.AppendChild(child)

            child = uExcelDoc.CreateElement("ProtectScenarios")
            child.InnerText = "False"
            options.AppendChild(child)
            ws.AppendChild(options)

            uExcelDoc.DocumentElement.AppendChild(ws)
            Return ws
        End Function

        Public Function GetAllWorkSheets() As ArrayList
            Dim workSheetList As New ArrayList
            For Each ws As XmlElement In uExcelDoc.FirstChild.ChildNodes
                If ws.Name.ToUpper = "WORKSHEET" Then
                    workSheetList.Add(ws)
                End If
            Next
            Return workSheetList
        End Function

        Public Function GetWorkSheet(ByVal WorkSheetIndex As Integer) As XmlElement
            Dim curIndex As Integer = 0
            For Each ws As XmlElement In uExcelDoc.FirstChild.ChildNodes
                If ws.Name.ToUpper = "WORKSHEET" Then
                    curIndex = curIndex + 1
                    If curIndex = WorkSheetIndex Then
                        Return ws
                    End If
                End If
            Next
            Return Nothing
        End Function

        Public Function GetWorkSheet(ByVal WorkSheetName As String) As XmlElement
            For Each rootChild As XmlNode In uExcelDoc.ChildNodes
                If rootChild.Name.ToUpper = "WORKBOOK" Then
                    For Each ws As XmlElement In rootChild.ChildNodes
                        If ws.Name.ToUpper = "WORKSHEET" Then
                            If ws.Attributes.GetNamedItem("ss:Name").Value.ToUpper = WorkSheetName.ToUpper Then
                                Return ws
                            End If
                        End If
                    Next
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' Add a row to a worksheet
        ''' </summary>
        ''' <param name="WorkSheetName"></param>
        ''' <param name="RowData">ArrayList of RowDataObject</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function AppendRow(ByVal WorkSheetName As String, ByVal RowData As ArrayList) As XmlElement
            Return AppendRow(GetWorkSheet(WorkSheetName), RowData)
        End Function

        ''' <summary>
        ''' Add a row to a worksheet
        ''' </summary>
        ''' <param name="WorkSheet"></param>
        ''' <param name="RowData">ArrayList of RowDataObject</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function AppendRow(ByVal WorkSheet As XmlElement, ByVal RowData As ArrayList) As XmlElement
            Dim row As XmlElement = uExcelDoc.CreateElement("Row")

            For index As Integer = 0 To RowData.Count - 1
                Dim cell As XmlElement = uExcelDoc.CreateElement("Cell")
                Dim data As XmlElement = uExcelDoc.CreateElement("Data")
                Dim att As XmlAttribute = uExcelDoc.CreateAttribute("ss", "Type", XLMNS_SS)
                Dim rowObj As RowDataObject = CType(RowData(index), RowDataObject)

                Select Case rowObj.DataType
                    Case RowDataObject.DataTypeEnum.DateTime
                        att.Value = "DateTime"
                        data.InnerText = DateTime.Parse(rowObj.ObjectValue.ToString).ToString

                    Case RowDataObject.DataTypeEnum.Number
                        att.Value = "Number"
                        data.InnerText = rowObj.ObjectValue.ToString

                    Case RowDataObject.DataTypeEnum.String
                        att.Value = "String"
                        data.InnerText = rowObj.ObjectValue.ToString
                End Select

                data.Attributes.Append(att)
                cell.AppendChild(data)
                row.AppendChild(cell)
            Next

            ' Append to Table element in Worksheet node
            WorkSheet.FirstChild.AppendChild(row)
            Return row
        End Function

#Region "Private Methods & Functions"

        Private Function initializeExcelDocument(ByVal author As String, ByVal createdDate As DateTime) As XmlDocument
            ' Add xml declaration
            Dim xmlDoc As XmlDocument = New XmlDocument
            Dim dec As XmlDeclaration = xmlDoc.CreateXmlDeclaration("1.0", "", "")
            xmlDoc.AppendChild(dec)

            ' Workbook root node and all namespace attributes
            Dim rootNode As XmlElement = xmlDoc.CreateElement("Workbook")
            Dim att As XmlAttribute = xmlDoc.CreateAttribute("xmlns")
            att.Value = XMLNS_URN
            rootNode.Attributes.Append(att)

            att = xmlDoc.CreateAttribute("xmlns:x")
            att.Value = XLMNS_X
            rootNode.Attributes.Append(att)

            att = xmlDoc.CreateAttribute("xmlns:o")
            att.Value = XLMNS_O
            rootNode.Attributes.Append(att)

            att = xmlDoc.CreateAttribute("xmlns:ss")
            att.Value = XLMNS_SS
            rootNode.Attributes.Append(att)

            att = xmlDoc.CreateAttribute("xmlns:html")
            att.Value = XLMNS_HTML
            rootNode.Attributes.Append(att)

            ' DocumentProperties node
            Dim docProperties As XmlElement = xmlDoc.CreateElement("DocumentProperties")
            att = xmlDoc.CreateAttribute("xmlns")
            att.Value = XLMNS_O
            docProperties.Attributes.Append(att)

            ' Author
            Dim el As XmlElement = xmlDoc.CreateElement("Author")
            el.InnerText = author
            docProperties.AppendChild(el)

            'Last Author
            el = xmlDoc.CreateElement("LastAuthor")
            el.InnerText = author
            docProperties.AppendChild(el)

            Dim strCreatedDate As String = Format(createdDate, "yyyy-mm-dd") & "T" & Format(createdDate, "HH:mm:ss") & "Z"

            ' Last Printed
            el = xmlDoc.CreateElement("LastPrinted")
            el.InnerText = strCreatedDate
            docProperties.AppendChild(el)

            ' Last Saved
            el = xmlDoc.CreateElement("LastSaved")
            el.InnerText = strCreatedDate
            docProperties.AppendChild(el)

            ' Company
            el = xmlDoc.CreateElement("Company")
            docProperties.AppendChild(el)

            ' Version
            el = xmlDoc.CreateElement("Version")
            el.InnerText = "10"
            docProperties.AppendChild(el)

            ' Created Date
            el = xmlDoc.CreateElement("Created")
            el.InnerText = strCreatedDate
            docProperties.AppendChild(el)

            rootNode.AppendChild(docProperties)

            ' Office Settings
            Dim settings As XmlElement = xmlDoc.CreateElement("OfficeDocumentSettings")
            att = xmlDoc.CreateAttribute("xmlns")
            att.Value = XLMNS_O
            settings.Attributes.Append(att)
            rootNode.AppendChild(settings)

            ' Add ExcelWorkbook node
            Dim wkBk As XmlElement = xmlDoc.CreateElement("ExcelWorkbook")
            att = xmlDoc.CreateAttribute("xmlns")
            att.Value = XLMNS_X
            wkBk.Attributes.Append(att)

            Dim child As XmlElement = xmlDoc.CreateElement("ProtectStructure")
            child.InnerText = "False"
            wkBk.AppendChild(child)
            child = xmlDoc.CreateElement("ProtectWindows")
            child.InnerText = "False"
            wkBk.AppendChild(child)
            rootNode.AppendChild(wkBk)

            ' Add styles node
            Dim styles As XmlElement = xmlDoc.CreateElement("Styles")
            rootNode.AppendChild(styles)

            xmlDoc.AppendChild(rootNode)
            Return xmlDoc
        End Function

        Private Sub finalizeExcelDocument()
            ' Count number of rows in each worksheet
            For Each ws As XmlElement In uExcelDoc.ChildNodes(1)
                If ws.Name.ToUpper = "WORKSHEET" Then
                    For Each table As XmlElement In ws.ChildNodes
                        If table.Name.ToUpper = "TABLE" Then
                            ' "child" is Table node
                            Dim numRows As Integer = 0
                            For Each row As XmlElement In table.ChildNodes
                                If row.Name.ToUpper = "ROW" Then
                                    numRows = numRows + 1
                                End If
                            Next
                            ' Update table row count
                            table.SetAttribute("ss:ExpandedRowCount", numRows.ToString)
                            Exit For
                        End If
                    Next
                End If
            Next
        End Sub

#End Region

#Region "Public Properties"

        Public ReadOnly Property ExcelXMLDocument() As XmlDocument
            Get
                finalizeExcelDocument()
                Return uExcelDoc
            End Get
        End Property

#End Region

    End Class

#Region "RowDataObject Class"

    Public Class RowDataObject

        Private uValue As Object
        Private uDataType As DataTypeEnum

        Public Enum DataTypeEnum
            Number
            [String]
            DateTime
        End Enum

        Public Property ObjectValue() As Object
            Get
                Return uValue
            End Get
            Set(ByVal value As Object)
                uValue = value
            End Set
        End Property

        Public Property DataType() As DataTypeEnum
            Get
                Return uDataType
            End Get
            Set(ByVal value As DataTypeEnum)
                uDataType = value
            End Set
        End Property

    End Class

#End Region

End Namespace