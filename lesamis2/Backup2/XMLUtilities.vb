Imports System.Xml

Namespace PogoloUtilities

    Public Class XMLUtilities

        ''' <summary>
        ''' Removes special unicode characters like 'smart quotes', etc.
        ''' http://www.roubaixinteractive.com/PlayGround/Binary_Conversion/The_Characters.asp
        ''' </summary>
        ''' <param name="StringValue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function RemoveInvalidXMLCharacters(ByVal StringValue As String) As String
            Try
                StringValue = StringValue.Replace(Chr(147), """")
                StringValue = StringValue.Replace(Chr(145), "`")
                StringValue = StringValue.Replace(Chr(146), "'")
                Return StringValue.Replace(Chr(148), """")
            Catch ex As Exception
                Return StringValue
            End Try
        End Function

        Public Shared Function ConvertXMLDocumentToString(ByVal XMLDoc As XmlDocument) As String
            Try
                Dim sw As New System.IO.StringWriter
                Dim xw As New XmlTextWriter(sw)
                XMLDoc.WriteTo(xw)
                Return sw.ToString()
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Public Shared Function ConvertStringToXMLDocument(ByVal XMLString As String) As XmlDocument
            Try
                Dim xmlDoc As New XmlDocument
                xmlDoc.LoadXml(XMLString)
                Return xmlDoc
            Catch ex As Exception
                Return New XmlDocument
            End Try
        End Function

        Public Shared Function GetNodeValueByName(ByVal Parent As XmlNode, ByVal NodeName As String) As String
            Try
                If Not Parent Is Nothing Then
                    Return Parent.SelectSingleNode(NodeName).InnerText()
                End If
                Return ""
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Public Shared Function GetNodeValueByName(ByVal XMLDoc As XmlDocument, ByVal NodeName As String) As String
            Try
                Dim nodeList As XmlNodeList = XMLDoc.GetElementsByTagName(NodeName)
                If nodeList.Count > 0 Then
                    Return nodeList.Item(0).InnerText
                End If
                Return ""
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Public Shared Function GetNodeValueByNameAsDouble(ByVal Parent As XmlNode, ByVal NodeName As String) As Double
            Try
                Return CType(GetNodeValueByName(Parent, NodeName), Double)
            Catch ex As Exception
                Return 0
            End Try
        End Function

        Public Shared Function GetNodeValueByNameAsDouble(ByVal XMLDoc As XmlDocument, ByVal NodeName As String) As Double
            Try
                Return CType(GetNodeValueByName(XMLDoc, NodeName), Double)
            Catch ex As Exception
                Return 0
            End Try
        End Function

        Public Shared Function GetNodeValueByNameAsInteger(ByVal Parent As XmlNode, ByVal NodeName As String) As Integer
            Try
                Return CType(GetNodeValueByName(Parent, NodeName), Integer)
            Catch ex As Exception
                Return 0
            End Try
        End Function

        Public Shared Function GetNodeValueByNameAsInteger(ByVal XMLDoc As XmlDocument, ByVal NodeName As String) As Integer
            Try
                Return CType(GetNodeValueByName(XMLDoc, NodeName), Integer)
            Catch ex As Exception
                Return 0
            End Try
        End Function

    End Class

End Namespace