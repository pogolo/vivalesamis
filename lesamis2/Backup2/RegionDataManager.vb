Imports Microsoft.VisualBasic

Namespace PogoloUtilities

    Public Class RegionDataManager

        Private uStateXMLPath As String
        Private uCountryXMLPath As String

        Public Sub New()
            ' Default expected path to data files
            Me.uCountryXMLPath = System.Web.HttpContext.Current.Server.MapPath("~/App_Data/Countries.xml")
            Me.uStateXMLPath = System.Web.HttpContext.Current.Server.MapPath("~/App_Data/States.xml")
        End Sub

        Public Sub New(ByVal CountryXMLPath As String, ByVal StateXMLPath As String)
            Me.uCountryXMLPath = CountryXMLPath
            Me.uStateXMLPath = StateXMLPath
        End Sub

        Public Function GetStateList() As ArrayList
            ' Load states from XML document
            Dim xmlStates As System.Xml.XmlDocument = New System.Xml.XmlDocument()
            xmlStates.Load(uStateXMLPath)

            Dim regionList As New ArrayList
            For Each childNode As System.Xml.XmlElement In xmlStates.DocumentElement.ChildNodes
                Dim regionObj As New RegionObject
                regionObj.FullName = childNode.GetAttribute("statename")
                regionObj.Abbreviation = childNode.GetAttribute("abbreviation")
                regionObj.IsDefault = False
                If childNode.HasAttribute("default") Then
                    If childNode.GetAttribute("default").ToUpper = "TRUE" Then
                        regionObj.IsDefault = True
                    End If
                End If
                regionList.Add(regionObj)
            Next

            Return regionList
        End Function

        Public Function GetCountryList() As ArrayList
            ' Load states from XML document
            Dim xmlStates As System.Xml.XmlDocument = New System.Xml.XmlDocument()
            xmlStates.Load(Me.uCountryXMLPath)

            Dim regionList As New ArrayList
            For Each childNode As System.Xml.XmlElement In xmlStates.DocumentElement.ChildNodes
                Dim regionObj As New RegionObject
                regionObj.FullName = childNode.GetAttribute("name")
                regionObj.Abbreviation = childNode.GetAttribute("abbreviation")
                regionObj.IsDefault = False
                If childNode.HasAttribute("default") Then
                    If childNode.GetAttribute("default").ToUpper = "TRUE" Then
                        regionObj.IsDefault = True
                    End If
                End If
                regionList.Add(regionObj)
            Next

            Return regionList
        End Function

        'Private Function sortRegionList(ByVal regionList As ArrayList, ByVal sortDefault As Boolean) As ArrayList
        '    Dim sorter As New PogoloDataProvider.Data.DataSetObject
        '    regionList = sorter.SortObjectArrayList(regionList, "FullName")

        '    ' Sort any/all default regions to top
        '    Dim extractList As New ArrayList
        '    Dim newList As New ArrayList
        '    For Each regionObj As RegionObject In regionList
        '        If regionObj.IsDefault Then
        '            extractList.Add(regionObj)
        '        Else
        '            newList.Add(regionObj)
        '        End If
        '    Next

        '    If extractList.Count > 0 Then
        '        For Each regionObj As RegionObject In extractList
        '            newList.Insert(0, regionObj)
        '        Next
        '    End If

        '    Return newList
        'End Function

    End Class

#Region "RegionObject"

    Public Class RegionObject
        Private uFullName As String
        Private uAbbreviation As String
        Private uIsDefault As Boolean

        Public Property FullName() As String
            Get
                Return uFullName
            End Get
            Set(ByVal value As String)
                uFullName = value
            End Set
        End Property

        Public Property Abbreviation() As String
            Get
                Return uAbbreviation
            End Get
            Set(ByVal value As String)
                uAbbreviation = value
            End Set
        End Property

        Public Property IsDefault() As Boolean
            Get
                Return uIsDefault
            End Get
            Set(ByVal value As Boolean)
                uIsDefault = value
            End Set
        End Property

    End Class

#End Region

End Namespace