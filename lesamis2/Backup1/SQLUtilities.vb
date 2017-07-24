Imports System.Data
Imports System.Collections
Imports System.Configuration
Imports Microsoft.ApplicationBlocks.Data

Namespace PogoloDataProvider.Data

    Public Class SQLUtilities

        Private uSqlConnect As SqlClient.SqlConnection
        Private uConnectionString As String

        Public Sub New()
            Try
                uConnectionString = ConfigurationManager.AppSettings.Get("SQLConnectionString")
                uSqlConnect = New SqlClient.SqlConnection(uConnectionString)
            Catch ex As Exception
                Throw ex
            End Try
        End Sub

        Public Sub New(ByVal AppConfigKey As String)
            Try
                uConnectionString = ConfigurationManager.AppSettings.Get(AppConfigKey)
                uSqlConnect = New SqlClient.SqlConnection(uConnectionString)
            Catch ex As Exception
                Throw ex
            End Try
        End Sub

        Public Sub New(ByVal ConnectionStringName As String, ByVal GetFromConnectionStrings As Boolean)
            If GetFromConnectionStrings Then
                If ConnectionStringName <> "" Then
                    uConnectionString = ConfigurationManager.ConnectionStrings(ConnectionStringName).ConnectionString
                Else
                    uConnectionString = ConfigurationManager.ConnectionStrings("LocalSqlServer").ConnectionString
                End If
            Else
                uConnectionString = ConfigurationManager.AppSettings.Get(ConnectionStringName)
            End If
                uSqlConnect = New SqlClient.SqlConnection(uConnectionString)
        End Sub

        Private Function getTableValue(ByVal reader As SqlClient.SqlDataReader, ByVal FieldName As String) As Object
            Dim columnNumber As Integer
            While reader.Read
                For columnNumber = 0 To reader.FieldCount - 1
                    If UCase(reader.GetName(columnNumber)) = UCase(FieldName) Then
                        Return reader.Item(columnNumber)
                    End If
                Next
            End While
            Return Nothing
        End Function

        Public Function ExecuteCommandString(ByVal queryString As String) As Integer
            Try
                Dim command As New SqlClient.SqlCommand(queryString, uSqlConnect)
                command.Connection = uSqlConnect
                uSqlConnect.Open()
                Dim obj As Object = command.ExecuteScalar()
                command.Connection.Close()
                If Not obj Is Nothing Then
                    Return CType(obj, Integer)
                Else
                    Return -1
                End If
            Catch
                Return -1
            End Try
        End Function

        Public Function ExecuteStoredProcedure(ByVal spName As String, ByVal valueList As ArrayList) As Integer
            Try
                Dim obj As Object = SqlHelper.ExecuteScalar(uConnectionString, spName, valueList)
                If Not obj Is Nothing Then
                    Return CType(obj, Integer)
                Else
                    Return -1
                End If
            Catch
                Return -1
            End Try
        End Function

        Public Function ExecuteQueryString(ByVal queryString As String) As ArrayList
            Try
                Dim command As New SqlClient.SqlCommand(queryString, uSqlConnect)
                command.Connection = uSqlConnect
                sqlConnect.Open()
                Dim reader As SqlClient.SqlDataReader = command.ExecuteReader
                Dim recordList As New ArrayList
                While reader.Read
                    Dim fieldValueList As New ArrayList
                    For columnNumber As Integer = 0 To reader.FieldCount - 1
                        Dim colName As String = reader.GetName(columnNumber)
                        Dim dic As New DictionaryEntry
                        dic.Key = colName
                        dic.Value = reader.GetValue(columnNumber)
                        fieldValueList.Add(dic)
                    Next
                    recordList.Add(fieldValueList)
                End While
                sqlConnect.Close()
                Return recordList
            Catch
                Return New ArrayList
            End Try
        End Function

#Region "Properties"

        Public ReadOnly Property ConnectionString() As String
            Get
                Return uConnectionString
            End Get
        End Property

        Public Property sqlConnect() As SqlClient.SqlConnection
            Get
                Return uSqlConnect
            End Get
            Set(ByVal Value As SqlClient.SqlConnection)
                uSqlConnect = Value
            End Set
        End Property

        Public ReadOnly Property ServerName() As String
            Get
                Return uSqlConnect.DataSource
            End Get
        End Property

#End Region

    End Class

End Namespace