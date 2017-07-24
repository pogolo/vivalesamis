Imports System.Data
Imports System.Data.OleDb

Namespace PogoloDataProvider.Data

    Public Class ExcelADO

        Private uExcelFilePath As String
        Private uDataReader As IDataReader
        Private uConnection As OleDbConnection

        ''' <summary>
        ''' Requires a server-bound copy of the spreadsheet.
        ''' </summary>
        ''' <param name="ExcelFilePath">System (not web) path to Excel file.</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal ExcelFilePath As String)
            uExcelFilePath = ExcelFilePath
        End Sub

        ''' <summary>
        ''' Run a SQL-style query on an Excel spreadsheet.
        ''' </summary>
        ''' <param name="ExcelQuery">Such as:<br />"Select * FROM [SKUs$]" - Selects all columns in a specified worksheet
        ''' <br />Select ID,Name FROM RangeName" - Selects specified columns from a previously defined range.
        ''' </param>
        ''' <param name="TargetObjectType"></param>
        ''' <returns>ArrayList of the object type you passed in.</returns>
        ''' <remarks>Uses "Microsoft.Jet.OLEDB.4.0" data provider for Excel access.</remarks>
        Public Function ExecuteQuery(ByVal ExcelQuery As String, ByVal TargetObjectType As Type) As ArrayList
            Try
                Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & uExcelFilePath & ";" & "Extended Properties=Excel 8.0;"
                uConnection = New OleDbConnection(connectionString)
                uConnection.Open()

                Dim selectSQL As New OleDbCommand(ExcelQuery, uConnection)
                Dim adapter As New OleDbDataAdapter(selectSQL)
                Dim dataset As New DataSet()
                adapter.Fill(dataset, "XLData")

                uDataReader = dataset.CreateDataReader
                Dim resultSet As ArrayList = CBO.FillCollection(uDataReader, TargetObjectType)
                Return resultSet
            Finally
                ' Close it all up
                If Not uDataReader Is Nothing Then
                    uDataReader.Close()
                    uDataReader.Dispose()
                    uDataReader = Nothing
                End If
                If Not uConnection Is Nothing Then
                    uConnection.Close()
                    uConnection.Dispose()
                    uConnection = Nothing
                End If
            End Try
        End Function

    End Class

End Namespace