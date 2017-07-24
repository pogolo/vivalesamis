Imports System
Imports System.Reflection

Namespace PogoloDataProvider.Data

    Public Class Null

        ' define application encoded null values
        Public Shared ReadOnly Property NullInteger() As Integer
            Get
                Return -1
            End Get
        End Property
        Public Shared ReadOnly Property NullDate() As Date
            Get
                Return Date.MinValue
            End Get
        End Property
        Public Shared ReadOnly Property NullString() As String
            Get
                Return ""
            End Get
        End Property
        Public Shared ReadOnly Property NullBoolean() As Boolean
            Get
                Return False
            End Get
        End Property
        Public Shared ReadOnly Property NullGuid() As Guid
            Get
                Return Guid.Empty
            End Get
        End Property

        Public Shared ReadOnly Property NullByte() As Byte()
            Get
                Return Nothing
            End Get
        End Property

        Public Shared ReadOnly Property NullDecimal() As Decimal
            Get
                Return New Decimal(-1)
            End Get
        End Property

        ' sets a field to an application encoded null value ( used in Presentation layer )
        Public Shared Function SetNull(ByVal objField As Object) As Object
            If Not objField Is Nothing And Not TypeOf (objField) Is System.DBNull Then
                If TypeOf objField Is Integer Then
                    SetNull = NullInteger
                ElseIf TypeOf objField Is Single Then
                    SetNull = NullInteger
                ElseIf TypeOf objField Is Double Then
                    SetNull = NullInteger
                ElseIf TypeOf objField Is Decimal Then
                    SetNull = NullInteger
                ElseIf TypeOf objField Is Date Then
                    SetNull = NullDate
                ElseIf TypeOf objField Is String Then
                    SetNull = NullString
                ElseIf TypeOf objField Is Boolean Then
                    SetNull = NullBoolean
                ElseIf TypeOf objField Is Guid Then
                    SetNull = NullGuid
                ElseIf TypeOf objField Is Byte() Then
                    SetNull = NullByte
                Else
                    Throw New NullReferenceException
                End If
            Else ' assume string
                SetNull = NullString
            End If
        End Function

        Public Enum DataTypes
            [Integer]
            [Double]
            [Single]
            [Decimal]
            [Date]
            [Boolean]
            [String]
            GUID
            [Byte]
        End Enum

        Private Shared Function convertDataType(ByVal dataType As Integer) As Object
            Select Case dataType
                Case DataTypes.Integer
                    Return NullInteger
                Case DataTypes.Single
                    Return NullInteger
                Case DataTypes.Double
                    Return NullInteger
                Case DataTypes.Decimal
                    Return NullInteger
                Case DataTypes.Date
                    Return NullDate
                Case DataTypes.String
                    Return NullString
                Case DataTypes.Boolean
                    Return NullBoolean
                Case DataTypes.GUID
                    Return NullGuid
                Case DataTypes.Byte
                    Return NullByte
                Case Else
                    Throw New NullReferenceException
            End Select
        End Function

        Public Shared Function SetNull(ByVal objField As Object, ByVal dataType As Integer) As Object
            If Not objField Is Nothing Then
                If TypeOf (objField) Is System.DBNull Then
                    Return convertDataType(dataType)
                Else ' assume string
                    Return objField
                End If
            End If
            Return Nothing
        End Function

        ' sets a field to an application encoded null value ( used in BLL layer )
        Public Shared Function SetNull(ByVal objPropertyInfo As PropertyInfo) As Object
            Select Case objPropertyInfo.PropertyType.ToString
                Case "System.Int16", "System.Int32", "System.Int64", "System.Single", "System.Double"
                    SetNull = NullInteger
                Case "System.Decimal"
                    SetNull = NullDecimal
                Case "System.DateTime"
                    SetNull = NullDate
                Case "System.String", "System.Char"
                    SetNull = NullString
                Case "System.Boolean"
                    SetNull = NullBoolean
                Case "System.Guid"
                    SetNull = NullGuid
                Case "System.Byte[]"
                    SetNull = NullByte
                Case Else
                    ' Enumerations default to the first entry
                    Dim pType As Type = objPropertyInfo.PropertyType
                    If pType.BaseType.Equals(GetType(System.Enum)) Then
                        Dim objEnumValues As System.Array = System.Enum.GetValues(pType)
                        Array.Sort(objEnumValues)
                        SetNull = System.Enum.ToObject(pType, objEnumValues.GetValue(0))
                    Else
                        Throw New NullReferenceException
                    End If
            End Select
        End Function

        Public Shared Function GetNull(ByVal objField As Object) As Object
            Return Null.GetNull(objField, DBNull.Value)
        End Function

        ' convert an application encoded null value to a database null value ( used in DAL )
        Public Shared Function GetNull(ByVal objField As Object, ByVal objDBNull As Object) As Object
            GetNull = objField
            If objField Is Nothing Or TypeOf (objField) Is DBNull Then
                GetNull = objDBNull
            ElseIf TypeOf objField Is Integer Then
                If Convert.ToInt32(objField) = NullInteger Then
                    GetNull = objDBNull
                End If
            ElseIf TypeOf objField Is Single Then
                If Convert.ToSingle(objField) = NullInteger Then
                    GetNull = objDBNull
                End If
            ElseIf TypeOf objField Is Double Then
                If Convert.ToDouble(objField) = NullInteger Then
                    GetNull = objDBNull
                End If
            ElseIf TypeOf objField Is Decimal Then
                If Convert.ToDecimal(objField) = NullInteger Then
                    GetNull = objDBNull
                End If
            ElseIf TypeOf objField Is Date Then
                If Convert.ToDateTime(objField) = NullDate Then
                    GetNull = objDBNull
                End If
            ElseIf TypeOf objField Is String Then
                If objField Is Nothing Then
                    GetNull = objDBNull
                Else
                    If objField.ToString = NullString Then
                        GetNull = objDBNull
                    End If
                End If
            ElseIf TypeOf objField Is Boolean Then
                If Convert.ToBoolean(objField) = NullBoolean Then
                    GetNull = objDBNull
                End If
            ElseIf TypeOf objField Is Guid Then
                If CType(objField, System.Guid).Equals(NullGuid) Then
                    GetNull = objDBNull
                End If
            ElseIf TypeOf objField Is System.Array Then
                If objField Is Nothing Then
                    GetNull = objDBNull
                End If
            Else
                Throw New NullReferenceException
            End If
        End Function

        ' checks if a field contains an application encoded null value
        Public Shared Function IsNull(ByVal objField As Object) As Boolean
            If objField.Equals(SetNull(objField)) Then
                IsNull = True
            Else
                IsNull = False
            End If
        End Function

    End Class

End Namespace