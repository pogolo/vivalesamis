Imports System.Reflection
Imports System.Collections

Namespace PogoloDataProvider.Data

    Public Class DataSetObject

#Region "Methods & Subs"

        Public Function SortObjectArrayList(ByVal Col As ArrayList, ByVal PropertyName As String, Optional ByVal BlnCompareNumeric As Boolean = False) As ArrayList
            Try
                ' Sorts any arraylist by the property name you specify
                Dim colNew As New ArrayList
                Dim objCurrent As Object
                Dim objCompare As Object
                Dim lngCompareIndex As Long
                Dim strCurrent As Object
                Dim strCompare As Object
                Dim blnGreaterValueFound As Boolean

                For Each objCurrent In Col
                    'get value of current item...
                    strCurrent = CallByName(objCurrent, PropertyName, vbGet)
                    'setup for compare loop
                    blnGreaterValueFound = False
                    lngCompareIndex = 0
                    For Each objCompare In colNew
                        Dim compareInfo As PropertyInfo = objCompare.GetType().GetProperty(PropertyName)
                        strCompare = CallByName(objCompare, PropertyName, vbGet)

                        ' Short
                        If strCompare.GetType Is GetType(Short) Then
                            If Short.Parse(strCompare.ToString()).CompareTo(Short.Parse(strCurrent.ToString())) > 0 Then
                                blnGreaterValueFound = True
                            End If

                            ' Integer 
                        ElseIf strCompare.GetType Is GetType(Integer) Then
                            If Integer.Parse(strCompare.ToString()).CompareTo(Integer.Parse(strCurrent.ToString())) > 0 Then
                                blnGreaterValueFound = True
                            End If

                            ' Double
                        ElseIf strCompare.GetType Is GetType(Double) Then
                            If Double.Parse(strCompare.ToString()).CompareTo(Double.Parse(strCurrent.ToString())) > 0 Then
                                blnGreaterValueFound = True
                            End If

                            ' Decimal
                        ElseIf strCompare.GetType Is GetType(Decimal) Then
                            If Decimal.Parse(strCompare.ToString()).CompareTo(Decimal.Parse(strCurrent.ToString())) > 0 Then
                                blnGreaterValueFound = True
                            End If

                            ' DateTime
                        ElseIf strCompare.GetType Is GetType(DateTime) Then
                            If DateTime.Parse(strCompare.ToString()).CompareTo(DateTime.Parse(strCurrent.ToString())) > 0 Then
                                blnGreaterValueFound = True
                            End If

                        Else
                            ' Convert to string
                            If strCompare.ToString().CompareTo(strCurrent.ToString()) > 0 Then
                                blnGreaterValueFound = True
                            End If
                        End If

                        ' Insert in arraylist
                        If blnGreaterValueFound Then
                            colNew.Insert(lngCompareIndex, objCurrent)
                            Exit For
                        End If
                        lngCompareIndex = lngCompareIndex + 1
                    Next
                    ' If we didn't find something bigger, just add it to the end of the new collection...
                    If blnGreaterValueFound = False Then
                        colNew.Add(objCurrent)
                    End If
                Next

                uDataList = colNew
                Return colNew
            Catch
                uDataList = Col
                Return Col
            End Try

        End Function

#End Region

#Region "Properties"

        Private uDataList As ArrayList

        Public Property DataList() As ArrayList
            Get
                Return uDataList
            End Get
            Set(ByVal Value As ArrayList)
                uDataList = Value
            End Set
        End Property

#End Region

    End Class

End Namespace
