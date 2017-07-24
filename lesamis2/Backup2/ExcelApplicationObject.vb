Imports System
Imports System.Data
Imports System.IO
Imports System.Web.ui
Imports Excel

Namespace PogoloUtilities

    Public Class ExcelApplicationObject

        Private uApplication As Excel.Application
        Private uWorkBook As Excel.Workbook
        Private uCurrentWorksheet As Excel.Worksheet
        Private uColumn As String
        Private uRow As Integer
        Private uAlphaNum As Collection

#Region "Methods & Functions"

        Public Function LoadExcelFromFilePath(ByVal filePath As String) As Boolean
            Try
                uApplication = New Excel.Application
                uWorkBook = uApplication.Workbooks.Open(filePath)
                Return True
            Catch
                Return False
            End Try
        End Function

        Public Function getNextRow(ByVal currentRowIndex As Integer, ByVal firstColumn As Integer, ByVal lastColumn As Integer) As ArrayList
            currentRowIndex = currentRowIndex - 1
            Dim row As Excel.Range = CType(uCurrentWorksheet.Rows(currentRowIndex), Excel.Range)
            Dim rowList As New ArrayList
            Dim columnCount As Integer
            For columnCount = firstColumn To lastColumn
                Dim curCell As Range = CType(row.Cells(currentRowIndex, columnCount), Excel.Range)
                rowList.Add(curCell.Value)
            Next
            Return rowList
        End Function

        Public Function getPreviousRow(ByVal currentRowIndex As Integer, ByVal firstColumn As Integer, ByVal lastColumn As Integer) As ArrayList
            currentRowIndex = currentRowIndex - 1
            Dim row As Excel.Range = CType(uCurrentWorksheet.Rows(currentRowIndex), Excel.Range)
            Dim rowList As New ArrayList
            Dim columnCount As Integer
            For columnCount = firstColumn To lastColumn
                Dim curCell As Range = CType(row.Cells(currentRowIndex, columnCount), Excel.Range)
                rowList.Add(curCell.Value)
            Next
            Return rowList
        End Function

        Public Function AllValidRows(ByVal firstColumn As Integer, ByVal lastColumn As Integer) As ArrayList
            Try
                Dim RowList As New ArrayList
                Dim RowNumber As Integer
                For RowNumber = 0 To getLastRowInSheet(1, 1, uCurrentWorksheet)
                    Dim CurrentRowList As New ArrayList
                    CurrentRowList.Add(getNextRow(RowNumber, firstColumn, lastColumn))
                Next
                Return RowList
            Catch
                Return New ArrayList
            End Try
        End Function

        Public Function isValidExcelFile(ByVal file As Web.UI.HtmlControls.HtmlInputFile) As Boolean
            Dim contentType As String = file.PostedFile.ContentType
            If InStr(LCase(contentType), "excel") > 0 Then
                Return True
            Else
                Return False
            End If
        End Function

        Private Function getLastRowInSheet(ByVal startRow As Integer, ByVal startColumn As Integer, ByVal sheet As Worksheet) As Integer
            Dim last As Integer = 0
            Sheet.Select()
            sheet.Range(startColumn, startRow).Select() ' starting point
            Return last
        End Function

        Public Function getSheets() As ArrayList
            Dim sheet As Worksheet
            Dim sheetList As New ArrayList
            For Each sheet In uApplication.ActiveWorkbook.Worksheets
                sheetList.Add(sheet)
            Next
            Return sheetList
        End Function

        Public Function getSheetNames() As ArrayList
            Dim sheet As Worksheet
            Dim sheetNames As New ArrayList
            For Each sheet In uApplication.ActiveWorkbook.Worksheets
                sheetNames.Add(sheet.Name)
            Next
            Return sheetNames
        End Function

        Public Function getSheetByName(ByVal Name As String) As Excel.Worksheet
            Dim sheetList As ArrayList = getSheets()
            Dim sheet As Excel.Worksheet
            For Each sheet In sheetList
                If LCase(Name) = LCase(sheet.Name) Then
                    Return sheet
                End If
            Next
            Return Nothing
        End Function

        Function getFirstSheet() As Excel.Worksheet
            Dim sheetList As ArrayList = getSheets()
            If sheetList.Count > 0 Then
                Return CType(sheetList.Item(0), Excel.Worksheet)
            End If
            Return Nothing
        End Function

        Public Function getWorksheetByName(ByVal name As String, ByVal CreateOnFail As Boolean) As Excel.Worksheet
            Try
                Dim sheet As Excel.Worksheet
                For Each sheet In uWorkBook.Sheets
                    If sheet.Name = name Then Return sheet
                Next
                If CreateOnFail Then
                    Return addSheet(name)
                End If
                Return Nothing
            Catch
                Return Nothing
            End Try
        End Function

        Public Function addSheet(ByVal name As String) As Excel.Worksheet
            Dim sheet As Excel.Worksheet = CType(uWorkBook.Sheets.Add, Excel.Worksheet)
            sheet.Name = name
            Return sheet
        End Function

        Private Sub buildAlphabet()
            If Not uAlphaNum Is Nothing Then Exit Sub
            uAlphaNum = New Collection
            uAlphaNum.Add("a")
            uAlphaNum.Add("b")
            uAlphaNum.Add("c")
            uAlphaNum.Add("d")
            uAlphaNum.Add("e")
            uAlphaNum.Add("f")
            uAlphaNum.Add("g")
            uAlphaNum.Add("h")
            uAlphaNum.Add("i")
            uAlphaNum.Add("j")
            uAlphaNum.Add("k")
            uAlphaNum.Add("l")
            uAlphaNum.Add("m")
            uAlphaNum.Add("n")
            uAlphaNum.Add("o")
            uAlphaNum.Add("p")
            uAlphaNum.Add("q")
            uAlphaNum.Add("r")
            uAlphaNum.Add("s")
            uAlphaNum.Add("t")
            uAlphaNum.Add("u")
            uAlphaNum.Add("v")
            uAlphaNum.Add("w")
            uAlphaNum.Add("x")
            uAlphaNum.Add("y")
            uAlphaNum.Add("z")
        End Sub

        Public Function columnNumbers(ByVal letter As String) As Integer
            Dim i As Integer
            For i = 1 To uAlphaNum.Count
                If CType(uAlphaNum(i), String) = letter Then
                    Return i
                End If
            Next
        End Function

        Public Function convertColumnToInteger(ByVal val As String) As Integer
            If uAlphaNum Is Nothing Then buildAlphabet()
            Dim strLen As Integer = val.Length
            If strLen > 2 Then Exit Function
            Dim total As Integer
            Dim first As String = val.Substring(0, 1)
            Dim factor As Integer = 26

            If strLen = 2 Then
                Dim second As String = val.Substring(1, 1)
                total = (columnNumbers(first) * factor) + columnNumbers(second)
            Else
                total = columnNumbers(first)
            End If
            Return total
        End Function

        Private Function convertIntegerToColumnLetter(ByVal ColumnNumber As Integer) As String
            If uAlphaNum Is Nothing Then buildAlphabet()
            If ColumnNumber > 26 Then
                Dim factor As Double = ColumnNumber / 26
                Dim first As String = CStr(uAlphaNum(factor))
                Dim sec As String = CStr(uAlphaNum(Math.IEEERemainder(ColumnNumber, 26)))
                Return first & sec
            Else
                Return CStr(uAlphaNum(ColumnNumber))
            End If

        End Function

        Public Sub SaveWorkbook(ByVal filePath As String)
            If filePath <> "" Then
                uWorkBook.SaveAs(filePath)
                uWorkBook.Close()
            Else
                Quit()
            End If
        End Sub

        Public Sub Quit()
            Application.Quit()
        End Sub

#End Region

#Region "Properties"

        Public Property Application() As Excel.Application
            Get
                Return uApplication
            End Get
            Set(ByVal Value As Excel.Application)
                uApplication = Value
            End Set
        End Property

        Public Property WorkBook() As Excel.Workbook
            Get
                Return uWorkBook
            End Get
            Set(ByVal Value As Excel.Workbook)
                uWorkBook = Value
            End Set
        End Property

        Public Property CurrentWorksheet() As Excel.Worksheet
            Get
                Return uCurrentWorksheet
            End Get
            Set(ByVal Value As Excel.Worksheet)
                uCurrentWorksheet = Value
            End Set
        End Property

#End Region

    End Class


End Namespace
