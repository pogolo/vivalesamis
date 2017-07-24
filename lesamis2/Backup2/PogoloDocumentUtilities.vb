Imports System.Web.UI.WebControls
Imports System.Reflection
Imports System.Web.UI
Imports System.Configuration

Namespace PogoloUtilities

    Public Class PogoloDocumentUtilities

        Public Shared Function RemoveSpecialCharacters(ByVal TargetString As String) As String
            Return System.Text.RegularExpressions.Regex.Replace(TargetString, "[^0-9a-zA-Z ]+?", "")
        End Function

        Public Shared Sub AddBackGroundImageToTableCell(ByVal SelectedCell As HtmlControls.HtmlTableCell, ByVal ImagePath As String)
            Try
                SelectedCell.Attributes.Add("Style", "BACKGROUND-POSITION: 0px 0px; BACKGROUND-ATTACHMENT: scroll; BACKGROUND-IMAGE: url(" & ImagePath & "); BACKGROUND-REPEAT: no-repeat")
            Catch ex As Exception

            End Try
        End Sub

        Public Shared Function GetApplicationURL(ByVal CurrentPage As Page) As String
            Try
                Return "http://" & CurrentPage.Request.Url.Host & CurrentPage.Request.ApplicationPath
            Catch
                Return ""
            End Try
        End Function

        Public Shared Function GetSecondLevelDomain(ByVal CurrentPage As Page) As String
            Try
                Dim hostList() As String = CurrentPage.Request.Url.Host.Split(CType(".", Char))
                If hostList.Length > 1 Then
                    System.Array.Reverse(hostList)
                    Return hostList(1) & "." & hostList(0) ' Get last 2 elements
                End If
                Return hostList(0)
            Catch ex As Exception
                Return CurrentPage.Request.Url.Host
            End Try
        End Function

        Public Function GenerateRandomStringUsingGUID(ByVal OutputLength As Integer) As String
            Try
                Dim guidResult As String = System.Guid.NewGuid().ToString()
                guidResult = guidResult.Replace("-", String.Empty)
                If OutputLength > guidResult.Length Then OutputLength = guidResult.Length
                Return guidResult.Substring(0, OutputLength)
            Catch
                Return ""
            End Try
        End Function

        Public Shared Function ReplaceQueryStringValue(ByVal QueryStringKeys As Collections.Specialized.NameValueCollection, ByVal KeyToReplace As String, ByVal NewValue As String) As String
            Try
                ' Replace/add a query string key without losing other keys
                ' Build useful ArrayList
                Dim qsList As New ArrayList
                Dim replaceFlag As Boolean = False
                For i As Integer = 0 To QueryStringKeys.Count - 1
                    Dim curKey As String = QueryStringKeys.AllKeys(i)
                    Dim curValue As String = QueryStringKeys.Item(i)
                    If curKey.ToUpper = KeyToReplace.ToUpper Then
                        curValue = NewValue
                        replaceFlag = True
                    End If
                    Dim namePair As New DictionaryEntry(curKey, curValue)
                    qsList.Add(namePair)
                Next
                ' Add new key value if not found
                If replaceFlag = False Then
                    qsList.Add(New DictionaryEntry(KeyToReplace, NewValue))
                End If
                ' Build the string to return
                Dim qs As String = ""
                For Each namePair As DictionaryEntry In qsList
                    If qs <> "" Then
                        qs += "&" & CType(namePair.Key, String) & "=" & CType(namePair.Value, String)
                    Else
                        qs = CType(namePair.Key, String) & "=" & CType(namePair.Value, String)
                    End If
                Next
                Return qs
            Catch
                Return ""
            End Try
        End Function

        Public Shared Function GetConfigurationSetting(ByVal SettingName As String) As String
            Try
                Return ConfigurationManager.AppSettings(SettingName)
            Catch
                Return ""
            End Try
        End Function

        Public Shared Function GetFullDateString(ByVal SelectedDate As DateTime) As String
            Try
                Return MonthName(SelectedDate.Month) & " " & SelectedDate.Day & ", " & SelectedDate.Year
            Catch
                Return SelectedDate.ToShortDateString
            End Try
        End Function

        Public Shared Function GetQueryStringKeyValue(ByVal CurrentPage As Page, ByVal KeyValue As String) As String
            Try
                Return CurrentPage.Request.QueryString(KeyValue)
            Catch
                Return ""
            End Try
        End Function

        Public Shared Function GetQueryStringIntegerValue(ByVal CurrentPage As Page, ByVal KeyValue As String) As Integer
            Try
                Return CType(CurrentPage.Request.QueryString(KeyValue), Integer)
            Catch
                Return -1
            End Try
        End Function

        Public Shared Function GetQueryStringDoubleValue(ByVal CurrentPage As Page, ByVal KeyValue As String) As Double
            Try
                Return CType(CurrentPage.Request.QueryString(KeyValue), Double)
            Catch
                Return -1
            End Try
        End Function

        Public Shared Function Implode(ByVal ValueList As IList, ByVal Separator As String) As String
            Try
                Dim newList As String = ""
                For Each val As String In ValueList
                    If newList = "" Then
                        newList = val.Trim
                    Else
                        newList = newList & Separator & val.Trim
                    End If
                Next
                Return newList
            Catch
                Return ""
            End Try
        End Function

        Public Shared Function Implode(ByVal Value As String, ByVal ValueDelimiter As String, ByVal NewSeparator As String) As String
            Dim newList() As String = Value.Split(CType(ValueDelimiter, Char))
            Dim newString As String = ""
            Dim ct As Integer = 0
            For Each val As String In newList
                If ct < newList.Length Then
                    newString += val.Trim & NewSeparator
                Else ' Last one
                    newString += val.Trim
                End If
                ct = ct + 1
            Next
            Return newString
        End Function

        Public Shared Function GetDifferenceInDays(ByVal DateToSubtractFrom As DateTime, ByVal DateToSubtract As DateTime) As Long
            Return DateDiff(DateInterval.Day, DateToSubtractFrom, DateToSubtract)
        End Function

        Public Shared Function GetControlFromCell(ByVal ControlName As String, ByVal Cell As System.Web.UI.WebControls.TableCell) As WebControl
            Try
                Return CType(Cell.FindControl(ControlName), WebControl)
            Catch
                Return Nothing
            End Try
        End Function

        Public Shared Function StripHTMLTagsFromString(ByVal StringToStrip As String) As String
            Try
                ' Special characters
                StringToStrip = StringToStrip.Replace("&nbsp;", " ")
                StringToStrip = StringToStrip.Replace("&lt;", "<")
                StringToStrip = StringToStrip.Replace("&gt;", ">")
                StringToStrip = StringToStrip.Replace("&amp;", "&")
                StringToStrip = StringToStrip.Replace("&quot;", "'")
                StringToStrip = StringToStrip.Replace("<br>", vbCrLf)
                StringToStrip = StringToStrip.Replace("<br />", vbCrLf)
                StringToStrip = StringToStrip.Replace("<br/>", vbCrLf)
                StringToStrip = StringToStrip.Replace("<p>", vbCrLf & vbCrLf)
                ' Strips all opening and closing Tags
                Dim myRegEx As New System.Text.RegularExpressions.Regex("<(.|\n)+?>")
                Return myRegEx.Replace(StringToStrip, " ")
            Catch
                Return StringToStrip
            End Try
        End Function

        Public Shared Function ComputeLastDayOfMonth(ByVal TheDate As DateTime) As Integer
            Return DatePart("d", DateAdd("d", -1, DateAdd("m", 1, DateAdd("d", -DatePart("d", TheDate) + 1, TheDate))))
        End Function

        Public Shared Function GetAllParentUserControls(ByVal SelectedControl As Control) As ArrayList
            ' Get all parent User Controls of selected control
            ' Bottom-most parent to top-most
            Try
                Dim controlList As New ArrayList
                Dim parent As Control = SelectedControl.Parent
                Do While Not parent Is Nothing
                    If TypeOf (parent) Is Web.UI.UserControl Or TypeOf (parent) Is System.Web.UI.WebControls.ContentPlaceHolder Or TypeOf (parent) Is MasterPage Then
                        controlList.Add(parent)
                    End If
                    parent = parent.Parent
                Loop
                Return controlList
            Catch
                Return New ArrayList
            End Try
        End Function

        Public Shared Function GetAllParentControlIDsAsString(ByVal SelectedControl As Control, ByVal Separator As String) As String
            Try
                Dim controlList As ArrayList = GetAllParentUserControls(SelectedControl)
                Dim controlID As String = SelectedControl.ID
                For Each ctrl As Control In controlList
                    controlID = ctrl.ID & Separator & controlID
                Next
                Return controlID
            Catch
                Return SelectedControl.ID
            End Try
        End Function

        Public Shared Sub GenerateFocusScript(ByVal curControl As Control, ByVal curPage As Page)
            Try
                Dim controlID As String = GetAllParentControlIDsAsString(curControl, "_")
                Dim newControl As New LiteralControl("<script language='javascript'>setFocusOnControl('" & controlID & "');</script>")
                If TypeOf (curPage) Is System.Web.UI.Page Then
                    Dim myPage As Page = CType(curPage, Page)
                    myPage.Controls.Add(newControl)
                End If
            Catch

            End Try
        End Sub

        Public Shared Sub AddJavaScriptCodeToPage(ByVal CurControl As UserControl, ByVal Code As String)
            Try
                CurControl.Controls.Add(New LiteralControl("<script language='javascript'>" & Code & "</script>"))
            Catch ex As Exception

            End Try
        End Sub

        Public Shared Sub AddJavaScriptCodeToPage(ByVal CurPage As Page, ByVal Code As String)
            Try
                CurPage.Controls.Add(New LiteralControl("<script language='javascript'>" & Code & "</script>"))
            Catch ex As Exception

            End Try
        End Sub

        Public Shared Sub AddDefaultReturnClickAction(ByVal curPage As Page, ByVal TextBoxToLoad As TextBox, ByVal Button As LinkButton)
            Dim jScript As String = "if ((event.which == 13) || (event.keyCode == 13)) {document.forms[0]." & TextBoxToLoad.ID & ".click();} else {return true;}"
            TextBoxToLoad.Attributes.Add("onkeyup", jScript)
        End Sub

        Public Shared Sub AddJavascriptAlertToPage(ByVal CurrentPage As Page, ByVal alertMessage As String)
            Dim JSCode As String = "<script language=javascript>alert('" & alertMessage & "');</script>"
            CurrentPage.Controls.Add(New LiteralControl(JSCode))
        End Sub

        Public Shared Sub AddJavascriptAlertToWebControl(ByVal Ctrl As Web.UI.WebControls.WebControl, ByVal alertMessage As String)
            Ctrl.Attributes.Add("onClick", "javascript:alert('" & alertMessage & "');")
        End Sub

        Public Shared Sub AddJSConfirmToPage(ByVal CurrentPage As Web.UI.Page, ByVal ConfirmMessage As String, ByVal JSCodeToExecute As String)
            If Not CurrentPage.ClientScript.IsStartupScriptRegistered("ConfirmScript") Then
                Dim JSCode As String = "<script language='javascript'>if(confirm('" & ConfirmMessage & "')){" & JSCodeToExecute & "}</script>"
                CurrentPage.ClientScript.RegisterClientScriptBlock(CurrentPage.GetType, "ConfirmScript", JSCode)
            End If
        End Sub

        Public Shared Sub AddJSConfirmToWebControl(ByVal ctrl As Web.UI.WebControls.WebControl, ByVal confirmMessage As String)
            ctrl.Attributes.Add("onClick", "javascript:return confirm('" & confirmMessage & "');")
        End Sub

        Public Shared Sub AddJSConfirmWithCloseToWebControl(ByVal ctrl As Web.UI.WebControls.WebControl, ByVal confirmMessage As String)
            ctrl.Attributes.Add("onClick", "javascript:if(confirm('" & confirmMessage & "')){return self.close();}")
        End Sub

        Public Shared Sub AddJSHelpPopUpToButton(ByVal button As Web.UI.WebControls.WebControl, ByVal pageName As String)
            button.Attributes.Add("onClick", "javascript:return helpPop('" & pageName & "');")
        End Sub

        Public Shared Sub AddJSPopUpWindowToPage(ByVal curPage As Page, ByVal NavigateTo As String, ByVal WindowName As String, ByVal height As Integer, ByVal width As Integer, Optional ByVal ExtraCode As String = "", Optional ByVal WithScroll As Boolean = False)
            Try
                Dim openCode As String = "openPopup('" & NavigateTo & "','" & WindowName & "','" & height & "','" & width & "');"
                If WithScroll Then
                    openCode = "openPopupWithScroll('" & NavigateTo & "','" & WindowName & "','" & height & "','" & width & "');"
                Else
                    openCode = "openPopup('" & NavigateTo & "','" & WindowName & "','" & height & "','" & width & "');"
                End If
                If ExtraCode = "" Then
                    openCode = openCode
                Else
                    ExtraCode = ExtraCode
                End If
                openCode = "<script language='javascript'>" & openCode & "</script>"
                curPage.Controls.Add(New LiteralControl(openCode & ExtraCode))
            Catch

            End Try
        End Sub

        Public Shared Sub AddJSPopUpWindowToButton(ByVal button As Web.UI.WebControls.WebControl, ByVal NavigateTo As String, ByVal WindowName As String, ByVal height As Integer, ByVal width As Integer, Optional ByVal ExtraCode As String = "", Optional ByVal WithScroll As Boolean = False)
            Try
                Dim openCode As String
                If WithScroll Then
                    openCode = "openPopupWithScroll('" & NavigateTo & "','" & WindowName & "','" & height & "','" & width & "');"
                Else
                    openCode = "openPopup('" & NavigateTo & "','" & WindowName & "','" & height & "','" & width & "');"
                End If
                If ExtraCode = "" Then
                    openCode = "return " & openCode
                Else
                    ExtraCode = "return " & ExtraCode
                End If
                button.Attributes.Add("onClick", openCode & ExtraCode)
            Catch

            End Try
        End Sub

        Public Shared Function CurrencyFormat() As System.Globalization.NumberFormatInfo
            Dim numFormat As New System.Globalization.NumberFormatInfo
            numFormat.CurrencyDecimalDigits = 2
            numFormat.CurrencyDecimalSeparator = "."
            numFormat.CurrencySymbol = "$"
            Return numFormat
        End Function

        Public Shared Function PrependZeros(ByVal numZeros As Integer, ByVal strValue As Integer) As String
            Try
                Dim strDocID As String = CType(strValue, String)
                Dim length As Integer = strDocID.Length
                If length >= 7 Then
                    Return strDocID
                Else
                    numZeros = numZeros - length
                    For ct As Integer = 1 To numZeros
                        strDocID = "0" & strDocID
                    Next
                End If
                Return strDocID
            Catch
                Return strValue.ToString
            End Try
        End Function

        Public Shared Function DefaultListItem() As ListItem
            Dim tmpItem As New ListItem
            tmpItem.Text = "< Choose One >"
            tmpItem.Value = "0"
            Return tmpItem
        End Function

        Public Shared Function DefaultListItem(ByVal ListItemText As String) As ListItem
            Dim tmpItem As New ListItem
            tmpItem.Text = "< " & ListItemText & " >"
            tmpItem.Value = "0"
            Return tmpItem
        End Function

        Public Shared Function DefaultListItem(ByVal ListItemText As String, ByVal DefaultValue As String) As ListItem
            Dim tmpItem As New ListItem
            tmpItem.Text = "< " & ListItemText & " >"
            tmpItem.Value = DefaultValue
            Return tmpItem
        End Function

        Public Shared Sub SelectDropDownValue(ByVal DropDown As DropDownList, ByVal StrVal As String)
            Try
                DropDown.ClearSelection()
                Dim item As ListItem
                item = DropDown.Items.FindByValue(StrVal)
                If Not item Is Nothing Then
                    DropDown.SelectedValue = item.Value
                End If
            Catch

            End Try
        End Sub

        Public Shared Sub SelectDropDownText(ByVal DropDown As DropDownList, ByVal StrVal As String)
            Try
                DropDown.ClearSelection()
                Dim item As ListItem
                item = DropDown.Items.FindByText(StrVal)
                If Not item Is Nothing Then
                    DropDown.SelectedValue = item.Value
                End If
            Catch

            End Try
        End Sub

        Public Shared Sub SelectListBoxValue(ByVal curListBox As ListBox, ByVal StrVal As String)
            Try
                curListBox.ClearSelection()
                Dim item As ListItem
                item = curListBox.Items.FindByValue(StrVal)
                If Not item Is Nothing Then
                    curListBox.SelectedValue = item.Value
                End If
            Catch ex As Exception

            End Try
        End Sub

        Public Shared Sub SelectListBoxText(ByVal curListBox As ListBox, ByVal StrVal As String)
            Try
                curListBox.ClearSelection()
                Dim item As ListItem
                item = curListBox.Items.FindByText(StrVal)
                If Not item Is Nothing Then
                    curListBox.SelectedValue = item.Value
                End If
            Catch ex As Exception

            End Try
        End Sub

        Public Shared Sub SelectRadioButtonValue(ByVal rbList As RadioButtonList, ByVal StrVal As String)
            Try
                rbList.ClearSelection()
                Dim item As ListItem = rbList.Items.FindByValue(StrVal)
                If Not item Is Nothing Then
                    rbList.SelectedValue = item.Value
                End If
            Catch ex As Exception

            End Try
        End Sub

        Public Shared Function ConvertBooleanToString(ByVal yesNo As Boolean) As String
            If yesNo Then
                Return "1"
            Else
                Return "0"
            End If
        End Function

        Public Shared Function ConvertStringToBoolean(ByVal StrVal As String) As Boolean
            Try
                If StrVal = "0" Or StrVal.ToUpper = "FALSE" Then
                    Return False
                End If
                If StrVal = "1" Or StrVal.ToUpper = "TRUE" Then
                    Return True
                End If
                Return False
            Catch
                Return False
            End Try
        End Function

        Public Shared Sub LoadStateNames(ByVal DropDown As DropDownList)
            DropDown.Items.Clear()
            Dim stateList As String() = {"Alabama^Alabama", "Alaska^Alaska", "Arkansas^Arkansas", "Arizona^Arizona", _
                  "California^California", "Colorado^Colorado", "Connecticut^Connecticut", "Delaware^Delaware", "Florida^Florida", "Georgia^Georgia", "Hawaii^Hawaii", "Idaho^Idaho", "Illinois^Illinois", _
                  "Indiana^Indiana", "Iowa^Iowa", "Kansas^Kansas", "Kentucky^Kentucky", "Louisiana^Louisiana", "Massachusetts^Massachusetts", "Maine^Maine", "Maryland^Maryland", "Michigan^Michigan", _
                  "Minnesota^Minnesota", "Missouri^Missouri", "Mississippi^Mississippi", "Montana^Montana", "North Carolina^North Carolina", "North Dakota^North Dakota", "Nebraska^Nebraska", "New Hampshire^New Hampshire", "New Jersey^New Jersey", _
                  "New Mexico^New Mexico", "Nevada^Nevada", "New York^New York", "Ohio^Ohio", "Oklahoma^Oklahoma", "Oregon^Oregon", "Pennsylvania^Pennsylvania", "Rhode Island^Rhode Island", "South Carolina^South Carolina", _
                  "South Dakota^South Dakota", "Tennessee^Tennessee", "Texas^Texas", "Utah^Utah", "Virginia^Virginia", "West Virginia^West Virginia", "Vermont^Vermont", "Washington^Washington", "Washington D.C.^Washington D.C.", "Wisconsin^Wisconsin", "Wyoming^Wyoming"}
            Dim state As String
            Dim newItem As ListItem = DefaultListItem()
            DropDown.Items.Add(newItem)
            For Each state In stateList
                Dim stateTuplet() As String = state.Split(CType("^", Char))
                newItem = New ListItem
                newItem.Text = stateTuplet(0)
                newItem.Value = stateTuplet(1)
                DropDown.Items.Add(newItem)
            Next
        End Sub

        Public Shared Sub LoadCanadianProvinces(ByVal DropDown As DropDownList)
            DropDown.Items.Clear()
            Dim stateList As String() = {"Alberta^Alberta", "British Columbia^British Columbia", "Manitoba^Manitoba", "New Brunswick^New Brunswick", "Newfoundland^Newfoundland", "Northwest Territories^Northwest Territories", _
                "Nova Scotia^Nova Scotia", "Nunavut^Nunavut", "Ontario^Ontario", "Prince Edward Island^Prince Edward Island", "Quebec^Quebec", "Saskatchewan^Saskatchewan", "Yukon^Yukon"}
            Dim state As String
            Dim newItem As ListItem = DefaultListItem()
            DropDown.Items.Add(newItem)
            For Each state In stateList
                Dim stateTuplet() As String = state.Split(CType("^", Char))
                newItem = New ListItem
                newItem.Text = stateTuplet(0)
                newItem.Value = stateTuplet(1)
                DropDown.Items.Add(newItem)
            Next
        End Sub

        Public Shared Function TrimComplete(ByVal sValue As String) As String
            Dim sAns As String
            Dim sChar As String
            Dim lLen As Integer
            Dim lCtr As Integer

            sAns = sValue
            lLen = Len(sValue)

            If lLen > 0 Then
                'Ltrim
                For lCtr = 1 To lLen
                    sChar = Mid(sAns, lCtr, 1)
                    If Asc(sChar) > 32 Then Exit For
                Next

                sAns = Mid(sAns, lCtr)
                lLen = Len(sAns)

                'Rtrim
                If lLen > 0 Then
                    For lCtr = lLen To 1 Step -1
                        sChar = Mid(sAns, lCtr, 1)
                        If Asc(sChar) > 32 Then Exit For
                    Next
                End If
                sAns = Left$(sAns, lCtr)
            End If

            TrimComplete = sAns
        End Function

    End Class

End Namespace