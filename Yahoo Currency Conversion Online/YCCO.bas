Attribute VB_Name = "YahooCurrencyConversionOnline"
Option Explicit

Public Enum YCCO_REFRESH_INTERVAL
   YCCO_EVERY_DAY = 0
   YCCO_EVERY_HOUR
   YCCO_EVERY_MINUTE
   YCCO_EVERY_MONTH
   YCCO_EVERY_TIME
End Enum

Public Enum YCCO_CONVERT_RESPONSE
   YCCO_SUCCEEDED
   YCCO_FAILED
   YCCO_NO_DATA_AVAILABLE
   YCCO_BUSY_PROCESSING
   YCCO_INVALID_PARAMETER
End Enum

Private YCCO_internet_transfer_control As Inet
Private YCCO_internet_transfer_control_defined As Boolean

Private YCCO_source_currency_control As Control
Private YCCO_destination_currency_control As Control
Private YCCO_currency_controls_defined As Boolean

Private YCCO_internet_refresh_interval As YCCO_REFRESH_INTERVAL

Private YCCO_registry_usage_enabled As Boolean

Private YCCO_last_search_string$
Private YCCO_last_registry_record$

Private YCCO_is_busy As Boolean

Private Function GetInternetSiteContent(URL_path As String) As String
   Dim website_data As String
   
   On Error GoTo ERROR_GetInternetSiteContent
   
   GetInternetSiteContent = ""
   
   YCCO_internet_transfer_control.Cancel
   YCCO_internet_transfer_control.Protocol = icHTTP
   YCCO_internet_transfer_control.URL = URL_path
   YCCO_internet_transfer_control.RequestTimeout = 5000
   
   website_data = YCCO_internet_transfer_control.OpenURL(, icString)

   GetInternetSiteContent = website_data

EXIT_GetInternetSiteContent:
   
   Exit Function
   
ERROR_GetInternetSiteContent:
   Select Case MsgBox("Unable to download data Yahoo website!" & vbCrLf & "[ " & CStr(Err.Number) & " - " & Err.Description & " ]", vbRetryCancel + vbCritical + vbApplicationModal, "YCCO Error!")
      Case vbCancel
         Resume EXIT_GetInternetSiteContent
      Case vbRetry
         Resume
   End Select
   
   Resume EXIT_GetInternetSiteContent
   
End Function

Private Function ExtractCurrencyData(ByVal data_string$, ByVal data_pos&, ByRef currency_alias$, ByRef currency_title$) As Boolean
   Dim marker_pos&, title_string$
   
   On Error Resume Next
   
   ExtractCurrencyData = False

   currency_alias = Trim(Mid$(data_string, data_pos, 3))

   If currency_alias = "" Then Exit Function
   
   title_string = Mid$(data_string, data_pos + 4)
   
   If title_string = "" Then Exit Function
   
   marker_pos = InStr(title_string, "<")
   
   If marker_pos = 0 Then Exit Function
   
   currency_title = Trim(Mid$(data_string, data_pos + 4, marker_pos - 1))
   
   If currency_title = "" Then Exit Function

   ExtractCurrencyData = True

End Function

Private Function GetRegistryDateString() As String
   GetRegistryDateString = Format(Date, "yyyy") & Format(Date, "mm") & Format(Date, "dd") & Format(Time, "hh") & Format(Time, "nn")
End Function

Private Function GetConversionSearchString() As String
   Dim source_currency_alias$, destination_currency_alias$
   
   On Error Resume Next
   
   GetConversionSearchString = ""
   
   source_currency_alias = Mid$(YCCO_source_currency_control.Tag, YCCO_source_currency_control.ItemData(YCCO_source_currency_control.ListIndex), 3)
   destination_currency_alias = Mid$(YCCO_destination_currency_control.Tag, YCCO_destination_currency_control.ItemData(YCCO_destination_currency_control.ListIndex), 3)
   
   If source_currency_alias <> destination_currency_alias Then
      GetConversionSearchString = source_currency_alias & destination_currency_alias
   Else
      GetConversionSearchString = "------"
   End If

End Function

Private Function GetInvertedConversionSearchString() As String
   Dim source_currency_alias$, destination_currency_alias$
   
   On Error Resume Next
   
   GetInvertedConversionSearchString = ""
   
   source_currency_alias = Mid$(YCCO_destination_currency_control.Tag, YCCO_destination_currency_control.ItemData(YCCO_destination_currency_control.ListIndex), 3)
   destination_currency_alias = Mid$(YCCO_source_currency_control.Tag, YCCO_source_currency_control.ItemData(YCCO_source_currency_control.ListIndex), 3)
   
   If source_currency_alias <> destination_currency_alias Then
      GetInvertedConversionSearchString = source_currency_alias & destination_currency_alias
   Else
      GetInvertedConversionSearchString = "------"
   End If

End Function

Private Function GetConversionRatio(ByVal data_string$, Optional return_conversion_ratio) As Boolean
   Dim search_string$, marker_pos&, value_string$, conversion_ratio!
   
   On Error Resume Next

   GetConversionRatio = False
   
   search_string = GetConversionSearchString
                   
   If search_string = "------" Then
      conversion_ratio = 1!
   Else
      marker_pos = InStr(data_string, ">" & search_string & "=")
      
      If marker_pos = 0 Then Exit Function
   
      marker_pos = InStr(marker_pos + 1, data_string, "<td>")
      
      If marker_pos = 0 Then Exit Function
   
      marker_pos = InStr(marker_pos + 1, data_string, "<td>")
      
      If marker_pos = 0 Then Exit Function
   
      value_string = Mid$(data_string, marker_pos + Len("<td><b>"))
      
      If value_string = "" Then Exit Function
      
      marker_pos = InStr(value_string, "</b>")
      
      If marker_pos = 0 Then Exit Function
      
      conversion_ratio = CCur(Left$(value_string, marker_pos - 1))
   
      If conversion_ratio < 0! Then Exit Function
      
      YCCO_last_search_string = search_string
      YCCO_last_registry_record = GetRegistryDateString & "-" & CStr(conversion_ratio)
      
      If YCCO_registry_usage_enabled = True Then
         SaveSetting "YCCO", "ConversionTable", search_string, YCCO_last_registry_record
      End If
   End If
   
   If IsMissing(return_conversion_ratio) = False Then
      return_conversion_ratio = conversion_ratio
    End If
   
   GetConversionRatio = True
End Function

'--------------------------------------------------------------------------------------------------------------------------
'   Global functions - called from Forms etc to utilise the YCCO module's functionality
'--------------------------------------------------------------------------------------------------------------------------

Public Sub YCCO_RegisterInternetTransferControl(internet_transfer_control As Inet)
   
   On Error Resume Next
   
   Err.Clear
   Set YCCO_internet_transfer_control = internet_transfer_control
   If Err.Number Then
      On Error GoTo 0
      Err.Raise vbObjectError + 1, "YCCO", "Unable to register the Internet Transfer control!"
   End If
   
   YCCO_internet_transfer_control_defined = True
End Sub

Public Sub YCCO_RegisterCurrencySelectionControls(source_currency_control As Control, destination_currency_control As Control)
   
   On Error Resume Next
   
   If TypeOf source_currency_control Is ComboBox Or TypeOf source_currency_control Is ListBox Then
      Err.Clear
      Set YCCO_source_currency_control = source_currency_control
      If Err.Number Then
         On Error GoTo 0
         Err.Raise vbObjectError + 2, "YCCO", "Unable to register the source currency selection " & TypeName(source_currency_control) & " control!"
      End If
      
      If TypeOf YCCO_source_currency_control Is ComboBox Then
         If YCCO_source_currency_control.Style <> vbComboDropdownList Then
            On Error GoTo 0
            Err.Raise vbObjectError + 3, "YCCO", "The source currency ComboBox must have its 'Style' property set to 'vbComboDropdownList'!"
         End If
      End If
   Else
         On Error GoTo 0
         Err.Raise vbObjectError + 4, "YCCO", "Unable to register the source currency control!" & vbCrLf & "Must be either a ComboBox or a ListBox control."
   End If
      
   If TypeOf destination_currency_control Is ComboBox Or TypeOf destination_currency_control Is ListBox Then
      Err.Clear
      Set YCCO_destination_currency_control = destination_currency_control
      If Err.Number Then
         On Error GoTo 0
         Err.Raise vbObjectError + 5, "YCCO", "Unable to register the destination currency  " & TypeName(source_currency_control) & " control!"
      End If
      
      If TypeOf YCCO_source_currency_control Is ComboBox Then
         If YCCO_destination_currency_control.Style <> vbComboDropdownList Then
            On Error GoTo 0
            Err.Raise vbObjectError + 6, "YCCO", "The destination currency combobox must have its 'Style' property set to 'vbComboDropdownList'!"
          End If
      End If
      
      YCCO_currency_controls_defined = True
   Else
         On Error GoTo 0
         Err.Raise vbObjectError + 7, "YCCO", "Unable to register the destination currency control!" & vbCrLf & "Must be either a ComboBox or a ListBox"
   End If
End Sub

Public Function YCCO_PopulateCurrencySelectionControls() As Boolean
   Dim data_string$, text_pos&, end_text_pos&, currency_alias$, currency_title$
   
   On Error Resume Next

   If YCCO_is_busy = True Then
      YCCO_PopulateCurrencySelectionControls = True
      Exit Function
   End If

   YCCO_is_busy = True
   YCCO_PopulateCurrencySelectionControls = False
   
   If YCCO_internet_transfer_control_defined = False Then
      On Error GoTo 0
      Err.Raise vbObjectError + 10, "YCCO", "Call YCCO_RegisterInternetTransferControl() before YCCO_PopulateCurrencySelectionControls()!"
   End If
   If YCCO_currency_controls_defined = False Then
      On Error GoTo 0
      Err.Raise vbObjectError + 11, "YCCO", "Call YCCO_RegisterCurrencySelectionControls() before YCCO_PopulateCurrencySelectionControls()!"
   End If

   data_string = GetInternetSiteContent("http://uk.finance.yahoo.com/m5?a=1&s=GBP&t=USD")
   
   If data_string = "" Then
      YCCO_is_busy = False
      Exit Function
   End If
   
   text_pos = InStr(data_string, "<select name=s>")
   
   If text_pos = 0 Then
      YCCO_is_busy = False
      Exit Function
   End If
   
   data_string = Mid$(data_string, text_pos + Len("<select name=s>"))
   
   end_text_pos = InStr(data_string, "<select name=t>")
   
   If end_text_pos = 0 Then
      YCCO_is_busy = False
      Exit Function
   End If
   
   text_pos = InStr(data_string, "<option selected value=")
   
   If text_pos = 0 Or text_pos > end_text_pos Then
      YCCO_is_busy = False
      Exit Function
   End If
   
   If ExtractCurrencyData(data_string, text_pos + Len("<option selected value="), currency_alias, currency_title) = False Then
      YCCO_is_busy = False
      Exit Function
   End If
   
   YCCO_source_currency_control.Clear
   YCCO_source_currency_control.Tag = ""
   
   YCCO_source_currency_control.AddItem currency_title
   YCCO_source_currency_control.ItemData(YCCO_source_currency_control.NewIndex) = Len(YCCO_source_currency_control.Tag) + 1
   YCCO_source_currency_control.Tag = YCCO_source_currency_control.Tag & currency_alias

   text_pos = InStr(data_string, "<option value=")
   
   Do While text_pos > 0 And text_pos < end_text_pos
      If ExtractCurrencyData(data_string, text_pos + Len("<option value="), currency_alias, currency_title) = False Then
         YCCO_is_busy = False
         Exit Function
      End If
      
      YCCO_source_currency_control.AddItem currency_title
      YCCO_source_currency_control.ItemData(YCCO_source_currency_control.NewIndex) = Len(YCCO_source_currency_control.Tag) + 1
      YCCO_source_currency_control.Tag = YCCO_source_currency_control.Tag & currency_alias
      
      text_pos = InStr(text_pos + 1, data_string, "<option value=")
   Loop

   data_string = Mid$(data_string, end_text_pos + Len("<select name=t>"))
   
   end_text_pos = InStr(data_string, "</select>")
   
   If end_text_pos = 0 Then
      YCCO_is_busy = False
      Exit Function
   End If
   
   text_pos = InStr(data_string, "<option selected value=")
   
   If text_pos = 0 Or text_pos > end_text_pos Then
      YCCO_is_busy = False
      Exit Function
   End If
   
   If ExtractCurrencyData(data_string, text_pos + Len("<option selected value="), currency_alias, currency_title) = False Then
      YCCO_is_busy = False
      Exit Function
   End If
   
   YCCO_destination_currency_control.Clear
   YCCO_destination_currency_control.Tag = ""
   
   YCCO_destination_currency_control.AddItem currency_title
   YCCO_destination_currency_control.ItemData(YCCO_destination_currency_control.NewIndex) = Len(YCCO_destination_currency_control.Tag) + 1
   YCCO_destination_currency_control.Tag = YCCO_destination_currency_control.Tag & currency_alias

   text_pos = InStr(data_string, "<option value=")
   
   Do While text_pos > 0 And text_pos < end_text_pos
      If ExtractCurrencyData(data_string, text_pos + Len("<option value="), currency_alias, currency_title) = False Then
         YCCO_is_busy = False
         Exit Function
      End If
      
      YCCO_destination_currency_control.AddItem currency_title
      YCCO_destination_currency_control.ItemData(YCCO_destination_currency_control.NewIndex) = Len(YCCO_destination_currency_control.Tag) + 1
      YCCO_destination_currency_control.Tag = YCCO_destination_currency_control.Tag & currency_alias
      
      text_pos = InStr(text_pos + 1, data_string, "<option value=")
   Loop

   YCCO_source_currency_control.ListIndex = 0
   YCCO_destination_currency_control.ListIndex = 0
   
   GetConversionRatio data_string

   YCCO_PopulateCurrencySelectionControls = True
      
   YCCO_is_busy = False
End Function

Public Sub YCCO_SetInternetRefreshInterval(ByVal refresh_interval As YCCO_REFRESH_INTERVAL)
   YCCO_internet_refresh_interval = refresh_interval
End Sub

Public Sub YCCO_EnableRegistryCache(ByVal chache_enabled As Boolean)
   YCCO_registry_usage_enabled = chache_enabled
End Sub

Public Function YCCO_ConvertGeneric(source_currency_object, destination_currency_object) As Integer
   Dim source_currency_units!, destination_currency_units!, retval%
   Dim conversion_required As Boolean
   
   On Error Resume Next
   
   YCCO_ConvertGeneric = YCCO_INVALID_PARAMETER
   conversion_required = True
   
   Select Case TypeName(source_currency_object)
      Case "Byte", "Integer", "Long", "Single", "Double", "Currency"
         Err.Clear
         source_currency_units = CSng(source_currency_object)
         If Err.Number Then
            Exit Function
         End If
         If source_currency_units < 0! Then
            Exit Function
         ElseIf source_currency_units = 0! Then
            destination_currency_units = 0!
            conversion_required = False
         End If
   
      Case "String"
         Err.Clear
         source_currency_units = CCur(Trim(source_currency_object))
         If Err.Number Then
            Exit Function
         End If
         If source_currency_units < 0! Then
            Exit Function
         ElseIf source_currency_units = 0! Then
            destination_currency_units = 0!
            conversion_required = False
         End If
   
      Case "Empty"   ' unassigned variant
         destination_currency_units = 0!
         conversion_required = False
      
      Case "TextBox"
         Err.Clear
         source_currency_units = CCur(Trim(source_currency_object.Text))
         If Err.Number Then
            Exit Function
         End If
         If source_currency_units < 0! Then
            Exit Function
         ElseIf source_currency_units = 0! Then
            source_currency_object.Text = "0"
            destination_currency_units = 0!
            conversion_required = False
         End If
   
      Case "ComboBox", "ListBox"
         Err.Clear
         source_currency_units = CCur(Trim(source_currency_object.List(source_currency_object.ListIndex)))
         If Err.Number Then
            Exit Function
         End If
         If source_currency_units <= 0! Then
            Exit Function
         End If
   End Select

   If conversion_required = True Then
      On Error GoTo 0
      retval = YCCO_Convert(source_currency_units, destination_currency_units)
      On Error Resume Next
      Select Case retval
         Case YCCO_BUSY_PROCESSING, YCCO_FAILED
            YCCO_ConvertGeneric = retval
            Exit Function
      End Select
   End If

   Select Case TypeName(destination_currency_object)
      Case "Byte"
         Err.Clear
         destination_currency_object = CByte(destination_currency_units)
         If Err.Number Then
            destination_currency_object = 0
            Exit Function
         End If
         YCCO_ConvertGeneric = retval
   
      Case "Integer"
         Err.Clear
         destination_currency_object = CInt(destination_currency_units)
         If Err.Number Then
            destination_currency_object = 0
            Exit Function
         End If
         YCCO_ConvertGeneric = retval
   
      Case "Long"
         Err.Clear
         destination_currency_object = CLng(destination_currency_units)
         If Err.Number Then
            destination_currency_object = 0
            Exit Function
         End If
         YCCO_ConvertGeneric = retval
   
      Case "Single"
         destination_currency_object = destination_currency_units
         YCCO_ConvertGeneric = retval
   
      Case "Double"
         Err.Clear
         destination_currency_object = CDbl(destination_currency_units)
         If Err.Number Then
            destination_currency_object = 0
            Exit Function
         End If
         YCCO_ConvertGeneric = retval
      
      Case "Currency"
         Err.Clear
         destination_currency_object = CCur(destination_currency_units)
         If Err.Number Then
            destination_currency_object = 0
            Exit Function
         End If
         YCCO_ConvertGeneric = retval
   
      Case "Empty"   ' unassigned variant
         Err.Clear
         destination_currency_object = CVar(destination_currency_units)
         If Err.Number Then
            Exit Function
         End If
         YCCO_ConvertGeneric = retval
   
      Case "String"
         Err.Clear
         destination_currency_object = CStr(destination_currency_units)
         If Err.Number Then
            destination_currency_object = ""
            Exit Function
         End If
         YCCO_ConvertGeneric = retval
      
      Case "TextBox"
         Err.Clear
         destination_currency_object.Text = CStr(destination_currency_units)
         If Err.Number Then
            destination_currency_object.Text = ""
            Exit Function
         End If
         YCCO_ConvertGeneric = retval
   End Select

End Function

Public Function YCCO_Convert(ByVal source_currency_units!, ByRef destination_currency_units!) As Integer
   Dim search_string$, registry_string$, url_string$, data_string$, conversion_ratio!
   Dim web_update_required As Boolean
   Dim invert_conversion As Boolean
   
   On Error Resume Next

   If YCCO_is_busy = True Then
      YCCO_Convert = YCCO_BUSY_PROCESSING
      Exit Function
   End If

   YCCO_is_busy = True
   YCCO_Convert = YCCO_FAILED
   
   If YCCO_internet_transfer_control_defined = False Then
      On Error GoTo 0
      Err.Raise vbObjectError + 20, "YCCO", "Call YCCO_RegisterInternetTransferControl() before YCCO_Convert()!"
   End If
   If YCCO_currency_controls_defined = False Then
      On Error GoTo 0
      Err.Raise vbObjectError + 21, "YCCO", "Call YCCO_RegisterCurrencySelectionControls() before YCCO_Convert()!"
   End If
   If YCCO_source_currency_control.ListCount = 0 Or YCCO_destination_currency_control.ListCount = 0 Then
      On Error GoTo 0
      Err.Raise vbObjectError + 22, "YCCO", "Call YCCO_PopulateCurrencySelectionControls() before YCCO_Convert()!"
   End If
   
   search_string = GetConversionSearchString
   invert_conversion = False
   
   If Len(search_string) <> 6 Then
      On Error GoTo 0
      Err.Raise vbObjectError + 22, "YCCO", "Call YCCO_PopulateCurrencySelectionControls() before YCCO_Convert()!"
   End If
   
   If search_string = "------" Then
      destination_currency_units = source_currency_units
      YCCO_Convert = YCCO_SUCCEEDED
   Else
      web_update_required = True

      If YCCO_registry_usage_enabled = True Then
         registry_string = GetSetting("YCCO", "ConversionTable", search_string, "?")
      Else
         If search_string = YCCO_last_search_string Then
            registry_string = YCCO_last_registry_record
         Else
            registry_string = "?"
         End If
      End If
      
      If registry_string = "?" Then
         search_string = GetInvertedConversionSearchString
         
         If YCCO_registry_usage_enabled = True Then
            registry_string = GetSetting("YCCO", "ConversionTable", search_string, "?")
            invert_conversion = True
         Else
            If search_string = YCCO_last_search_string Then
               registry_string = YCCO_last_registry_record
               invert_conversion = True
            Else
               registry_string = "?"
            End If
         End If
      End If
         
      If registry_string <> "?" Then
         If Mid$(registry_string, 13, 1) = "-" Then
            
            conversion_ratio = CCur(Mid$(registry_string, 14))
            
            Select Case YCCO_internet_refresh_interval
               Case YCCO_EVERY_MONTH
                  If Left$(registry_string, 6) = Left$(GetRegistryDateString, 6) Then
                     web_update_required = False
                  End If
               Case YCCO_EVERY_DAY
                  If Left$(registry_string, 8) = Left$(GetRegistryDateString, 8) Then
                     web_update_required = False
                  End If
               Case YCCO_EVERY_HOUR
                  If Left$(registry_string, 10) = Left$(GetRegistryDateString, 10) Then
                     web_update_required = False
                  End If
               Case YCCO_EVERY_MINUTE
                  If Left$(registry_string, 12) = Left$(GetRegistryDateString, 12) Then
                     web_update_required = False
                  End If
               Case YCCO_EVERY_TIME
                  ' slip through with no test - web update alwats performed
            End Select
         End If
      End If
      
      If web_update_required = True Then
         url_string = "http://uk.finance.yahoo.com/m5?a=" & CStr(source_currency_units) & "&s=" & _
                      Mid$(YCCO_source_currency_control.Tag, YCCO_source_currency_control.ItemData(YCCO_source_currency_control.ListIndex), 3) & _
                      "&t=" & _
                      Mid$(YCCO_destination_currency_control.Tag, YCCO_destination_currency_control.ItemData(YCCO_destination_currency_control.ListIndex), 3)
                      
         data_string = GetInternetSiteContent(url_string)
         
         If data_string = "" Then
            YCCO_is_busy = False
            Exit Function
         End If
      
         If GetConversionRatio(data_string, conversion_ratio) = False Then
            YCCO_is_busy = False
            Exit Function
         End If
      
         invert_conversion = False
      End If
   
      If invert_conversion = True Then
         If conversion_ratio <> 0! Then
            conversion_ratio = 1! / conversion_ratio
         End If
      End If
         
      destination_currency_units = source_currency_units * conversion_ratio
   
      If conversion_ratio > 0! Then
         YCCO_Convert = YCCO_SUCCEEDED
      Else
         YCCO_Convert = YCCO_NO_DATA_AVAILABLE
      End If
   End If
            
   YCCO_is_busy = False
End Function
   
