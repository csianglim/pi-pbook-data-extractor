Attribute VB_Name = "TagExtractor"
Sub StartExtraction(isPreview As Boolean)

    ' Vars
    Dim srv As Server
    Dim tagName As String
    Dim tagPath As String
    Dim tagServer As String
    Dim tag As PIPoint
    Dim tagValue As PIValue
    Dim tagValueArray As PIValues
    Dim i As Integer, blockCount As Integer
    Dim blockSize As Integer
    Dim nvsSum As New NamedValues
    Dim totalValues As Long
    Dim DateStart As Date, DateEnd As Date
    Dim fileName As String, textData As String, textRow As String, fileNo As Integer
    Dim num_days As Long
    
    ' New vars for custom extraction duration
    Dim extractionMethod As String
    Dim retValue As NamedValues
    Dim pdata As PIData
    Dim ipid2 As IPIData2
    Dim duration As String
    Dim calcType As CalculationBasisConstants
    
    ' Clear previous values
    ThisDisplay.listValues.Clear
    fileNo = 0
    
    ' Get block size
    blockSize = CInt(ThisDisplay.comboBlockSize.value)
    
    ' Get extraction method
    extractionMethod = CStr(ThisDisplay.comboExtractionMode.value)
    
    ' Init controls

    On Error GoTo ErrHandler
        
    ' Retrieve fetch details
    DateStart = CDate(ThisDisplay.DateStart.value & " " & ThisDisplay.TimeStart.value)
    DateEnd = CDate(ThisDisplay.DateEnd.value & " " & ThisDisplay.TimeEnd.value)
        
    For i = ThisDisplay.ListSelectedTags.ListCount - 1 To 0 Step -1
    
        ' Retrieve tag name and server
        tagPath = ThisDisplay.ListSelectedTags.List(i, 0)
        Call Tools.SplitTagServer(tagPath, tagServer, tagName)
    
    
        ' TODO: Validate the input
    
        ' Do Extraction
        If tagServer <> "" Then
            Set srv = Servers(tagServer)
        Else
            Set srv = Servers.DefaultServer
        End If
        
        ' Check if tag exists
        If srv.GetPoints("tag = '" & tagName & "'").Count = 1 Then
        
        
            Set tag = srv.PIPoints(tagName)
            'Set tagValue = tag.Data.Snapshot
            
            ' Get number of values (not used)
            'totalValues = tag.Data.Summary(dateStart, dateEnd, btInside, astTotal)
            
            
            ' If preview
            If isPreview Then
                
                Dim lst As MSForms.ListBox
                Set lst = ThisDisplay.ListSelectedTags
                
                If lst.Selected(i) Then
                    
                    ' TODO: Respect blockSize when retrieving data using averages
                    If extractionMethod = "Recorded Values" Then
                        Set tagValueArray = tag.Data.RecordedValuesByCount(DateStart, blockSize, dForward, btInside)
                    Else
                        Set pdata = tag.Data
                        Set ipid2 = pdata ' get pointer to IPIData2 Interface https://techsupport.osisoft.com/Documentation/PI-SDK/PI_SDK_Interfaces/PI-SDK_Objects/IPIData2/IPIData2Ex.htm
                        duration = ThisDisplay.editInterval.value
                        If extractionMethod = "Time-Weighted Averages" Then
                            calcType = cbTimeWeighted
                        Else
                            calcType = cbEventWeighted
                        End If
                        Set retValue = ipid2.Summaries2(DateStart, DateEnd, duration, asAverage, calcType) 'Summaries2 returns a NamedValues collection of NamedValues, e.g. the Average.
                        Set tagValueArray = retValue("Average").value 'Get the PIValues from the NamedValues
                    End If
                    
                    ' Display results
                    For Each tagValue In tagValueArray
                      ThisDisplay.listValues.AddItem CStr(tagValue.TimeStamp.LocalDate) + vbTab + CStr(tagValue.value)
                    Next
                    
                    ' Display statistics
                    ThisDisplay.textTotalExtracted = CStr(tagValueArray.Count)
                
                    Exit Sub
                End If
                
                
            ' Continue to Standard Extraction
            Else
            
            
                ' Save to file
                totalValues = 0
                
                ' Remove unallowed chars from tagname
                fileName = Replace(tagName, "/", "_")
                fileName = Replace(fileName, "%", "_")
                fileName = Replace(fileName, ":", "_")
                
                
                fileName = ThisDisplay.editTargetDir.Text + "\" + ThisDisplay.editSaveFilePrefix.value + fileName + ".csv"
                fileNo = FreeFile 'Get first free file number
                Open fileName For Output As #fileNo
                
                Print #fileNo, "time,value," + tagName
                
                ' Iterate on blocks
                Do While DateStart < DateEnd
                    
                    ' TODO: Respect blockSize when retrieving data using averages
                    If extractionMethod = "Recorded Values" Then
                        Set tagValueArray = tag.Data.RecordedValuesByCount(DateStart, blockSize, dForward, btInside)
                    Else
                        Set pdata = tag.Data
                        Set ipid2 = pdata ' get pointer to IPIData2 Interface https://techsupport.osisoft.com/Documentation/PI-SDK/PI_SDK_Interfaces/PI-SDK_Objects/IPIData2/IPIData2Ex.htm
                        duration = ThisDisplay.editInterval.value
                        If extractionMethod = "Time-Weighted Averages" Then
                            calcType = cbTimeWeighted
                        Else
                            calcType = cbEventWeighted
                        End If
                        Set retValue = ipid2.Summaries2(DateStart, DateEnd, duration, asAverage, calcType) 'Summaries2 returns a NamedValues collection of NamedValues, e.g. the Average.
                        Set tagValueArray = retValue("Average").value 'Get the PIValues from the NamedValues
                        
                        ' To Exit the Loop
                        DateStart = DateEnd
                    End If
                    
                    ' Check if not more values available
                    If tagValueArray.Count <= 1 Then
                        Exit Do
                    End If
                    
                
                    ' Store block to file
                    For Each tagValue In tagValueArray
                        Print #fileNo, CStr(tagValue.TimeStamp.LocalDate) + "," + CStr(tagValue.value)
                    Next
                
                
                    ' Get last date from block
                    DateStart = DateAdd("s", 1, tagValueArray.Item(tagValueArray.Count).TimeStamp.LocalDate)
        
                    totalValues = totalValues + tagValueArray.Count
                    Set tagValueArray = Nothing
                    ThisDisplay.editStatus.value = tagPath + vbTab + "(" + CStr(totalValues) + ")"
                    
                    DoEvents
                
                Loop
                
                
                Close #fileNo
                    
                ' Remove item from selected column and add it to processed
                ThisDisplay.ListSelectedTags.RemoveItem (i)
                ThisDisplay.listExtractedTags.AddItem (tagPath + vbTab + CStr(totalValues))
        
            End If
       End If ' (tag exist check)
    Next i
    
    ThisDisplay.editStatus.value = ""
    MsgBox ("Extraction finished successfully!")
    
    
Exit Sub
ErrHandler:

    MsgBox "Tag: " & tagPath & vbCrLf & _
            "Error Line: " & Erl & vbCrLf & _
            "Error: (" & Err.Number & ") " & Err.Description, vbCritical, "Critical Error"

   'Msg = Err.Source & Err.Description
   'MsgBox Msg

   If fileNo > 0 Then
        Close #fileNo
   End If
   
  
End Sub
