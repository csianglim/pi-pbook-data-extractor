Attribute VB_Name = "PISymbolsAnalyzer"
Option Explicit

'ProcessBook Objects
Dim objPIPBTOC As ProcBook
Dim objDsp As Display
Dim objPIPB As Object
Dim objPIPBDisplay As Object


Public Sub ListPISymbols()

    Dim answer As Integer
    

    'This subroutine loops through the entries in ProcessBook.  When it finds
    'a display or linked display it passes the name of the ProcessBook Workbook
    'and the display to another subroutine (OutputSymbolsToExcel).  The OutputSymbolsToExcel
    'routine then loops through all the display symbols and outputs to Excel the
    'Process book name, the display name, and the tags associated with the symbols in the display.
    'It is envoked via a Macro button on the ListPISymbols.PDI display.
    
    ' Index for entry in ProcessBook objPIPBTOC
    Dim varEntry As Entry
    
    'Text of entry as it appears in the ProcessBook window
    Dim strEntryName As String
        
    'ProcessBook and Display name variables
    Dim strpipbPath As String ' PB (.piw) filename
    Dim strDisplayName As String
    
        
    'If a copy of ProcessBook is not running want to advise user to open _
    the desired workbook file.
    On Error GoTo ErrHandlerPBWB
    
    'Get PI-ProcessBook application object
    Set objPIPB = GetObject(, "PIProcessBook.Application.2") 'This will generate an error if a PIW file is not open
        
    'Reference PI_ProcessBook current active display
    Set objPIPBDisplay = objPIPB.ActiveDisplay
       
    'Get Entries Objects for ProcessBook
    Set objPIPBTOC = Application.ProcBooks(1)
    
    'Get PIW filename
    strpipbPath = objPIPBTOC.Path
    
    On Error Resume Next 'If made it this far set error behavior to normal
    
    ' Iterate on open Diplays
    Dim currDisplay As Display
    For Each currDisplay In objPIPBTOC.Displays
    
   
        ' Ask if we need to process this display
        answer = MsgBox("Imort tags from the following display " + currDisplay.Path + " ?", vbYesNo + vbQuestion)
        If answer = vbYes Then
        
            Call GetTags(currDisplay)
        End If
        
    Next currDisplay
    
    'Exiting subroutine so want to release system resources associated with
    'ProcessBook and Excel objects
    Set objPIPB = Nothing
    Set objPIPBTOC = Nothing

    
Exit Sub
    
ErrHandlerPBWB:

    MsgBox ("This utility requires that you have a copy of PI-ProcessBook (.PIW file) running.  Open the desired PIW file."), vbExclamation
    Exit Sub
    
End Sub


Private Sub GetTags(Current_Display As Display)

    'Display objects
    Dim Current_Symbols As Object
    Dim Symbol_Object() As Object

    'Counters
    Dim Object_Count As Integer
    Dim i As Integer, J As Integer, L As Integer, R As Integer

    'Symbol property variables
    Dim strTagId(10) As String
    Dim strServerId(10) As String
    Dim Server_Tag_ComboName(10) As String
    Dim value As String
    Dim Text_String As String

    On Error Resume Next

    'Get a handle on the current ProcessBook display _
    and on the symbols within the display
    'Current_Display.Activate = True
    Set Current_Symbols = Current_Display.Symbols

    'Figure out how many object exist on the current display
    Object_Count = Current_Symbols.Count
    ReDim Symbol_Object(Object_Count)
    
    'Go through all symbols in display and set Symbol_Object(I) _
    to the properties of that symbol
    i = 1
    For i = 1 To Object_Count
        Set Symbol_Object(i) = Current_Display.Symbols.Item(i)
    Next i




    'Reset counter
    i = 1
    
    'Loop through each display and process all object symbols
    For i = 1 To Object_Count
    
        'Type of symbol
        'Line       = 1
        'Rectangle  = 2
        'Ellipse    = 3
        'Text       = 4
        'Polygon    = 6
        'Value      = 7
        'Button     = 9
        'Arc        = 11
        'Trend      = 10
        'Bar        = 12
        'Bitmap     = 13
    
    
        'Check for symbol type 4 which is text. Output to Worksheet cell 8
        'If Symbol_Object(i).Type = 4 Then
        '    Text_String = Symbol_Object(i).Contents
        '    ThisDisplay.ListSelectedTags.AddItem (Text_String)
        'End If
    
        'If a Value is also a multi_state then PtCount will be to capture both TagNames
        If Symbol_Object(i).IsMultiState = True And Symbol_Object(i).Type = 7 Then 'two Tags are associated with this Value Object
        
            J = 1
            'Get Tag Name
            Server_Tag_ComboName(J) = Symbol_Object(i).GetTagName(J)
            'Parse node name and tag name
            For J = 1 To Symbol_Object(i).PtCount
        
                If Server_Tag_ComboName(J) <> "" Then
                    For L = Len(Server_Tag_ComboName(J)) To 1 Step -1
                        value = Mid(Server_Tag_ComboName(J), L, 1)
                        'Check for split
                        If Mid(Server_Tag_ComboName(J), L, 1) = "\" Then
                        'split between server id and Tag Name found
                            strTagId(J) = Mid(Server_Tag_ComboName(J), L + 1, Len(Server_Tag_ComboName(J)))
                            strServerId(J) = Mid(Server_Tag_ComboName(J), 3, L - 3)
                            GoTo WriteServerAndTag7
                        ElseIf L <= 1 Then
                            strTagId(J) = Server_Tag_ComboName(J)
                            strServerId(J) = ""
          
                        End If
                    Next
                End If
            
WriteServerAndTag7:
            
            ' Check if tag exists
            'if ThisDisplay.ListSelectedTags.
            
            'ThisDisplay.ListSelectedTags.AddItem (strServerId(1) + "/" + strTagId(J))
            
            Next J
        
        'Trend Object
        ElseIf Symbol_Object(i).Type = 10 Then
            J = 1
            For J = 1 To Symbol_Object(i).PtCount
                'Get tag name
                Server_Tag_ComboName(J) = Symbol_Object(i).GetTagName(J)
            
                If Server_Tag_ComboName(J) <> "" Then
                    For L = Len(Server_Tag_ComboName(J)) To 1 Step -1
                        value = Mid(Server_Tag_ComboName(J), L, 1)
                    
                        'Check for split between node and tag name
                        If Mid(Server_Tag_ComboName(J), L, 1) = "\" Then
                    
                            'Tagname
                            strTagId(J) = Mid(Server_Tag_ComboName(J), L + 1, Len(Server_Tag_ComboName(J)))
                            'Node name
                            strServerId(J) = Mid(Server_Tag_ComboName(J), 3, L - 3)
                            GoTo WriteServerAndTag10
                        ElseIf L <= 1 Then
                            strTagId(J) = Server_Tag_ComboName(J)
                            strServerId(J) = ""
          
                        End If
                    Next
                End If
            
WriteServerAndTag10:
        
        
            'ThisDisplay.ListSelectedTags.AddItem (strServerId(1) + "/" + strTagId(J))
        
            Next J
        
        'Tag is associated with this Object
        ElseIf Symbol_Object(i - 1).PtCount >= 1 Then
            J = 1
            'Get tag name
            Server_Tag_ComboName(J) = Symbol_Object(i).GetTagName(J)
        
            If Server_Tag_ComboName(J) <> "" Then
        
                For L = Len(Server_Tag_ComboName(J)) To 1 Step -1
                    value = Mid(Server_Tag_ComboName(J), L, 1)
                
                    'check for split between node and tag name
                    If Mid(Server_Tag_ComboName(J), L, 1) = "\" Then
                        'get tagname
                        strTagId(J) = Mid(Server_Tag_ComboName(J), L + 1, Len(Server_Tag_ComboName(J)))
                        'get node name
                        strServerId(J) = Mid(Server_Tag_ComboName(J), 3, L - 3)
                        GoTo WriteServerAndTag11
                    ElseIf L = 1 Then
                        strTagId(J) = Server_Tag_ComboName(J)
                        strServerId(J) = ""
          
                    End If
                Next
            End If
        
WriteServerAndTag11:
        
            'ThisDisplay.ListSelectedTags.AddItem (strServerId(1) + "/" + strTagId(J))
            ThisDisplay.ListSelectedTags.AddItem (Server_Tag_ComboName(J))
            
            

                
            
        
        End If

    Next ' Next object
   
End Sub

