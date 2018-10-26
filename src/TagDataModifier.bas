Attribute VB_Name = "TagDataModifier"
Public Sub AddNewTagValue(ByRef tagPath As String, ByVal value As Double)
    
       
    
    ' Vars
    Dim pipt As PIPoint
    Dim pdata As pidata
    Dim srv As Server
    Dim nam As Integer
    Dim YesOrNoAnswerToMessageBox As String
    Dim QuestionToMessageBox As String
    Dim tagName As String
    Dim tagServer As String
  '  'Code to set target
  '  'Set refrences of PI SDK & PI Time
    On Error GoTo eh
  '
  '
    Call Tools.SplitTagServer(tagPath, tagServer, tagName)
  
    If tagServer <> "" Then
        Set srv = Servers(tagServer)
    Else
        Set srv = Servers.DefaultServer
    End If
    Set pipt = srv.PIPoints(tagName)
    Set pdata = pipt.Data
      
    QuestionToMessageBox = "Are you sure you are willing to save the value to PI Server?"
    YesOrNoAnswerToMessageBox = MsgBox(QuestionToMessageBox, vbYesNo, "Save Value to PI Server")
    If YesOrNoAnswerToMessageBox = vbNo Then
       MsgBox "Value Not Changed"
    Else
      pdata.UpdateValue value, pt, dmInsertDuplicates   ' Update PITag value in Pi server
      MsgBox "Value saved as :" & value
          
    End If
    Exit Sub
eh:
MsgBox Err.Description

End Sub
