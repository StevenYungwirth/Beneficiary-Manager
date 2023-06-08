Attribute VB_Name = "XMLProcedures"
Option Explicit

Private Property Get XMLClientList() As DOMDocument60
    Set XMLClientList = ProjectGlobals.ClientListFile
End Property

Public Function XPathExpression(value As String) As String
    'If the account/bene name has single or double quotes, wrap the search string in the opposite
    If InStr(value, "'") = 0 Then
        XPathExpression = "'" & value & "'"
    ElseIf InStr(value, """") = 0 Then
        XPathExpression = """" & value & """"
    Else
        XPathExpression = "concat('" & Replace(value, "'", "',""'"",'") & "')"
    End If
End Function

Public Sub FlagNodeTypeInList(nodeType As String, sheetName As String, flagValue As Boolean)
    'Select the nodes
    Dim nodesToFlag As IXMLDOMNodeList
    Set nodesToFlag = XMLClientList.SelectNodes("//" & nodeType)
    
    'Add the flag to each node
    Dim nodeItemFound As Variant
    For Each nodeItemFound In nodesToFlag
        Dim nodeItem As IXMLDOMElement: Set nodeItem = nodeItemFound
        FlagNodeInList nodeItem, sheetName, flagValue
    Next nodeItemFound
End Sub

Public Sub FlagNodeInList(selectedNode As IXMLDOMElement, sheetName, flagValue As Boolean)
    If sheetName <> vbNullString Then
        selectedNode.setAttribute "In_" & sheetName, CStr(flagValue)
    End If
End Sub

Public Sub DifferingInfoCheck(infoType As String, sheetInfo As Variant, listInfo As Variant, DifferingInfoDict As Dictionary, dictKey As Variant, _
                              sheetName As String, componentNode As IXMLDOMNode, identifyingType As String, IdentifyingData As String)
    'Make note that the info on this row is different than what's in the list, unless the node was already added in the process
    If sheetInfo <> listInfo And Not DifferingInfoDict.Exists(dictKey) And XMLProcedures.GetAddDate(componentNode) < ProjectGlobals.ImportTime Then
        Dim dictString As String
        If sheetInfo = componentNode.SelectSingleNode(infoType).Text Then
            'The data node was updated with the info from the importing sheets
            If IdentifyingData = vbNullString Then
                dictString = identifyingType & " - " & infoType & " - Old: " & listInfo & " | New: " & sheetInfo
            Else
                dictString = identifyingType & " - " & infoType & " - " & IdentifyingData & " - Old: " & listInfo & " | New: " & sheetInfo
            End If
        Else
            'The data node wasn't updated with the info from the importing sheets because a higher priority sheet disallowed the change
            'Get the higher priority sheet name
            Dim higherPriorityName As String
            If componentNode.SelectSingleNode(infoType).Attributes.getNamedItem("Updated_By") Is Nothing Then
                higherPriorityName = componentNode.SelectSingleNode(infoType).Attributes.getNamedItem("Added_By").Text
            Else
                higherPriorityName = componentNode.SelectSingleNode(infoType).Attributes.getNamedItem("Updated_By").Text
            End If
            If IdentifyingData = vbNullString Then
                dictString = identifyingType & " - " & infoType & " - XML: " & listInfo & " | " & sheetName & " : " & sheetInfo
            Else
                dictString = identifyingType & " - " & infoType & " - " & IdentifyingData & " - XML: " & listInfo & " | " & sheetName & " : " & sheetInfo
            End If
        End If
        
        DifferingInfoDict.Add dictKey, dictString
    End If
End Sub

Public Function GetAddDate(componentNode As IXMLDOMNode) As Date
    If Not componentNode Is Nothing Then
        If Not componentNode.Attributes.getNamedItem("Added_On") Is Nothing Then
            If IsDate(componentNode.Attributes.getNamedItem("Added_On").Text) Then
                GetAddDate = CDate(componentNode.Attributes.getNamedItem("Added_On").Text)
            Else
            End If
        Else
        End If
    Else
    End If
End Function

Public Function MoveNode(nodeToMove As IXMLDOMNode, newParent As IXMLDOMNode) As IXMLDOMNode
    'Duplicate the node
    Dim nodeClone As IXMLDOMNode
    Set nodeClone = nodeToMove.CloneNode(True)
    
    'Add it to the new parent
    newParent.appendChild nodeClone
    
    'Delete the original node
    nodeToMove.parentNode.RemoveChild nodeToMove
    
    'Return the moved node
    Set MoveNode = nodeClone
End Function

Public Sub FormatAndSaveXML()
    PrettyPrintXML XMLClientList.XML
    PrettyPrint XMLClientList
    XMLClientList.Save ProjectGlobals.ClientListFilePath
End Sub

Private Function PrettyPrintXML(XML As String) As String
  Dim Reader As New SAXXMLReader60
  Dim Writer As New MXXMLWriter60

  Writer.Indent = True
  Writer.standalone = False
  Writer.omitXMLDeclaration = False
  Writer.Encoding = "utf-8"

  Set Reader.contentHandler = Writer
  Set Reader.dtdHandler = Writer
  Set Reader.errorHandler = Writer

  Call Reader.putProperty("http://xml.org/sax/properties/declaration-handler", _
          Writer)
  Call Reader.putProperty("http://xml.org/sax/properties/lexical-handler", _
          Writer)

  Call Reader.Parse(XML)

  PrettyPrintXML = Writer.output
End Function

Private Sub PrettyPrint(Parent As IXMLDOMNode, Optional Level As Integer)
  Dim Node As IXMLDOMNode
  Dim Indent As IXMLDOMText

  If Not Parent.parentNode Is Nothing And Parent.ChildNodes.Length > 0 Then
    For Each Node In Parent.ChildNodes
      Set Indent = Node.OwnerDocument.createTextNode(vbNewLine & String(Level, vbTab))

      If Node.nodeType = NODE_TEXT Then
        If Trim(Node.Text) = "" Then
          Parent.RemoveChild Node
        End If
      ElseIf Node.PreviousSibling Is Nothing Then
        Parent.InsertBefore Indent, Node
      ElseIf Node.PreviousSibling.nodeType <> NODE_TEXT Then
        Parent.InsertBefore Indent, Node
      End If
    Next Node
  End If

  If Parent.ChildNodes.Length > 0 Then
    For Each Node In Parent.ChildNodes
      If Node.nodeType <> NODE_TEXT Then PrettyPrint Node, Level + 1
    Next Node
  End If
End Sub
