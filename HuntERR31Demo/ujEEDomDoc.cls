VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ujEEDomDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
'                               HuntERR
'                   Error Handling and Reporting Library
'                    from URFIN JUS (www.urfinjus.net)
'                  Copyright 2001-2002. All rights reserved.
'version 3.1
'Error Extractor class for MSXML DOMDocument object
'=========================================================================================
Option Explicit
Public Function Extract(ByVal COMServer As Object, ByVal Param) As String
    Dim ObjParseErr As Object, strPointer As String
    If InException Or (COMServer Is Nothing) Then Exit Function
    On Error GoTo errHandler
    Set ObjParseErr = COMServer.parseError
    If ObjParseErr Is Nothing Then Exit Function
    With ObjParseErr
        If .errorCode <> 0 Then
            If .linepos > 0 Then strPointer = Space(.linepos - 1) & "^" & vbNewLine
            Extract = "Description: " & Replace(.reason, vbNewLine, "") & vbNewLine & _
                      "Source Line: " & Replace(.srcText, vbTab, " ") & vbNewLine & _
                      ".            " & strPointer & _
                      " Line=" & .Line & " Pos=" & .linepos
        End If
    End With
    Exit Function
errHandler:
    Extract = "  ErrExtractXML.Extract failed: " & Err.Description
End Function

'Note: HuntERR 3.1 removes indentation from message, and re-indents it properly,
'so we need to put a dot at the beginning of pointer line

