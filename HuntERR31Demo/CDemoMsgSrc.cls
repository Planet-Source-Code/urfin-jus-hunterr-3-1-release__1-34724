VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDemoMsgSrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
'                               HuntERR
'                   Error Handling and Reporting Library
'                    from URFIN JUS (www.urfinjus.net)
'                  Copyright 2001-2002. All rights reserved.
'version 3.1, 04/25/2002
'Demo Application, Demo Message Source class
'=========================================================================================
Option Explicit
'GetMessage is the only function that must be provided in a class.
'It must return message for specified key. Note that all keys should start with #,
'which is indication that this is message key, not hard-coded message.
'In our demo class we show how you can cover a situation when message is not found
'in your message source (databas table), just because you forgot to put it there.
'In our implementation application provides Key which consists of
'message ID followed by "||" and then hard-coded default that should be
'shown if message is not found, for example:
' "#1||This is default for message number 1."
'Our version of GetMessage parses message ID from Key, tries to retrieve
'message by ID, and if fails then returns default from Key.

Public Function GetMessage(ByVal Key As String) As String
    Const Delim = "||"
    Dim P As Long, MsgID As Long, Msg As String
    On Error GoTo errHandler
    Key = Mid$(Key, 2)
    P = InStr(1, Key & Delim, Delim) 'Account for situation when there is no subst message, id only
    MsgID = SafeLng(Left$(Key, P - 1))
    If MsgID = 0 Then
        Msg = Key 'No key was provided; this is hard-coded message
        Else
        Msg = GetMessageByID(MsgID) 'If not found returns empty string
        If Msg = "" Then
            'put notice that message was not found, and attach default message
            Msg = "(From CDemoMsgSrc: Message ID=" & MsgID & " not found, using default) " & _
                vbNewLine & Mid$(Key, P + Len(Delim))
        End If 'Msg = "" ...
    End If 'MsgID = 0 ...
    GetMessage = Msg
    Exit Function
errHandler:
    GetMessage = "Error in CDemoMsgSrc(Key=" & Key & "): " & Err.Description
End Function

Private Function SafeLng(ByVal V) As Long
    On Error Resume Next
    SafeLng = CLng(Trim$(V))
End Function


Private Function GetMessageByID(ByVal MsgID As Long) As String
    On Error GoTo errHandler
    Select Case MsgID
        Case 1
            'Lets pretend we queried messages database table and
            'retrieved the following message with MsgID=1
            GetMessageByID = "This is message from CDemoMsgSrc with params [%1] and [%2] ||... and line break."
        Case Else
            'Lets pretend message was not found in database
            GetMessageByID = ""
    End Select
    Exit Function
errHandler:
    GetMessageByID = "Failed to retrieve message: " & Err.Description
End Function
