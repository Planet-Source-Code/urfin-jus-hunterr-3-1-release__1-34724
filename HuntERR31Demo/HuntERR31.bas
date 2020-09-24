Attribute VB_Name = "HuntERR31"
'========================================================================================
'                               HuntERR
'                   Error Handling and Reporting Library
'                    from URFIN JUS (www.urfinjus.net)
'                  Copyright 2001, 2002. All rights reserved.
'version 3.1
'=========================================================================================
'Conditional compilation constants
'  H_NOCOMPLUS = 1     -- No COM+ support
'  H_NOENUMS = 1       -- Exclude declarations of ENUM_ERRMAP enumeration.
'  H_EXTBASE = 1       -- ERRMAP_BASE constant is defined in other module.
'  H_NOSTOP  = 1       -- Don't stop on errors. If set, ErrMustStop returns false.
'
'Attention! For error logging to Oracle database
'comment/uncomment lines in ErrSaveToDB method.
'
'Public members =========================================================================
'
'Public Enum ENUM_ERRMAP
'Public Enum ENUM_ERROR_ACTION

'Public Function ErrorIn(ByVal MethodHeader As String, _
'                        Optional ByVal arrArgs, _
'                        Optional ByVal ErrorAction As Long = EA_DEFAULT, _
'                        Optional ByVal DbObject As Object, _
'                        Optional ByVal EnvVarNames As String, _
'                        Optional ByVal arrEnvVars, _
'                        Optional ByVal TransControlObject As Object) As String
'Public Sub Check(ByVal Cond As Boolean, _
'                 ByVal AnErrNumber As Long, _
'                 ByVal AnErrDescr As String, _
'                 Optional ByVal Values, _
'                 Optional ByVal AHelpFile, _
'                 Optional ByVal AHelpContext)
'
'Public Sub ErrPreserve()
'Public Sub ErrClear()
'Public Sub ErrRestore()
'Public Sub ErrContinue(ByVal AnErrNum As Long, ByVal AnErrReport)
'Public Sub ErrGetFromServer(ByVal Extractor, ByVal COMServer As Object, _
'        Optional ByVal Param, Optional ByVal Comment As String)

'Public Function IsException(ByVal ErrNumber As Long) As Boolean
'Public Function InException() As Boolean
'Public Function InPropagation() As Boolean

'Public Property Get ErrReport() As String
'Public Property Get ErrReportHTML() As String
'Public Property Get ErrNumber() As Long
'Public Property Get ErrSource() As String
'Public Property Get ErrDescription() As String
'Public Property Get ErrOrigSource() As String
'Public Property Get ErrOrigDescription() As String
'Public Property Get/Let ErrAccumBuffer() As String
'Public Function ErrExtractFromReport(ByVal AReport As String, ByVal FromStr As String, ByVal TillStr As String)
'
'Public Property Get/Set ErrMessageSource() As Object
'
'Public Sub ErrRlsObjs(Optional ByRef Obj1 As Object, ...)
'Public Sub ErrCloseFiles(ParamArray Files())

'Public Function ErrMustStop() As Boolean
'Public Property Get ErrStopFlag() As Boolean

'Public Function ErrSaveToEventLog() As Boolean
'Public Function ErrSaveToFile(Optional ByVal ErrFileName As String = "Errors.txt") As Boolean
'Public Function ErrSaveToDB(ByVal ConnectString As String, _
'                       Optional ByVal AppID As Long = 1, _
'                       Optional ByVal ProcName As String = "spErrorLogInsert") As Boolean

'Public Function ErrInIDE() As Boolean
'Public Function ErrInsInto(ByVal IntoStr As String, ByVal Values) As String

Option Explicit
'Error numbers (vbObjectError + [1..4095])  are used by OLE DB;
'See support.microsoft.com/support/kb/articles/Q168/3/54.ASP
'If you use OracleObjects: OO errors are just  above vbObjectError + 4096,
'so you need to redefine ERRMAP_BASE
#If H_CUSTOMDEF = 1 Then
    'App defines this constant in some Bas module
#Else
    Const ERRMAP_BASE = vbObjectError + 4096  'vbObjectError = $H80040000 = -2147221504
#End If

'Error Map. Defines error ranges for errors and exceptions
'Define your custom errors starting with ERRMAP_APP_FIRST + 1.
#If H_NOENUMS = 1 Then
    'ENUM_ERRMAP is declared in other module (we recommend public COM class)
#Else
Public Enum ENUM_ERRMAP
        ERR_ACCUMULATE = 0                          'Check Sub accumulates messages
    ERRMAP_FIRST = ERRMAP_BASE
    ERRMAP_RESERVED_FIRST = ERRMAP_FIRST            'Errors reserved for HuntERR and UJ apps.
       ERR_SYSEXCEPTION                             'System exception
    ERRMAP_RESERVED_LAST = ERRMAP_RESERVED_FIRST + 100
    ERRMAP_EXC_FIRST = ERRMAP_RESERVED_LAST + 1     'Exceptions - reraised by ErrorIn
        EXC_GENERAL = ERRMAP_EXC_FIRST              'Use it if you don't need specific number
        EXC_VALIDATION                              'User input validation exception
        EXC_MULTIPLE                                'Multiple messages in error description
        EXC_CANCELLED                               'Cancelled operation, silent exception
    ERRMAP_EXC_LAST = ERRMAP_EXC_FIRST + 1000
    ERRMAP_APP_FIRST
        ERR_GENERAL = ERRMAP_APP_FIRST              'Use it if you don't need specific number
        'Application errors here
End Enum
'Flags controlling actions of ErrorIn, through ErrorAction parameter
Public Enum ENUM_ERROR_ACTION
    EA_RERAISE = 1         'Reraise error
    EA_ADVANCED = 2        'Build Error report
    EA_SET_ABORT = 4       'Call SetAbort on current object's context.
    EA_DISABLE_COMMIT = 8  'Call DisableCommit on current objects' context. Recommended.
    EA_ROLLBACK = &H10     'Call Connection.Rollback
    EA_WEBINFO = &H20      'Add web request information
    EA_CONN_CLOSE = &H40   'Close connection
    EA_DEFAULT = EA_ADVANCED + EA_RERAISE + EA_WEBINFO + EA_DISABLE_COMMIT 'Default
    'The following constants are defined for convenience
    EA_NORERAISE = EA_ADVANCED + EA_WEBINFO + EA_DISABLE_COMMIT
    EA_DFTRBK = EA_ADVANCED + EA_RERAISE + EA_WEBINFO + EA_ROLLBACK
    EA_DFTRBKCLS = EA_ADVANCED + EA_RERAISE + EA_WEBINFO + EA_ROLLBACK + EA_CONN_CLOSE
End Enum

#End If 'H_NOENUMS = 1 ... Else ...

Private mErrNumber As Long, mErrSource As String, mErrDescr As String, mErl As Long
Private mErrHelpFile As String, mErrHelpContext As String
Private mErrReport As String, mErrAPI As Long
Private mErrPreserved As Boolean 'Set by ErrPreserve, cleared by ErrorIn at the end of processing
Private mExtBuffer As String     'Errors from COMServers, added by ErrGetFromServer
Private mMsgSrc As Object 'External message source
Private mRlsdObjs As String, mSavedConn As Object
Const MAX_NON_LONG_DATA = 40
Private mLDBuffer As String 'Long data buffer
Private mAccumBuffer As String 'Error accumulation buffer

Public Const _
    SRC_ERRORIN = "HuntERR.ErrorIn", _
    SRC_CHECK = "HuntERR.Check", _
    SRC_SYSHANDLER = "HuntERR.SysExcHandler", _
    SUBST_CRLF = "||", _
    STR_NOSTOP = "NoStop"
    
'API declarations
Private Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function FormatMessageAPI Lib "kernel32" Alias "FormatMessageA" _
    (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
    Arguments As Long) As Long
Private Declare Sub APISleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Const MAX_RETRY = 1# / 24# / 60# / 60#  ' 1000 ms

'The main super-function. Should be called in error handling blocks
Public Function ErrorIn(ByVal MethodHeader As String, _
                        Optional ByVal arrArgs, _
                        Optional ByVal ErrorAction As Long = EA_DEFAULT, _
                        Optional ByVal DBObject As Object, _
                        Optional ByVal EnvVarNames As String, _
                        Optional ByVal arrEnvVars, _
                        Optional ByVal TransControlObject As Object) As String
    Dim MethodName As String, ArgNames As String, objConn As Object, strMsg As String
    If (Not mErrPreserved) And (Err.Number = 0) Then
        Err.Number = ERR_GENERAL
        Err.Description = "ErrorIn: Error information was lost. " & _
            "To fix: call ErrPreserve before doing anything in error handler."
    End If
    ErrPreserve
    Set objConn = ErrGetConnectionObject(DBObject)
    If InException Then
        mErrReport = ""
        ErrTerminateTrans ErrorAction, objConn, TransControlObject
        Else
        If FlagSet(ErrorAction, EA_ADVANCED) Then
            ParseMethodHeader MethodHeader, MethodName, ArgNames
            If InPropagation Then
                mErrReport = mErrDescr
                Else 'Initial processing
                ReportInit MethodName
                ReportAddAPIError
                ReportAddADOInfo objConn, DBObject
                If FlagSet(ErrorAction, EA_WEBINFO) Then ReportAddWebInfo
            End If 'InPropagation...
            ReportAddCallStackInfo MethodName, ArgNames, arrArgs, EnvVarNames, arrEnvVars
            ReportAddExtErrors
            ReportAddRlsdObjsList
        End If
        ErrTerminateTrans ErrorAction, objConn, TransControlObject
        If FlagSet(ErrorAction, EA_CONN_CLOSE) Then ErrConnClose objConn
        If FlagSet(ErrorAction, EA_ADVANCED) Then
            mErrSource = SRC_ERRORIN
            mErrDescr = mErrReport
            ErrorIn = mErrReport
        End If
    End If 'InException ... else ...
    Set mSavedConn = Nothing
    mExtBuffer = ""
    ErrRestore
    mErrPreserved = False 'Drop the flag, as we made use of err info
    If FlagSet(ErrorAction, EA_RERAISE) Then _
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Sub Check(ByVal Cond As Boolean, _
                 ByVal AnErrNumber As Long, _
                 ByVal AnErrDescr As String, _
                 Optional ByVal Values, _
                 Optional ByVal AHelpFile, _
                 Optional ByVal AHelpContext)
If Not Cond Then
    If Not mMsgSrc Is Nothing And Left$(AnErrDescr, 1) = "#" Then AnErrDescr = mMsgSrc.GetMessage(AnErrDescr)
    AnErrDescr = Replace(AnErrDescr, SUBST_CRLF, vbNewLine)
    If Not IsMissing(Values) Then AnErrDescr = ErrInsInto(AnErrDescr, Values)
    If AnErrNumber = 0 Then
        'Accumulate message in buffer
        mAccumBuffer = mAccumBuffer & IIf(mAccumBuffer = "", "", vbNewLine) & AnErrDescr
        Else
        Err.Raise AnErrNumber, SRC_CHECK, AnErrDescr, AHelpFile, AHelpContext
    End If 'AnErrNumber = 0
End If ' Cond
End Sub

'Preserves Err object properties for later use by ErrorIn
Public Sub ErrPreserve()
    If (Err.Number <> 0) And (Not mErrPreserved) Then
        mErrNumber = Err.Number
        mErrDescr = Err.Description
        mErrSource = Err.Source
        mErrHelpFile = Err.HelpFile
        mErrHelpContext = Err.HelpContext
        mErl = Erl
        mErrAPI = Err.LastDllError 'We need to do it here, GetLastError information is vulnerable
        mErrPreserved = True
    End If
End Sub

Public Sub ErrClear()
    mErrNumber = 0
    mErrDescr = ""
    mErrSource = ""
    mErrHelpFile = ""
    mErrHelpContext = 0
    mErl = 0
    mErrAPI = 0
    mErrPreserved = False
End Sub

'Restores Err object properties
Public Sub ErrRestore()
    If mErrPreserved Then
        Err.Clear
        Err.Number = mErrNumber
        Err.Source = mErrSource
        Err.Description = mErrDescr
        If mErrHelpContext <> "" Then Err.HelpContext = mErrHelpContext
        Err.HelpFile = mErrHelpFile
    End If
End Sub

'Continues propagation process
Public Sub ErrContinue(ByVal AnErrNum As Long, ByVal AnErrReport)
    Err.Raise AnErrNum, SRC_ERRORIN, AnErrReport
End Sub

'Retrieves error information from custom COM object (server).
'Uses extractor class provided by application.
'Extactor object should have Extract method returning formatted error information.
Public Sub ErrGetFromServer(ByVal Extractor, ByVal COMServer As Object, _
        Optional ByVal Param, Optional ByVal Comment As String)
    Dim objExtr As Object, sMsg As String, sHdr As String
    ErrPreserve
    On Error GoTo errHandler
    sHdr = "    COM Server Errors: Server=" & ErrVarToString(COMServer) & _
        " Extractor=" & ErrVarToString(Extractor) & IIf(Comment = "", "", "  [" & Comment & "]")
    If IsObject(Extractor) Then Set objExtr = Extractor Else Set objExtr = CreateObject(CStr(Extractor))
    ErrRestore 'Set Err object with original error information, in case if extractor needs it
    sMsg = objExtr.Extract(COMServer, Param)
    If sMsg <> "" Then
        sMsg = Unindent(Trim$(sMsg))
        If Right$(sMsg, Len(vbNewLine)) = vbNewLine Then sMsg = Left$(sMsg, Len(sMsg) - 1) 'cut off new line char
        mExtBuffer = mExtBuffer & sHdr & vbNewLine & Indent(sMsg, 6) & vbNewLine
    End If
    GoTo ExitSub 'We cannot use Exit Sub - it clears Err object
errHandler:
    mExtBuffer = mExtBuffer & sHdr & vbNewLine & _
        "    ErrorIn failed to extract error information: " & Err.Description
ExitSub:
    ErrRestore 'Restore err object for code in error handler
End Sub

'Returns true if parameter is in range reserved for Exceptions
Public Function IsException(ByVal AnErrNumber As Long) As Boolean
    IsException = (AnErrNumber >= ERRMAP_EXC_FIRST) And (AnErrNumber <= ERRMAP_EXC_LAST)
End Function

'Returns true if ErrNumber is in range reserved for Exceptions
Public Function InException() As Boolean
    InException = IsException(ErrNumber)
End Function

'Returns true if error was raised by ErrorIn, thus propagation is in progress
Public Function InPropagation() As Boolean
    InPropagation = (ErrSource = SRC_ERRORIN)
End Function

'Returns report prepared by last call to ErrorIn, or current Err.Description
Public Property Get ErrReport() As String
    If mErrReport <> "" Then
        ErrReport = mErrReport
    ElseIf InPropagation Then
        ErrReport = Err.Description
    Else
        ErrReport = ""
    End If
End Property

'Returns HTML-formatted error report
Public Property Get ErrReportHTML() As String
    Dim sHTML As String
    sHTML = ErrReport
    sHTML = Replace(sHTML, "&", "&amp;")
    sHTML = Replace(sHTML, "<", "&lt;")
    sHTML = Replace(sHTML, ">", "&gt;")
    sHTML = Replace(sHTML, """", "&quot;")
    sHTML = Replace(sHTML, " ", "&nbsp;")
    sHTML = Replace(sHTML, vbNewLine, "<br>" & vbNewLine)
    ErrReportHTML = sHTML
End Property

'Returns error number saved by last call to ErrPreserve or ErrorIn, or curre
Public Property Get ErrNumber() As Long
    ErrNumber = mErrNumber
End Property

'Returns error source saved by last call to ErrPreserve or ErrorIn, or curre
Public Property Get ErrSource() As String
    ErrSource = mErrSource
End Property

'Returns error description saved by last call to ErrPreserve or ErrorIn, or curre
Public Property Get ErrDescription() As String
    ErrDescription = mErrDescr
End Property

Public Property Get ErrHelpContext() As String
    ErrHelpContext = mErrHelpContext
End Property

Public Property Get ErrHelpFile() As String
    ErrHelpFile = mErrHelpFile
End Property

'Extracts original error source from error report
Public Property Get ErrOrigSource() As String
    ErrOrigSource = IIf(InPropagation, ErrExtractFromReport(vbNewLine & "  Source: ", vbNewLine), mErrSource)
End Property

'Extracts original error description from error report
Public Property Get ErrOrigDescription() As String
    On Error GoTo errHandler
    Dim P As Long
    If InPropagation Then
        P = InStr(1, ErrReport & vbNewLine, vbNewLine)
        ErrOrigDescription = Left$(ErrReport, P - 1)
    Else
        ErrOrigDescription = mErrDescr
    End If
    Exit Property
errHandler:
End Property

Public Property Get ErrAccumBuffer() As String
    ErrAccumBuffer = mAccumBuffer
End Property

Public Property Let ErrAccumBuffer(ByVal AValue As String)
    mAccumBuffer = AValue
End Property

'Extracts substring from error report
Public Function ErrExtractFromReport(ByVal FromStr As String, ByVal TillStr As String)
    Dim PStart As Long, PEnd As Long, Rpt As String
    Rpt = ErrReport
    If Rpt = "" Then Exit Function
    PStart = InStr(1, Rpt, FromStr) + Len(FromStr)
    PEnd = InStr(PStart, Rpt, TillStr) - 1
    If (PStart > 0) And (PEnd >= PStart) Then ErrExtractFromReport = Mid$(Rpt, PStart, PEnd - PStart + 1)
End Function

'Message Source object ==========================================
Public Property Get ErrMessageSource() As Object
    Set ErrMessageSource = mMsgSrc
End Property

Public Property Set ErrMessageSource(ByVal Obj As Object)
    Set mMsgSrc = Obj
End Property

'Objects Release ========================================================
Public Sub ErrRlsObjs(Optional ByRef Obj1 As Object, _
                            Optional ByRef Obj2 As Object, _
                            Optional ByRef Obj3 As Object, _
                            Optional ByRef Obj4 As Object, _
                            Optional ByRef Obj5 As Object, _
                            Optional ByRef Obj6 As Object, _
                            Optional ByRef Obj7 As Object, _
                            Optional ByRef Obj8 As Object)
    ErrPreserve
    On Error GoTo EndProc
    mRlsdObjs = ""
    ReleaseObj Obj1
    ReleaseObj Obj2
    ReleaseObj Obj3
    ReleaseObj Obj4
    ReleaseObj Obj5
    ReleaseObj Obj6
    ReleaseObj Obj7
    ReleaseObj Obj8
EndProc:
    ErrRestore
    'Nothing to do
End Sub

Public Sub ErrCloseFiles(ParamArray Files())
    Dim i
    ErrPreserve
    On Error Resume Next
    For i = LBound(Files) To UBound(Files)
        If Files(i) <> 0 Then Close Files(i)
    Next i
End Sub

Public Function ErrMustStop() As Boolean
    Dim Res As Long
    Const STR_STOPMSG = "Stopped on error: $Descr$||"
    ErrPreserve
    If ErrInIDE And (Not InException) And ErrStopFlag And mErrHelpFile <> STR_NOSTOP Then
        Debug.Print
        Debug.Print Format(Now, "hh:nn:ss") & " Stopped on error:" & vbNewLine & ErrDescription
        Select Case MsgBox(ErrStopPrompt, vbYesNoCancel Or vbCritical, "Stopped on Error")
            Case vbYes: ErrMustStop = True: mErrPreserved = False 'must clear this flag
            Case vbNo: ErrMustStop = False
            Case vbCancel: ErrMustStop = False: mErrHelpFile = STR_NOSTOP
        End Select
        ErrRestore
    End If
End Function

Private Property Get ErrStopPrompt() As String
    ErrStopPrompt = "Error: " & ErrOrigDescription & IIf(InPropagation, " (Propagated)", "") & vbNewLine & _
        "Do you want to retry the operation in step mode?" & vbNewLine & _
        "Click YES to retry, NO to move to the caller, CANCEL for no more stops"
End Property

Public Property Get ErrStopFlag() As Boolean
#If H_NOSTOP = 1 Then
    ErrStopFlag = False
#Else
    ErrStopFlag = True
#End If
End Property

'Logs error report to system event log
'Logging is ignored from within VB IDE!
Public Function ErrSaveToEventLog() As Boolean
    On Error GoTo errHandler
    If ErrReport <> "" Then
        App.StartLogging "", vbLogToNT
        App.LogEvent ErrReport
    End If 'mErrReport...
    ErrSaveToEventLog = True
    Exit Function
errHandler:
    'nothing to do...
End Function

'Appends error report to text file
Public Function ErrSaveToFile(Optional ByVal ErrFileName As String = "Errors.txt") As Boolean
    Dim F As Long, FName As String
    On Error GoTo errHandler
    If ErrReport <> "" Then
        F = FreeFile
        FName = IIf(InStr(1, ErrFileName, "\") > 0, ErrFileName, App.Path & "\" & ErrFileName)
        If ErrOpenErrorFile(FName, F) Then
            Print #F, ErrReport & vbNewLine & vbNewLine
            Close #F
        End If
    End If 'mErrReport...
    ErrSaveToFile = True
    Exit Function
errHandler:
    'nothing to do...
End Function

'Logs error report to database table
Public Function ErrSaveToDB(ByVal ConnectString As String, _
                       Optional ByVal AppID As Long = 1, _
                       Optional ByVal ProcName As String = "spErrorLogInsert") As Boolean
    Dim Cmd As Object, SQL As String
    On Error GoTo errHandler
    If ErrReport <> "" Then
        'SQL Server:
        SQL = "Exec " & ProcName & " " & AppID & ", " & "'" & Replace(ErrReport, "'", "''") & "'"
        'Oracle:
        'SQL = "Call " & ProcName & " (" & AppID & ", " & "'" & Replace(ErrReport, "'", "''") & "')"
        Set Cmd = CreateObject("ADODB.Command")
        Cmd.CommandType = 1 ' adCmdText
        Cmd.CommandText = SQL
        Cmd.ActiveConnection = ConnectString
        Cmd.Execute
    End If
    ErrSaveToDB = True
    Exit Function
errHandler:
    App.StartLogging "", vbLogToNT 'Try to save message to Event Log
    App.LogEvent "Failed to save error report to database. " & vbNewLine & _
            "ConnectString = '" & ConnectString & "'"
End Function

Private Function ErrInsInto(ByVal IntoStr As String, ByVal Values) As String
    Dim i As Long, ChrX As String
    On Error Resume Next
    ChrX = Chr$(vbKeyBack) 'Spec char to act instead of % during manipulations
    IntoStr = Replace(IntoStr, "%", ChrX)
    If Not IsArray(Values) Then Values = Array(Values)
    For i = LBound(Values) To UBound(Values)
        IntoStr = Replace(IntoStr, ChrX & (i + 1), SafeStr(Values(i)))
    Next i
    ErrInsInto = Replace(IntoStr, ChrX, "%") 'replace back
End Function

'This version of IDE detection was submitted to PSC by Dan F
Public Function ErrInIDE() As Boolean
    Dim boolVar As Boolean
    Debug.Assert SetToTrue(boolVar)
    ErrInIDE = boolVar
End Function

Private Function SetToTrue(ByRef boolVar As Boolean) As Boolean
    boolVar = True
    SetToTrue = True
End Function

'##################################### Private methods #########################################
Private Sub ReleaseObj(ByRef AnObj As Object)
    On Error GoTo EndProc
    If Not AnObj Is Nothing Then
        If mRlsdObjs <> "" Then mRlsdObjs = mRlsdObjs & ", "
        mRlsdObjs = mRlsdObjs & "[" & TypeName(AnObj) & "]"
        If (mSavedConn Is Nothing) Then
            If InStr(1, "Connection,Command,Recordset", TypeName(AnObj)) > 0 Then _
                Set mSavedConn = ErrGetConnectionObject(AnObj, True)
        End If '(mSavedConn...
    End If 'Not AnObj...
EndProc:
    Set AnObj = Nothing
End Sub

'Prepares initial report
Private Sub ReportInit(ByVal MethodName As String)
Const TEMPLATE_REPORT = _
  "%Descr% %nl%" & _
  "  Time='%Time%' App='%App%:%Ver%' ADO-version='%ADOVersion%' Computer='%Comp%' %nl%" & _
  "  Method: %MethodName% %nl%" & _
  "  Number: %ErrNum% = &H%ErrHex% = vbObjectError %ErrNumRel1% = ERRMAP_APP_FIRST %ErrNumRel2% %ErrStd%%nl%" & _
  "  Source: %Source% %nl%" & _
  "  Description: %Descr%%nl%"
    On Error GoTo errHandler
    mErrReport = Replace(TEMPLATE_REPORT, "%", ChrBk)
    ReportSet "nl", vbNewLine
    ReportSet "MethodName", MethodName
    ReportSet "Comp", ErrGetComputerName
    ReportSet "Time", Format(Now, "Mm/Dd/yy Hh:Nn:Ss")
    ReportSet "App", App.EXEName
    ReportSet "Ver", ErrGetAppVersion
    ReportSet "ADOVersion", ErrGetADOVersion
    ReportSet "ErrNum", mErrNumber
    ReportSet "ErrHex", Hex$(mErrNumber)
    ReportSet "ErrNumRel1", FormatNum(mErrNumber - vbObjectError)
    ReportSet "ErrNumRel2", FormatNum(mErrNumber - ERRMAP_APP_FIRST)
    ReportSet "ErrStd", IIf(mErrNumber = ERR_GENERAL, "= ERR_GENERAL", "")
    ReportSet "Source", mErrSource
    ReportSet "Descr", mErrDescr
    mErrReport = Replace(mErrReport, ChrBk, "%")
    Exit Sub
errHandler:
    'Nothing to do...
End Sub

Private Function ErrGetADOVersion() As String
    ErrGetADOVersion = "?"
    On Error Resume Next
    ErrGetADOVersion = CreateObject("ADODB.Connection").version
End Function

Private Sub ReportAddAPIError()
    On Error GoTo errHandler
    '203 is "System cannot find the environment", not so much meaningful
    If (mErrAPI <> 0) And (mErrAPI <> 203) Then
        ReportAdd "  API Error: (" & mErrAPI & ") " & FormatMessage(mErrAPI)
    End If
    Exit Sub
errHandler:
    'Nothing to do...
End Sub

Private Sub ReportAddExtErrors()
    On Error GoTo errHandler
    If mExtBuffer <> "" Then
        mErrReport = mErrReport & mExtBuffer 'Ext buffer must already have vbNewLine at the end.
        mExtBuffer = ""
    End If
    Exit Sub
errHandler:
    'Nothing to do...
End Sub

Private Sub ReportAddRlsdObjsList()
    If mRlsdObjs <> "" Then
        ReportAdd "    Released Objects: " & mRlsdObjs
        mRlsdObjs = ""
    End If
End Sub

Private Function FormatMessage(ByVal ErrNum As String) As String
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
    Dim strBuffer As String * 512, strMsg As String
    On Error GoTo errHandler
    FormatMessageAPI FORMAT_MESSAGE_FROM_SYSTEM, Null, ErrNum, 0, strBuffer, 512, 0
    strMsg = strBuffer
    'Strange but necessary manipulations
    strMsg = Replace(strMsg, vbNewLine, "")
    strMsg = Replace(strMsg, Chr(0), "")
    FormatMessage = strMsg
    Exit Function
errHandler:
    'Nothing to do...
End Function

Private Sub ErrTerminateTrans(ByVal ErrorAction As Long, _
                              ByVal objConn As Object, _
                              ByVal TransControlObject As Object)
    Dim Ctx As Object, strMsg As String
    On Error GoTo errHandler
    Set Ctx = ErrGetContext
    If Not TransControlObject Is Nothing Then
        strMsg = "Attempt to call TransControlObject.SetAbort: " & ErrSafeCallMethod(TransControlObject, "SetAbort")
    ElseIf (Not Ctx Is Nothing) And FlagSet(ErrorAction, EA_SET_ABORT Or EA_DISABLE_COMMIT) Then
        If FlagSet(ErrorAction, EA_SET_ABORT) Then
            strMsg = "Attempt to call ObjectContext.SetAbort " & ErrSafeCallMethod(Ctx, "SetAbort")
            Else
            strMsg = "Attempt to call ObjectContext.DisableCommit " & ErrSafeCallMethod(Ctx, "DisableCommit")
        End If
    ElseIf FlagSet(ErrorAction, EA_ROLLBACK) And Not (objConn Is Nothing) Then
        If objConn.State = 0 Then
            strMsg = "Could not call RollbackTrans: Connection is closed "
            Else
            strMsg = "Attempt to call RollbackTrans " & ErrSafeCallMethod(objConn, "RollbackTrans")
        End If
    End If 'Not TransControlObject ....
    If strMsg <> "" Then ReportAdd "    Transaction: " & strMsg
    Exit Sub
errHandler:
    'Nothing to do...
End Sub

Private Sub ErrConnClose(ByVal objConn As Object)
    On Error Resume Next
    If Not (objConn Is Nothing) Then
        If objConn.State <> 0 Then
            ReportAdd "    Connection: Attempt to call Close " & ErrSafeCallMethod(objConn, "Close")
        End If
    End If 'Not (objConn....
End Sub

Private Function ErrSafeCallMethod(ByVal Obj As Object, ByVal Method As String) As String
    On Error Resume Next
    Select Case Method
        Case "Close":          Obj.Close
        Case "DisableCommit":  Obj.DisableCommit
        Case "SetAbort":       Obj.SetAbort
        Case "RollbackTrans":  Obj.RollbackTrans
        Case Else: Err.Raise ERR_GENERAL, , "Unknown method"
    End Select
    ErrSafeCallMethod = IIf(Err.Number = 0, "succeeded", "failed with error '" & Err.Description & "'")
End Function

Private Function ErrGetConnectionObject(ByVal DBObject As Object, Optional ByVal ClearProp As Boolean) As Object
    On Error GoTo errHandler
    If DBObject Is Nothing Then
        If Not mSavedConn Is Nothing Then Set ErrGetConnectionObject = mSavedConn
        Else
        Select Case TypeName(DBObject)
            Case "Connection":  Set ErrGetConnectionObject = DBObject
            Case "Command", "Recordset":
                Set ErrGetConnectionObject = DBObject.ActiveConnection
                If ClearProp Then Set DBObject.ActiveConnection = Nothing
            Case Else:
                Set ErrGetConnectionObject = DBObject.Connection 'Custom class, try to get its Connection property
                If ClearProp Then Set DBObject.Connection = Nothing
        End Select
    End If
    Exit Function
errHandler:
End Function

Private Sub ReportAddADOInfo(ByVal objConn As Object, ByVal DBObject As Object)
    Dim E As Object, strState As String
    On Error GoTo errHandler
    If objConn Is Nothing Then Exit Sub
    ReportAdd "  ADO Info: "
    ReportAdd "    ADO Version:   " & ErrGetADOVersion
    ReportAdd "    DbObject:      " & TypeName(DBObject) & IIf(mSavedConn Is Nothing, "", _
        " (Connection object was preserved internally)")
    ReportAdd "    Conn. String: '" & objConn.ConnectionString & "'"
    ReportAdd "    Conn. State:   " & ConnStateAsString(objConn.State)
    If objConn.Errors.Count = 0 Then Exit Sub
    For Each E In objConn.Errors
        ReportAdd "    Error:         " & E.Description
    Next E
    Exit Sub
errHandler:
    'Nothing to do: failed for whatever reason, so no ADO errors
End Sub

'Reads information from IIS request object
Private Sub ReportAddWebInfo()
Const TEMPLATE_WEBINFO = _
  "%nl%" & _
  "  Web Info: %nl%" & _
  "    RequestMethod='%RequestMethod%'%nl%" & _
  "    QueryString: '%WebServer%%URL%%QS%' %nl%" & _
  "    FormData:    '%FormData%' %nl%" & _
  "    Cookies:     '%Cookies%'  %nl%"
    Dim Ctx As Object, IISRequestObj As Object
    Dim strReqMethod As String, strCookies As String, strServer As String
    On Error GoTo errHandler
    'Try to get Request object through ObjectContext
    Set Ctx = ErrGetContext
    If IsEmpty(Ctx) Or (Ctx Is Nothing) Then Exit Sub
    Set IISRequestObj = Ctx("Request")
    If IISRequestObj Is Nothing Then Exit Sub
    ReportAdd Replace(TEMPLATE_WEBINFO, "%", ChrBk)
    With IISRequestObj
        ReportSet "WebServer", .ServerVariables("SERVER_NAME")
        ReportSet "RequestMethod", .ServerVariables("REQUEST_METHOD")
        ReportSet "URL", .ServerVariables("URL")
        ReportSet "QS", IIf(.queryString = "", "", "?" & .queryString)
        ReportSet "FormData", CStr(.Form)
        ReportSet "Cookies", .Cookies
        ReportSet "nl", vbNewLine
    End With
    mErrReport = Replace(mErrReport, ChrBk, "%")
    Exit Sub
errHandler:
End Sub

Private Sub ParseMethodHeader(ByVal MethodHeader As String, ByRef MethodName As String, _
                ByRef ArgNames As String)
    Dim arrBuf() As String
    On Error GoTo errHandler
    arrBuf = Split(MethodHeader, "(")
    If UBound(arrBuf) >= 0 Then MethodName = arrBuf(0) Else MethodName = ""
    If UBound(arrBuf) <= 0 Then ArgNames = "" Else ArgNames = Left$(arrBuf(1), Len(arrBuf(1)) - 1)    'get rid of ")"
    Exit Sub
errHandler:
End Sub

Private Sub ReportAddCallStackInfo(ByVal MethodName As String, _
                                 ByVal ArgNames As String, _
                                 ByVal arrArgs, _
                                 ByVal EnvVarNames As String, _
                                 ByVal arrEnvVars)
    Dim S As String
    On Error GoTo errHandler
    S = "  Call Stack: " & MethodName & "(" & ErrCreateNameValueList(ArgNames, arrArgs) & ")" _
        & IIf(mErl = 0, "", "  at Line " & mErl) & " "
    ReportAdd Pad(S, "-", 100)
    If mLDBuffer <> "" Then ReportAdd mLDBuffer: mLDBuffer = ""
    If EnvVarNames <> "" Then ReportAdd "    Env: " & ErrCreateNameValueList(EnvVarNames, arrEnvVars)
    If mLDBuffer <> "" Then ReportAdd mLDBuffer: mLDBuffer = ""
    Exit Sub
errHandler:
End Sub

Private Function ErrCreateNameValueList(ByVal strNames As String, ByVal arrValues) As String
    Dim arrNames() As String, i As Long, strList As String
    Dim strName As String, strValue As String, strNameValue As String
    On Error GoTo errHandler
    mLDBuffer = ""
    'arrValues maybe array of values, or a single value.
    If Not IsArray(arrValues) Then arrValues = Array(arrValues)
    arrNames = Split(strNames, ",")
    For i = 0 To UBound(arrValues)
        If i <= UBound(arrNames) Then strName = arrNames(i) Else strName = ""
        strValue = ErrVarToString(arrValues(i))
        If Len(strValue) > MAX_NON_LONG_DATA Or InStr(1, strValue, vbNewLine) > 0 Then
            If (Left$(strName, 3) = "xml") And (Left(strValue, 2) = "'<") Then strValue = ErrFormatXML(strValue)
            strValue = Space(6) & Replace(strValue, vbNewLine, vbNewLine & Space(6)) 'Make indent
            mLDBuffer = mLDBuffer & IIf(mLDBuffer = "", "", vbNewLine) & "    Value Of " & strName & ":" & vbNewLine & strValue
            strValue = "{Text}"
        End If
        strNameValue = IIf(strName = "", strValue, strName & "=" & strValue)
        If strList <> "" Then strList = strList & ", "
        strList = strList & strNameValue
    Next i
    ErrCreateNameValueList = strList
    Exit Function
errHandler:
End Function

Private Function ErrFormatXML(ByVal xml As String) As String
    Dim arrTmp() As String, NestLvl As Long, NewLvl As Long, i As Long
    On Error GoTo errHandler
    xml = Mid$(xml, 2, Len(xml) - 2) 'Strip off quotes
    ErrFormatXML = xml
    arrTmp = Split(xml, "<") ' break into segments
    For i = 1 To UBound(arrTmp) 'arrTmp(0) should be empty string, just ignore it
        If Left(arrTmp(i), 1) = "/" Then
            NestLvl = NestLvl - 1 'This is closing tag, it belongs to upper level
            NewLvl = NestLvl
        ElseIf InStr(1, arrTmp(i), "/>") > 0 Then
            'This is opening tag, but it is closed in this line, so don't change nest level
        Else
            NewLvl = NestLvl + 1 'This is opening tag, inc nest level for followers
        End If
        arrTmp(i) = IIf(i > 1, vbNewLine, "") & Space(NestLvl * 2) & "<" & arrTmp(i)
        NestLvl = NewLvl
    Next i
    ErrFormatXML = Join(arrTmp, "")
    Exit Function
errHandler:
End Function

 'Utilities =========================================================================
Private Function FlagSet(ByVal Value As Long, ByVal Flag As Long) As Boolean
    FlagSet = ((Value And Flag) <> 0)
End Function

Private Sub ReportSet(ByVal Tag As String, ByVal Value As String)
    mErrReport = Replace(mErrReport, ChrBk & Tag & ChrBk, Value)
End Sub

Private Sub ReportAdd(ByVal Info As String)
    If Not InException Then mErrReport = mErrReport & Info & vbNewLine
End Sub

Private Function ErrGetComputerName() As String
    Dim sBuffer As String * 255, lLen As Long
    lLen = Len(sBuffer)
    If CBool(GetComputerNameAPI(sBuffer, lLen)) Then ErrGetComputerName = Left$(sBuffer, lLen)
End Function

Private Function ErrGetAppVersion() As String
    ErrGetAppVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

Private Function ErrVarToString(ByVal V) As String
  Dim L As Long, U As Long
  On Error GoTo errHandler
  If IsArray(V) Then
        ErrVarToString = "{Array}"
  Else 'If IsArray(...
    Select Case VarType(V)
        Case vbInteger, vbLong, vbByte, _
             vbSingle, vbDouble, vbCurrency, _
             vbBoolean, vbDecimal: ErrVarToString = CStr(V)
        Case vbDate:      ErrVarToString = "'" & CStr(V) & "'"
        Case vbError:     ErrVarToString = "" 'Missing arg falls here
        Case vbEmpty:     ErrVarToString = "{Empty}"
        Case vbNull:      ErrVarToString = "{Null}"
        Case vbString:    ErrVarToString = "'" & V & "'"
        Case vbObject:    ErrVarToString = "{" & TypeName(V) & "}" 'Value of Nothing will be shown as "Nothing"
        Case Else:        ErrVarToString = "{?}"
        End Select
    End If 'IsArray...
  Exit Function
errHandler:
  ErrVarToString = "{?}"
  End Function
 
Private Function FormatNum(ByVal L As Long) As String
    FormatNum = IIf(L >= 0, "+ " & L, "- " & Abs(L))
End Function

'File may be opened by other component. Keep trying for MAX_RETRY to open.
Private Function ErrOpenErrorFile(ByVal FileName As String, ByVal F As Long) As Boolean
    Dim StartTime As Date
    On Error GoTo errHandler
    ErrOpenErrorFile = True
    StartTime = Now
    Do
      If ErrTryOpenErrorFile(FileName, F) Then Exit Function
      APISleep 200
    Loop Until (Now - StartTime) > MAX_RETRY
    ErrOpenErrorFile = False
    Exit Function
errHandler:
    ErrOpenErrorFile = False
End Function

Private Function ErrTryOpenErrorFile(ByVal FileName As String, ByVal F As Long) As Boolean
    On Error Resume Next
    Open FileName For Append As #F
    ErrTryOpenErrorFile = (Err.Number = 0)
End Function

Private Function ErrGetContext() As Object
    'If this function doesn't compile do one of the following
    ' 1) Add reference to "COM+ Services Type Library" in Project|References box.
    ' 2) If you are not using COM+ in your project then add definition
    '       H_NOCOMPLUS=1 to "Conditional compilation arguments" box on Make page of
    '       Project Properties dialog.
#If H_NOCOMPLUS = 1 Then
    Set ErrGetContext = Nothing
#Else
    ' 3) or just comment the following line
    Set ErrGetContext = GetObjectContext
#End If
End Function

Private Function ConnStateAsString(ByVal AState As Long) As String
    Dim sState As String
    On Error GoTo errHandler
    If AState = 0 Then
        sState = "adStateClosed"
    Else
        If FlagSet(AState, 1) Then sState = "adStateOpen"
        If FlagSet(AState, 2) Then sState = sState & " + adStateConnecting"
        If FlagSet(AState, 4) Then sState = sState & " + adStateExecuting"
        If FlagSet(AState, 8) Then sState = sState & " + adStateFetching"
    End If
    ConnStateAsString = sState
    Exit Function
errHandler:
End Function

Private Function SafeStr(ByVal V) As String
    On Error Resume Next
    SafeStr = CStr(V)
End Function

'Special char to be used in string manipulations instead of %, to avoid substituting in already replaced values
Private Property Get ChrBk() As String
    ChrBk = Chr$(vbKeyBack)
End Property

Private Function Indent(ByVal Src As String, ByVal NumSp As Long) As String
    Indent = Space(NumSp) & Replace(Src, vbNewLine, vbNewLine & Space(NumSp))
End Function

Private Function Unindent(ByVal Src As String) As String
    While InStr(1, Src, vbNewLine & " ") > 0
        Src = Replace(Src, vbNewLine & " ", vbNewLine)
    Wend
    Unindent = Src
End Function

Private Function Pad(ByVal Src As String, ByVal Char As String, ByVal ToLen As Long) As String
    If Len(Src) < ToLen Then
        Pad = Src & Replace(Space(ToLen - Len(Src)), " ", Char)
        Else
        Pad = Src
    End If
End Function

