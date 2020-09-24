Attribute VB_Name = "HuntERR31SysH"
'========================================================================================
'                               HuntERR
'                   Error Handling and Reporting Library
'                       System exceptions handling module
'                        from URFIN JUS (www.urfinjus.net)
'                      Copyright 2001. All rights reserved.
'version 3.1 Preview, 01/12/2002
'=========================================================================================
'System exception handling allows to prevent application from crashing in case of serious
'errors like access violation.
'System exception handling is described in details in excellent article
'"No Exception Errors, My Dear Dr. Watson" by Jonathan Lunman
'May 99, Visual Basic Programmer's Journal (www.vbpj.com)

'Public members
'Public Property Get ErrSysHandlerWasSet() As Boolean
'Public Sub ErrSysHandlerSet()
'Public Sub ErrSysHandlerRelease()

Option Explicit

Public Enum ENUM_SYSEXC
    SYSEXC_ACCESS_VIOLATION = &HC0000005
    SYSEXC_DATATYPE_MISALIGNMENT = &H80000002
    SYSEXC_BREAKPOINT = &H80000003
    SYSEXC_SINGLE_STEP = &H80000004
    SYSEXC_ARRAY_BOUNDS_EXCEEDED = &HC000008C
    SYSEXC_FLT_DENORMAL_OPERAND = &HC000008D
    SYSEXC_FLT_DIVIDE_BY_ZERO = &HC000008E
    SYSEXC_FLT_INEXACT_RESULT = &HC000008F
    SYSEXC_FLT_INVALID_OPERATION = &HC0000090
    SYSEXC_FLT_OVERFLOW = &HC0000091
    SYSEXC_FLT_STACK_CHECK = &HC0000092
    SYSEXC_FLT_UNDERFLOW = &HC0000093
    SYSEXC_INT_DIVIDE_BY_ZERO = &HC0000094
    SYSEXC_INT_OVERFLOW = &HC0000095
    SYSEXC_PRIVILEGED_INSTRUCTION = &HC0000096
    SYSEXC_IN_PAGE_ERROR = &HC0000006
    SYSEXC_ILLEGAL_INSTRUCTION = &HC000001D
    SYSEXC_NONCONTINUABLE_EXCEPTION = &HC0000025
    SYSEXC_STACK_OVERFLOW = &HC00000FD
    SYSEXC_INVALID_DISPOSITION = &HC0000026
    SYSEXC_GUARD_PAGE_VIOLATION = &H80000001
    SYSEXC_INVALID_HANDLE = &HC0000008
    SYSEXC_CONTROL_C_EXIT = &HC000013A
End Enum

Private mSysHandlerWasSet As Boolean

Public Const SYSEXC_MAXIMUM_PARAMETERS = 15
'Not exactly as in API, shorter declaration, but internally the same
Type CONTEXT
  Dbls(0 To 66) As Double
  Longs(0 To 6) As Long
End Type

Type SYSEXC_RECORD
    ExceptionCode As Long
    ExceptionFlags As Long
    pExceptionRecord As Long
    ExceptionAddress As Long
    NumberParameters As Long
    ExceptionInformation(SYSEXC_MAXIMUM_PARAMETERS) As Long
End Type

Type SYSEXC_DEBUG_INFO
        pExceptionRecord As SYSEXC_RECORD
        dwFirstChance As Long
End Type

Type SYSEXC_POINTERS
    pExceptionRecord As SYSEXC_RECORD
    ContextRecord As CONTEXT
End Type

Private Declare Function SetUnhandledExceptionFilter Lib "kernel32" _
    (ByVal lpTopLevelExceptionFilter As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub CopyExceptionRecord Lib "kernel32" Alias "RtlMoveMemory" (pDest As SYSEXC_RECORD, ByVal LPSYSEXC_RECORD As Long, ByVal lngBytes As Long)

Public Property Get ErrSysHandlerWasSet() As Boolean
    ErrSysHandlerWasSet = mSysHandlerWasSet
End Property

Public Sub ErrSysHandlerSet()
    If mSysHandlerWasSet Then ErrSysHandlerRelease
    Call SetUnhandledExceptionFilter(AddressOf SysExcHandler)
    mSysHandlerWasSet = True
End Sub

Public Sub ErrSysHandlerRelease()
    ErrPreserve 'This Sub may be called from error handler, so preserve errors
    On Error Resume Next
    If mSysHandlerWasSet Then Call SetUnhandledExceptionFilter(0)
    mSysHandlerWasSet = False
    ErrRestore
End Sub

'========================== Private stuff ===========================================
Private Function SysExcHandler(ByRef ExcPtrs As SYSEXC_POINTERS) As Long
  Dim ExcRec As SYSEXC_RECORD, strExc As String
  ExcRec = ExcPtrs.pExceptionRecord
  Do Until ExcRec.pExceptionRecord = 0
    CopyExceptionRecord ExcRec, ExcRec.pExceptionRecord, Len(ExcRec)
  Loop
  strExc = GetExcAsText(ExcRec.ExceptionCode)
  Err.Raise ERR_GENERAL, SRC_SYSHANDLER, _
    "(&H" & Hex$(ExcRec.ExceptionCode) & ") " & strExc
End Function

Private Function GetExcAsText(ByVal ExcNum As Long) As String
    Select Case ExcNum
        Case SYSEXC_ACCESS_VIOLATION:          GetExcAsText = "Access violation"
        Case SYSEXC_DATATYPE_MISALIGNMENT:     GetExcAsText = "Datatype misalignment"
        Case SYSEXC_BREAKPOINT:                GetExcAsText = "Breakpoint"
        Case SYSEXC_SINGLE_STEP:               GetExcAsText = "Single step"
        Case SYSEXC_ARRAY_BOUNDS_EXCEEDED:     GetExcAsText = "Array bounds exceeded"
        Case SYSEXC_FLT_DENORMAL_OPERAND:      GetExcAsText = "Float Denormal Operand"
        Case SYSEXC_FLT_DIVIDE_BY_ZERO:        GetExcAsText = "Divide By Zero"
        Case SYSEXC_FLT_INEXACT_RESULT:        GetExcAsText = "Floating Point Inexact Result"
        Case SYSEXC_FLT_INVALID_OPERATION:     GetExcAsText = "Invalid Operation"
        Case SYSEXC_FLT_OVERFLOW:              GetExcAsText = "Float Overflow"
        Case SYSEXC_FLT_STACK_CHECK:           GetExcAsText = "Float Stack Check"
        Case SYSEXC_FLT_UNDERFLOW:             GetExcAsText = "Float Underflow"
        Case SYSEXC_INT_DIVIDE_BY_ZERO:        GetExcAsText = "Integer Divide By Zero"
        Case SYSEXC_INT_OVERFLOW:              GetExcAsText = "Integer Overflow"
        Case SYSEXC_PRIVILEGED_INSTRUCTION:    GetExcAsText = "Privileged Instruction"
        Case SYSEXC_IN_PAGE_ERROR:             GetExcAsText = "In Page Error"
        Case SYSEXC_ILLEGAL_INSTRUCTION:       GetExcAsText = "Illegal Instruction"
        Case SYSEXC_NONCONTINUABLE_EXCEPTION:  GetExcAsText = "Non Continuable Exception"
        Case SYSEXC_STACK_OVERFLOW:            GetExcAsText = "Stack Overflow"
        Case SYSEXC_INVALID_DISPOSITION:       GetExcAsText = "Invalid Disposition"
        Case SYSEXC_GUARD_PAGE_VIOLATION:      GetExcAsText = "Guard Page Violation"
        Case SYSEXC_INVALID_HANDLE:            GetExcAsText = "Invalid Handle"
        Case SYSEXC_CONTROL_C_EXIT:            GetExcAsText = "Control-C Exit"
    End Select
End Function


