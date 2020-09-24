VERSION 5.00
Begin VB.Form frmShowError 
   Caption         =   "Error Report"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   12420
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtError 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5520
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   12300
   End
End
Attribute VB_Name = "frmShowError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================
'                               HuntERR
'                   Error Handling and Reporting Library
'                    from URFIN JUS (www.urfinjus.net)
'                  Copyright 2001-2002. All rights reserved.
'version 3.1, 04/25/2002
'Simple window to show error report
'=========================================================================================
Option Explicit

Public Property Let ErrorReport(ByVal AReport As String)
    txtError.Text = AReport
    If Not Visible Then ShowSelf
    On Error Resume Next
    Me.SetFocus
End Property

'Try to show non-modal. If there is already modal window in application,
'then it will fail, and we'll show as modal.
Private Sub ShowSelf()
    On Error Resume Next
    Show
    If Err.Number <> 0 Then Show vbModal
End Sub

Private Function ShowNormal()
End Function

Private Sub Form_Resize()
  txtError.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

