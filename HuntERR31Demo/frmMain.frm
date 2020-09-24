VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HuntERR 3.1 Demo Application"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrevPage 
      Caption         =   "Previous"
      Height          =   330
      Left            =   6945
      TabIndex        =   33
      Top             =   4635
      Width           =   1200
   End
   Begin VB.CheckBox chkStopInProc 
      Caption         =   "Check this box to step into code"
      Height          =   255
      Left            =   165
      TabIndex        =   32
      Top             =   4680
      Width           =   3375
   End
   Begin VB.CommandButton cmdNextPage 
      Caption         =   "Next"
      Height          =   330
      Left            =   8265
      TabIndex        =   34
      Top             =   4620
      Width           =   1200
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   3975
      Left            =   45
      TabIndex        =   35
      Top             =   555
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   15
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "1. DB Connection"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frames(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "2. Error Log"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "3. HuntERR Basics"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frames(0)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "4. Loss of Err Info"
      TabPicture(3)   =   "frmMain.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frames(12)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "5. Long Strings"
      TabPicture(4)   =   "frmMain.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "frames(7)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "6. API Errors"
      TabPicture(5)   =   "frmMain.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "frames(4)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "7. System Exc."
      TabPicture(6)   =   "frmMain.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "frames(5)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "8. Extensions"
      TabPicture(7)   =   "frmMain.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame5"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "9. XML Formatting"
      TabPicture(8)   =   "frmMain.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "frames(9)"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "10. Msg Source"
      TabPicture(9)   =   "frmMain.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "frames(8)"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "11. Accumulation"
      TabPicture(10)  =   "frmMain.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "frames(1)"
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "12. Obj Release"
      TabPicture(11)  =   "frmMain.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "frames(6)"
      Tab(11).ControlCount=   1
      TabCaption(12)  =   "13. ADO"
      TabPicture(12)  =   "frmMain.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "frames(3)"
      Tab(12).ControlCount=   1
      TabCaption(13)  =   "14. Stop On Error"
      TabPicture(13)  =   "frmMain.frx":016C
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "frames(10)"
      Tab(13).ControlCount=   1
      TabCaption(14)  =   "About"
      TabPicture(14)  =   "frmMain.frx":0188
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "frames(11)"
      Tab(14).ControlCount=   1
      Begin VB.Frame frames 
         ForeColor       =   &H00FF0000&
         Height          =   2895
         Index           =   11
         Left            =   -74880
         TabIndex        =   85
         Top             =   960
         Width           =   9000
         Begin VB.Label Label44 
            Caption         =   "www.urfinjus.net"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   180
            TabIndex        =   95
            Top             =   1560
            Width           =   1635
         End
         Begin VB.Label Label43 
            Caption         =   "Copyright URFIN JUS, All rights reserved."
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   195
            TabIndex        =   94
            Top             =   1215
            Width           =   4020
         End
         Begin VB.Label Label39 
            Caption         =   "Demo Application"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   165
            TabIndex        =   93
            Top             =   840
            Width           =   2715
         End
         Begin VB.Label Label35 
            Caption         =   "Version 3.1"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   180
            TabIndex        =   92
            Top             =   510
            Width           =   2715
         End
         Begin VB.Label Label33 
            Caption         =   "HuntERR, Error Handling Solution for Visual Basic"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   165
            TabIndex        =   91
            Top             =   210
            Width           =   7005
         End
      End
      Begin VB.Frame frames 
         ForeColor       =   &H00FF0000&
         Height          =   2805
         Index           =   10
         Left            =   -74865
         TabIndex        =   82
         Top             =   1035
         Width           =   8985
         Begin VB.CommandButton btnStopOnErrExecute 
            Caption         =   "Execute"
            Height          =   315
            Left            =   7605
            TabIndex        =   31
            Top             =   330
            Width           =   1230
         End
         Begin VB.Label Label36 
            Caption         =   "New in 3.1: HuntERR provides new functions that allow to implement 'immediate stop on error' functionality."
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   120
            TabIndex        =   84
            Top             =   2325
            Width           =   8520
         End
         Begin VB.Label Label34 
            Caption         =   $"frmMain.frx":01A4
            Height          =   735
            Left            =   270
            TabIndex        =   83
            Top             =   300
            Width           =   6780
         End
      End
      Begin VB.Frame frames 
         ForeColor       =   &H00FF0000&
         Height          =   2790
         Index           =   3
         Left            =   -74880
         TabIndex        =   76
         Top             =   1005
         Width           =   8940
         Begin VB.CheckBox chkADOAbortTrans 
            Caption         =   "Abort transaction"
            Height          =   285
            Left            =   4920
            TabIndex        =   27
            Top             =   885
            Width           =   1575
         End
         Begin VB.CheckBox chkADOAutoRelease 
            Caption         =   "Release Connection object "
            Height          =   285
            Left            =   4920
            TabIndex        =   29
            Top             =   1440
            Width           =   2340
         End
         Begin VB.CheckBox chkADOAutoClose 
            Caption         =   "Close Connection to database"
            Height          =   315
            Left            =   4920
            TabIndex        =   28
            Top             =   1155
            Width           =   2625
         End
         Begin VB.CheckBox chkADOInTrans 
            Caption         =   "Execute in transaction"
            Height          =   285
            Left            =   285
            TabIndex        =   26
            Top             =   915
            Width           =   2355
         End
         Begin VB.TextBox txtSQL 
            Height          =   300
            Left            =   255
            TabIndex        =   25
            Text            =   "SELECT x,y,z  FROM Customers"
            Top             =   555
            Width           =   7005
         End
         Begin VB.CommandButton cmdExecuteSQL 
            Caption         =   "Execute"
            Height          =   330
            Left            =   7335
            TabIndex        =   30
            Top             =   540
            Width           =   1215
         End
         Begin VB.Label Label28 
            Caption         =   "Post-execution status:"
            Height          =   225
            Left            =   195
            TabIndex        =   81
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label lblADOStatus 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1890
            TabIndex        =   80
            Top             =   1785
            Width           =   5580
         End
         Begin VB.Label Label27 
            Caption         =   "ON ERROR:"
            Height          =   255
            Left            =   3750
            TabIndex        =   79
            Top             =   915
            Width           =   945
         End
         Begin VB.Label Label15 
            Caption         =   $"frmMain.frx":02BB
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   195
            TabIndex        =   78
            Top             =   2220
            Width           =   8400
         End
         Begin VB.Label Label3 
            Caption         =   "Enter invalid SQL statement to generate ADO/database errors"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   210
            TabIndex        =   77
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.Frame frames 
         ForeColor       =   &H00FF0000&
         Height          =   2805
         Index           =   6
         Left            =   -74880
         TabIndex        =   73
         Top             =   1005
         Width           =   8955
         Begin VB.CommandButton btnExecuteRlsObjs 
            Caption         =   "Execute"
            Height          =   315
            Left            =   7530
            TabIndex        =   24
            Top             =   285
            Width           =   1320
         End
         Begin VB.Label Label22 
            Caption         =   $"frmMain.frx":0377
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   270
            TabIndex        =   75
            Top             =   2010
            Width           =   8010
         End
         Begin VB.Label Label20 
            Caption         =   $"frmMain.frx":0466
            Height          =   810
            Left            =   240
            TabIndex        =   74
            Top             =   240
            Width           =   7170
         End
      End
      Begin VB.Frame frames 
         Height          =   2805
         Index           =   1
         Left            =   -74790
         TabIndex        =   67
         Top             =   1005
         Width           =   8865
         Begin VB.TextBox txtAccumName 
            Height          =   285
            Left            =   1875
            TabIndex        =   20
            Top             =   960
            Width           =   1725
         End
         Begin VB.TextBox txtAccumSSN 
            Height          =   285
            Left            =   1875
            TabIndex        =   21
            Text            =   "???"
            Top             =   1245
            Width           =   1725
         End
         Begin VB.TextBox txtAccumPhone 
            Height          =   285
            Left            =   1875
            TabIndex        =   22
            Text            =   "???"
            Top             =   1560
            Width           =   1725
         End
         Begin VB.CommandButton cmdAccumValidate 
            Caption         =   "Validate"
            Height          =   330
            Left            =   3720
            TabIndex        =   23
            Top             =   930
            Width           =   1335
         End
         Begin VB.Label Label16 
            Caption         =   "Name:"
            Height          =   270
            Left            =   1005
            TabIndex        =   72
            Top             =   945
            Width           =   645
         End
         Begin VB.Label Label18 
            Caption         =   "SSN:"
            Height          =   270
            Left            =   1020
            TabIndex        =   71
            Top             =   1305
            Width           =   540
         End
         Begin VB.Label Label19 
            Caption         =   "Phone:"
            Height          =   270
            Left            =   1005
            TabIndex        =   70
            Top             =   1620
            Width           =   540
         End
         Begin VB.Label Label17 
            Caption         =   "New in 3.1: Error accumulation; parameterized message in Check sub "
            ForeColor       =   &H00FF0000&
            Height          =   450
            Left            =   210
            TabIndex        =   69
            Top             =   2145
            Width           =   8145
         End
         Begin VB.Label Label14 
            Caption         =   $"frmMain.frx":055D
            Height          =   570
            Left            =   225
            TabIndex        =   68
            Top             =   180
            Width           =   8430
         End
      End
      Begin VB.Frame frames 
         ForeColor       =   &H00FF0000&
         Height          =   2700
         Index           =   8
         Left            =   -74805
         TabIndex        =   61
         Top             =   1125
         Width           =   8865
         Begin VB.TextBox txtMsgPrm2 
            Height          =   300
            Left            =   1875
            TabIndex        =   18
            Text            =   "Other String"
            Top             =   1530
            Width           =   1575
         End
         Begin VB.TextBox txtMsgPrm1 
            Height          =   300
            Left            =   1860
            TabIndex        =   17
            Text            =   "Strange %2 string"
            Top             =   1170
            Width           =   1575
         End
         Begin VB.ComboBox cmbMsgSrcTestType 
            Height          =   315
            ItemData        =   "frmMain.frx":0644
            Left            =   1875
            List            =   "frmMain.frx":0654
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   795
            Width           =   5355
         End
         Begin VB.CommandButton cmdMsgSrcExecute 
            Caption         =   "Execute"
            Height          =   345
            Left            =   7245
            TabIndex        =   19
            Top             =   300
            Width           =   1455
         End
         Begin VB.Label Label41 
            Caption         =   "(Notice that %2 is not being replaced by second parameter)"
            Height          =   270
            Left            =   3675
            TabIndex        =   88
            Top             =   1185
            Width           =   4515
         End
         Begin VB.Label Label30 
            Caption         =   "Parameter #2:"
            Height          =   270
            Left            =   225
            TabIndex        =   66
            Top             =   1545
            Width           =   1455
         End
         Begin VB.Label Label29 
            Caption         =   $"frmMain.frx":0704
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   165
            TabIndex        =   65
            Top             =   2235
            Width           =   7560
         End
         Begin VB.Label Label26 
            Caption         =   "Parameter #1:"
            Height          =   270
            Left            =   210
            TabIndex        =   64
            Top             =   1200
            Width           =   1395
         End
         Begin VB.Label Label25 
            Caption         =   "Error/Exc Description:"
            Height          =   270
            Left            =   225
            TabIndex        =   63
            Top             =   825
            Width           =   1710
         End
         Begin VB.Label Label24 
            Caption         =   " When you press Execute button Demo App will raise exception using Check Sub with selected source of exception message"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   255
            TabIndex        =   62
            Top             =   225
            Width           =   6570
         End
      End
      Begin VB.Frame frames 
         ForeColor       =   &H00FF0000&
         Height          =   2745
         Index           =   9
         Left            =   -74865
         TabIndex        =   58
         Top             =   1050
         Width           =   8970
         Begin VB.CommandButton btnXMLFmtExecute 
            Caption         =   "Execute"
            Height          =   315
            Left            =   7410
            TabIndex        =   15
            Top             =   315
            Width           =   1485
         End
         Begin VB.Label Label32 
            Caption         =   $"frmMain.frx":079F
            ForeColor       =   &H00FF0000&
            Height          =   555
            Left            =   285
            TabIndex        =   60
            Top             =   2010
            Width           =   7260
         End
         Begin VB.Label Label9 
            Caption         =   $"frmMain.frx":0835
            Height          =   630
            Left            =   270
            TabIndex        =   59
            Top             =   300
            Width           =   6780
         End
      End
      Begin VB.Frame Frame5 
         ForeColor       =   &H00FF0000&
         Height          =   2805
         Left            =   -74865
         TabIndex        =   56
         Top             =   1035
         Width           =   9000
         Begin VB.CommandButton btnParseXML 
            Caption         =   "Parse XML"
            Height          =   330
            Left            =   6975
            TabIndex        =   14
            Top             =   765
            Width           =   1380
         End
         Begin VB.TextBox txtXML 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1020
            Left            =   345
            MultiLine       =   -1  'True
            TabIndex        =   13
            Text            =   "frmMain.frx":0924
            Top             =   795
            Width           =   6165
         End
         Begin VB.Label Label42 
            Caption         =   "New in 3.1: Extractor class doesn't need to care about indentation. HuntERR indents error message automatically"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   165
            TabIndex        =   89
            Top             =   2295
            Width           =   8415
         End
         Begin VB.Label Label8 
            Caption         =   $"frmMain.frx":0965
            Height          =   420
            Left            =   150
            TabIndex        =   57
            Top             =   270
            Width           =   7755
         End
      End
      Begin VB.Frame frames 
         ForeColor       =   &H00FF0000&
         Height          =   2805
         Index           =   5
         Left            =   -74880
         TabIndex        =   53
         Top             =   1035
         Width           =   8970
         Begin VB.CommandButton btnSysHandlerExecute 
            Caption         =   "Execute"
            Height          =   315
            Left            =   7380
            TabIndex        =   12
            Top             =   315
            Width           =   1350
         End
         Begin VB.CheckBox chkSetHandler 
            Caption         =   "Set System Exception Handler"
            Height          =   255
            Left            =   375
            TabIndex        =   11
            Top             =   1155
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.Label Label31 
            Caption         =   "New in 3.1: ErrorIn does not release system exception handler automatically. Application should do it explicitly"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   180
            TabIndex        =   55
            Top             =   2295
            Width           =   8415
         End
         Begin VB.Label Label7 
            Caption         =   $"frmMain.frx":09EC
            Height          =   645
            Left            =   255
            TabIndex        =   54
            Top             =   240
            Width           =   6735
         End
      End
      Begin VB.Frame frames 
         ForeColor       =   &H00FF0000&
         Height          =   2790
         Index           =   4
         Left            =   -74850
         TabIndex        =   51
         Top             =   1035
         Width           =   8970
         Begin VB.CommandButton btnTestAPI 
            Caption         =   "Execute"
            Height          =   315
            Left            =   7560
            TabIndex        =   10
            Top             =   225
            Width           =   1320
         End
         Begin VB.Label Label5 
            Caption         =   $"frmMain.frx":0AF7
            Height          =   465
            Left            =   270
            TabIndex        =   52
            Top             =   300
            Width           =   6780
         End
      End
      Begin VB.Frame frames 
         ForeColor       =   &H00FF0000&
         Height          =   2805
         Index           =   7
         Left            =   -74865
         TabIndex        =   48
         Top             =   1050
         Width           =   8970
         Begin VB.CommandButton cmdLDExecute 
            Caption         =   "Execute"
            Height          =   345
            Left            =   7410
            TabIndex        =   9
            Top             =   225
            Width           =   1395
         End
         Begin VB.Label Label21 
            Caption         =   "New in 3.1: Long string values of proc parameters and env. variables are shown separately in error report"
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   225
            TabIndex        =   50
            Top             =   2295
            Width           =   7740
         End
         Begin VB.Label Label23 
            Caption         =   $"frmMain.frx":0BA3
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   225
            TabIndex        =   49
            Top             =   270
            Width           =   6180
         End
      End
      Begin VB.Frame frames 
         ForeColor       =   &H00FF0000&
         Height          =   2820
         Index           =   12
         Left            =   -74895
         TabIndex        =   45
         Top             =   1035
         Width           =   8955
         Begin VB.CommandButton btnErrClearedExecute 
            Caption         =   "Execute"
            Height          =   315
            Left            =   7560
            TabIndex        =   8
            Top             =   225
            Width           =   1320
         End
         Begin VB.Label Label38 
            Caption         =   $"frmMain.frx":0C34
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   165
            TabIndex        =   47
            Top             =   2205
            Width           =   8415
         End
         Begin VB.Label Label37 
            Caption         =   $"frmMain.frx":0CC5
            Height          =   990
            Left            =   270
            TabIndex        =   46
            Top             =   210
            Width           =   6780
         End
      End
      Begin VB.Frame frames 
         ForeColor       =   &H00FF0000&
         Height          =   2790
         Index           =   0
         Left            =   -74850
         TabIndex        =   42
         Top             =   1035
         Width           =   8955
         Begin VB.TextBox txtX 
            Height          =   285
            Left            =   1095
            TabIndex        =   6
            Text            =   "1"
            Top             =   945
            Width           =   870
         End
         Begin VB.CommandButton cmdCalc 
            Caption         =   "Calculate 1/(1-x)"
            Height          =   345
            Left            =   5835
            TabIndex        =   7
            Top             =   900
            Width           =   1725
         End
         Begin VB.Label Label40 
            Caption         =   "New in 3.1: Parameterized error/exception description parameter of Check Sub"
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   180
            TabIndex        =   86
            Top             =   2250
            Width           =   8415
         End
         Begin VB.Label Label1 
            Caption         =   "X = "
            Height          =   255
            Left            =   750
            TabIndex        =   44
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   $"frmMain.frx":0DF0
            Height          =   495
            Left            =   255
            TabIndex        =   43
            Top             =   285
            Width           =   8220
         End
      End
      Begin VB.Frame Frame6 
         ForeColor       =   &H00FF0000&
         Height          =   2760
         Left            =   -74835
         TabIndex        =   39
         Top             =   1050
         Width           =   8895
         Begin VB.CommandButton btnCreateDBObjects 
            Caption         =   "Create Database Log Objects"
            Height          =   390
            Left            =   5415
            TabIndex        =   5
            Top             =   2130
            Width           =   2505
         End
         Begin VB.CheckBox chkLogToDB 
            Caption         =   "Log to tblErrorLog in database."
            Height          =   270
            Left            =   270
            TabIndex        =   4
            Top             =   1215
            Width           =   2715
         End
         Begin VB.CheckBox chkLogToFile 
            Caption         =   "Log to file ""Errors.txt"""
            Height          =   255
            Left            =   270
            TabIndex        =   3
            Top             =   870
            Width           =   2085
         End
         Begin VB.CheckBox chkLogToEventLog 
            Caption         =   "Log to Event Log. Note: Logging to NT Event Log is disabled when run in VB IDE"
            Height          =   315
            Left            =   270
            TabIndex        =   2
            Top             =   480
            Width           =   7095
         End
         Begin VB.Label Label10 
            Caption         =   "Click ""Create..."" button to create tblErrorLog and spErrorLogInsert in MS SQL Server database. "
            Height          =   450
            Left            =   270
            TabIndex        =   87
            Top             =   2145
            Width           =   4920
         End
         Begin VB.Label Label11 
            Caption         =   "Check the boxes to activate error logging in demos."
            Height          =   345
            Left            =   120
            TabIndex        =   41
            Top             =   195
            Width           =   7545
         End
         Begin VB.Label Label12 
            Caption         =   $"frmMain.frx":0EBC
            Height          =   645
            Left            =   3075
            TabIndex        =   40
            Top             =   1200
            Width           =   5400
         End
      End
      Begin VB.Frame frames 
         ForeColor       =   &H00FF0000&
         Height          =   2775
         Index           =   2
         Left            =   180
         TabIndex        =   37
         Top             =   1050
         Width           =   8940
         Begin VB.CommandButton cmdTestConnection 
            Caption         =   "Test Connection"
            Height          =   330
            Left            =   7110
            TabIndex        =   1
            Top             =   1620
            Width           =   1635
         End
         Begin VB.TextBox txtConnectString 
            Height          =   330
            Left            =   120
            TabIndex        =   0
            Text            =   "Provider=SQLOLEDB.1;Password=;User ID=sa;Initial Catalog=Northwind"
            Top             =   1605
            Width           =   6960
         End
         Begin VB.Label Label13 
            Caption         =   $"frmMain.frx":0F50
            Height          =   555
            Left            =   180
            TabIndex        =   90
            Top             =   720
            Width           =   5940
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            Caption         =   $"frmMain.frx":100F
            Height          =   510
            Left            =   165
            TabIndex        =   38
            Top             =   210
            Width           =   7185
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.Label Label6 
      Caption         =   $"frmMain.frx":10D5
      Height          =   465
      Left            =   225
      TabIndex        =   36
      Top             =   75
      Width           =   9240
   End
End
Attribute VB_Name = "frmMain"
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
'Demo Application, main form
'=========================================================================================
Option Explicit
'This is the example of defining custom error numbers
Public Enum ENUM_ERRORS_DEMO
    ERRD_FIRST = ERRMAP_APP_FIRST
    ERRD_API
    ERRD_XMLLIB_NOTAVAILABLE
End Enum

Private Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private mConn As Object


'1. DB Connection ======================================================
Private Sub cmdTestConnection_Click()
    On Error GoTo errHandler
    Dim ErrMsg As String
    Screen.MousePointer = vbHourglass
    If TestConnection(txtConnectString.Text, ErrMsg) Then
        MsgBox "Connection tested successfully."
        Else
        MsgBox "Error: " & ErrMsg, vbCritical Or vbOKOnly, "Error"
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrorIn "frmMain.cmdTestConnection_Click", , EA_NORERAISE
    HandleError
End Sub

Private Function TestConnection(ByVal AConnString As String, ByRef ErrMsg As String) As Boolean
    On Error GoTo errHandler
    Set mConn = CreateObject("ADODB.Connection")
    mConn.Open AConnString
    TestConnection = True
    Exit Function
errHandler:
    ErrMsg = Err.Description
End Function


'2. Error Logging options =============================================
Private Sub btnCreateDBObjects_Click()
Const SQL_CREATETABLE = _
    "CREATE TABLE [dbo].[tblErrorLog] (" & vbNewLine & _
    "    [ErrorID] [int] IDENTITY (1, 1) NOT NULL ,  [DateCreated] [datetime] NOT NULL ," & vbNewLine & _
    "    [AppID] [int] NULL ,  [ErrorReport] [Text]" & vbNewLine & _
    ") ON [PRIMARY] "
Const SQL_CREATEPROC = _
    "CREATE PROCEDURE dbo.spErrorLogInsert " & vbNewLine & _
    "    @AppID INT, @ErrorReport Text AS " & vbNewLine & _
    "INSERT INTO tblErrorLog (DateCreated, AppID, ErrorReport) " & vbNewLine & _
    "VALUES  (GetDate(),@AppID, @ErrorReport) "
Const SQL_GRANTEXEC = "GRANT EXECUTE ON spErrorLogInsert to public " & vbNewLine
    On Error GoTo errHandler
    If MsgBox("Application will create log objects in Northwind database. OK to go on? ", vbYesNo, "Confirm") = vbYes Then
        Set mConn = CreateObject("ADODB.Connection")
        With mConn
            .Open txtConnectString.Text
            .Execute SQL_CREATETABLE
            .Execute SQL_CREATEPROC
            .Execute SQL_GRANTEXEC
        End With
        Set mConn = Nothing
        MsgBox "Database objects were created successfully."
    End If
    Exit Sub
errHandler:
    ErrorIn "btnCreateDBObjects_Click", , EA_NORERAISE
    HandleError
End Sub

'3. HuntERR Basic Features ======================================================
Private Sub cmdCalc_Click()
      Debug.Assert FalseIfWantStepIn
10    On Error GoTo errHandler
20    MsgBox "Result: " & CalcFormula(txtX.Text), , "Success"
30    Exit Sub
errHandler:
    ErrorIn "frmMain.cmdCalc_Click", , EA_NORERAISE, , "txtX.Text", txtX.Text
    HandleError
End Sub

Private Function CalcFormula(ByVal X As Double) As Double
11    On Error GoTo errHandler
12    Check X <= 100, EXC_VALIDATION, "Invalid value of X: %1. X must be less than 100", X 'Just example of using Check
13    CalcFormula = Divide(1, X - 1)
14    Exit Function
errHandler:
    ErrorIn "frmMain.CalcFormula(X)", X
End Function

Private Function Divide(ByVal A As Double, ByVal B As Double) As Double
    On Error GoTo errHandler
50    On Error GoTo errHandler
55      Divide = A / B
60    Exit Function
    Exit Function
errHandler:
    ErrorIn "frmMain.Divide(A,B)", Array(A, B)
End Function

'4. Loss of Err information ==============================================================
Private Sub btnErrClearedExecute_Click()
    Debug.Assert FalseIfWantStepIn
    On Error GoTo errHandler
    MsgBox "(Just generating Division by Zero)" & CStr(1 / 0)
    Exit Sub
errHandler:
    'Normally App should call ErrPreserve here, but we don't,
    'just to show how ErrorIn detects this situation
    DoSomeCleanUp
    'Err object is cleared. ErrorIn will recognize this and put a message in Error report
    ErrorIn "frmMain.btnErrClearedExecute_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub DoSomeCleanUp()
    On Error Resume Next 'This clears Err object
    'some stuff
End Sub


'5. Long Strings ====================================================================================
Private Sub cmdLDExecute_Click()
    Debug.Assert FalseIfWantStepIn
    On Error GoTo errHandler
    TestLongData "This is short parameter", "This parameter is a little longer than 40 characters.", _
        "And this has CRLF" & vbNewLine & "inside"
    Exit Sub
errHandler:
    ErrorIn "frmMain.cmdLDExecute_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub TestLongData(ByVal Prm1 As String, ByVal Prm2 As String, ByVal Prm3 As String)
    On Error GoTo errHandler
    MsgBox "(Just generating error)" & CStr(1 / 0)
    Exit Sub
errHandler:
    ErrorIn "frmMain.TestLongData(Prm1,Prm2,Prm3)", Array(Prm1, Prm2, Prm3)
End Sub

'6. API errors =============================================
'As we pass 0 as window handle GetWindowRect will fail and return zero, which
'is indication of error. Application should call GetLastError API function
'(in VB Err.LastDllError) to get error number, and then FormatMessage API
'function to get error description. ErrorIn does this for you automatically.
Private Sub btnTestAPI_Click()
    Dim rct As RECT
    Debug.Assert FalseIfWantStepIn
    On Error GoTo errHandler
    'GetWindowRect is doomed to fail.
    Check GetWindowRect(0, rct) <> 0, ERRD_API, "GetWindowRect returned 0."
    Exit Sub
errHandler:
    ErrorIn "btnTestAPI_Click", , EA_NORERAISE
    HandleError
End Sub

'7. System exceptions handling =============================================
'Demonstrates handling serious system exceptions like Acess Violation
Private Sub btnSysHandlerExecute_Click()
    'Dim sBuffer as String * 255  - This is correct declaration for API function call.
    Dim L As Long, sBuffer As String ' But we do it wrong
    Debug.Assert FalseIfWantStepIn
    On Error GoTo errHandler
    If chkSetHandler Then ErrSysHandlerSet 'Set handler only if box is checked
    L = 255 'We promise buffer of 255 chars, but provide empty string. We are heading for serious trouble...
    GetComputerNameAPI sBuffer, L 'Here is trouble
    ErrSysHandlerRelease 'This is what you should normally do
    Exit Sub
errHandler:
    ErrSysHandlerRelease 'New in 3.1: ErrorIn does not call this function automatically
    ErrorIn "btnSysHandlerExecute_Click", , EA_NORERAISE
    HandleError
End Sub

'8. HuntERR extensions, MSXML errors =============================================
Private Sub btnParseXML_Click()
    Debug.Assert FalseIfWantStepIn
    On Error GoTo errHandler
    ParseXML txtXML.Text
    MsgBox "XML string parsed successfully."
    Exit Sub
errHandler:
    ErrorIn "btnParseXML_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub ParseXML(ByVal xmlSample As String)
    Dim xmlDoc As Object
    On Error GoTo errHandler
    Check Trim$(xmlSample) <> "", EXC_VALIDATION, "XML string may not be empty."
    Set xmlDoc = CreateXmlDoc
    Check Not (xmlDoc Is Nothing), ERRD_XMLLIB_NOTAVAILABLE, "MSXML library is not available. Operation cancelled."
    Check xmlDoc.loadXML(xmlSample), ERR_GENERAL, "LoadXML method failed." 'Here we use general err number
    Exit Sub
errHandler:
    ErrGetFromServer New ujEEDomDoc, xmlDoc, , " This is comment "
    ErrorIn "ParseXML(xmlSample)", "..."
End Sub

Private Function CreateXmlDoc() As Object
    On Error Resume Next
    Set CreateXmlDoc = CreateObject("MSXML.DOMDocument")
End Function

'9. XML Formatting ======================================================
Private Sub btnXMLFmtExecute_Click()
    Dim strSample As String
    Debug.Assert FalseIfWantStepIn
    On Error GoTo errHandler
    strSample = LoadText(App.Path & "\Sample.xml")
    TestXMLFormatting strSample, strSample
    Exit Sub
errHandler:
    ErrorIn "frmMain.btnXMLFmtExecute_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub TestXMLFormatting(ByVal strParam1 As String, ByVal xmlParam2 As String)
    On Error GoTo errHandler
    MsgBox "(Just generating error)" & CStr(1 / 0)
    Exit Sub
errHandler:
    ErrorIn "frmMain.TestXMLFormatting(strParam1,xmlParam2)", Array(strParam1, xmlParam2)
End Sub

'10. Message Source ===============================================================
Private Sub cmdMsgSrcExecute_Click()
    On Error GoTo errHandler
    frmShowError.Hide
    ' "||" is replaced by vbNewLine automatically.
    ' We use HelpFile param to send name of control to set focus to
    Debug.Assert FalseIfWantStepIn
    Set ErrMessageSource = New CDemoMsgSrc 'Create and set custom message source
    Select Case cmbMsgSrcTestType.ListIndex
        Case 0
            'Hard-code the message, without use of Message Source.
            Check False, EXC_VALIDATION, _
                "This is hard-coded message with params [%1] and " & _
                "[%2] ||and line break.  Focus will be set to Param #1", _
                Array(txtMsgPrm1.Text, txtMsgPrm2.Text), "txtMsgPrm1"
        Case 1
            'Providing MsgID=1 instead of error description
            Check False, EXC_VALIDATION, "#1", Array(txtMsgPrm1.Text, txtMsgPrm2.Text)
        Case 2
            'Providing MsgID=1 and default message, in case if message not found.
            Check False, EXC_VALIDATION, _
                "#1|| This is default message for ID=1 with params [%1] and [%2]", _
                Array(txtMsgPrm1.Text, txtMsgPrm2.Text)
        Case 3
            'Providing MsgID=2 (it does not exist) and default message, which will be actually shown
            Check False, EXC_VALIDATION, _
                "#2|| This is default message for ID=2 with params [%1] and " & _
                "[%2]||... and line break. Focus will be set to Param #2", _
                Array(txtMsgPrm1.Text, txtMsgPrm2.Text), "txtMsgPrm2"
    End Select
    Set ErrMessageSource = Nothing
    Exit Sub
errHandler:
    ErrorIn "frmMain.cmdMsgSrcExecute_Click", , EA_NORERAISE
    HandleError
    Set ErrMessageSource = Nothing
End Sub

'11. Accumulation of Exception messages =============================================================
Private Sub cmdAccumValidate_Click()
    Debug.Assert FalseIfWantStepIn
    On Error GoTo errHandler
    ValidateAll
    MsgBox "Validation passed OK"
    Exit Sub
errHandler:
    ErrorIn "frmMain.cmdAccumValidate_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub ValidateAll()
    On Error GoTo errHandler
    ErrAccumBuffer = "" 'Clear the buffer
    'Start validating one by one, with 0 as error number (meaning accumulation of errors)
    Check Trim$(txtAccumName.Text) <> "", ERR_ACCUMULATE, "Name may not be empty"
    Check txtAccumSSN.Text Like "???-??-????", ERR_ACCUMULATE, _
        "SSN entered [%1] is invalid. Should be in format '123-45-6789'", txtAccumSSN.Text
    Check txtAccumPhone.Text Like "???-???-????", ERR_ACCUMULATE, _
        "Phone you entered [%1] is invalid. Should be in format '123-456-7890'", txtAccumPhone.Text
    'Now check if there is something in the buffer, and raise Exc if yes
    Check ErrAccumBuffer = "", EXC_MULTIPLE, "Please correct the following errors: ||" & ErrAccumBuffer
    Exit Sub
errHandler:
    ErrorIn "frmMain.ValidateAll"
End Sub

'12. Releasing objects ==========================================================================
Private Sub btnExecuteRlsObjs_Click()
    Debug.Assert FalseIfWantStepIn
    On Error GoTo errHandler
    TestRlsObjs
    Exit Sub
errHandler:
    ErrorIn "frmMain.btnExecuteRlsObjs_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub TestRlsObjs()
    Dim Obj1 As CTestClass, Obj2 As CTestClass
    On Error GoTo errHandler
    Set Obj1 = New CTestClass
    Obj1.ObjectName = "FirstObject"
    Set Obj2 = New CTestClass
    Obj2.ObjectName = "SecondObject"
    MsgBox "(Just generating error)" & CStr(1 / 0)
    Exit Sub
errHandler:
    ErrRlsObjs Obj1, Obj2 'Notice that messages are shown AFTER return from ErrRlsObjs call!
    ErrorIn "frmMain.TestRlsObjs"
End Sub

'13. ADO Errors ====================================================================================
'We declare ADO objects as Object on purpose, to show that HuntERR does not require
'references to ADO. This Demo project doesn't refer to ADO library although it uses ADO
Private Sub cmdExecuteSQL_Click()
    Debug.Assert FalseIfWantStepIn
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    ExecSQLExt txtConnectString.Text, txtSQL.Text, chkADOInTrans, _
        chkADOAbortTrans, chkADOAutoClose, chkADOAutoRelease
    Screen.MousePointer = vbDefault 'In case of error HandleError will restore the cursor
    MsgBox "Statement executed successfully."
    Exit Sub
errHandler:
    ErrorIn "frmMain.cmdExecuteSQL_Click", , EA_NORERAISE
    HandleError
    lblADOStatus.Caption = ReportConnStatus
End Sub

Private Sub ExecSQLExt(ByVal ConnectString As String, ByVal SQL As String, _
                        ByVal WithTrans As Boolean, ByVal TransAbort As Boolean, _
                        ByVal ConnClose As Boolean, ConnRelease As Boolean)
    On Error GoTo errHandler
    Dim EAction As Long
    'Setup appropriate ErrorAction parameter
    EAction = EA_DEFAULT
    If TransAbort Then EAction = EAction + EA_ROLLBACK
    If ConnClose Then EAction = EAction + EA_CONN_CLOSE
    'Create module-level Connection object, connect, and execute statement
    Set mConn = CreateObject("ADODB.Connection")
    mConn.Open ConnectString
    If WithTrans Then mConn.BeginTrans
    mConn.Execute SQL
    If WithTrans Then mConn.CommitTrans
    Exit Sub
errHandler:
    If ConnRelease Then ErrRlsObjs mConn
    ErrorIn "frmMain.ExecSQLExt(ConnectString,SQL,WithTrans,TransAbort,ConnClose,ConnRelease)", _
         Array(ConnectString, SQL, WithTrans, TransAbort, ConnClose, ConnRelease), EAction, mConn
End Sub


'14. Stop on Error ======================================================
Private Sub btnStopOnErrExecute_Click()
    Debug.Assert FalseIfWantStepIn
    On Error GoTo errHandler
    TestStopOnError
    MsgBox "No error occured or error was bypassed"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.btnStopOnErrExecute_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub TestStopOnError()
    Dim Z As Long
    On Error GoTo errHandler
    Z = TestDivide(1, 0)
    Exit Sub
errHandler:
    'Here you have a chance to correct 0 value to fix the "bug"
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.TestStopOnError"
End Sub

Private Function TestDivide(ByVal X As Long, Y As Long) As Long
    On Error GoTo errHandler
    TestDivide = X / Y
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.TestDivide(X,Y)", Array(X, Y)
End Function


'=============================== Utilities ==============================================
Private Sub Form_Load()
    sstMain.Tab = 0
    cmbMsgSrcTestType.ListIndex = 0
    chkStopInProc.Enabled = ErrInIDE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmShowError
End Sub

Private Sub chkSetHandler_Click()
    If chkSetHandler.Value = vbUnchecked Then
        MsgBox "Warning! We are heading for big trouble! If you press 'Execute' button now, " & vbNewLine & _
               "with this box unchecked, Demo application will crash together with VB IDE - " & vbNewLine & _
               "just to show you how bad it can be!", , "WARNING!"
    End If
End Sub

Private Function ReportConnStatus() As String
    On Error GoTo errHandler
    If mConn Is Nothing Then
        ReportConnStatus = "Connection object is Nothing."
        Else
        If mConn.State = 0 Then
            ReportConnStatus = "Connection object is not Nothing; it is Closed."
            Else
            ReportConnStatus = "Connection object is not Nothing; it is Opened."
        End If
    End If
    Exit Function
errHandler:
    'Nothing to do; this Sub is called from error handler
End Function

Private Function FalseIfWantStepIn() As Boolean
    FalseIfWantStepIn = Not (chkStopInProc.Value = vbChecked)
End Function

Public Function LoadText(ByVal FileName As String) As String
    On Error GoTo errHandler
    Dim S As Object
    Set S = CreateObject("ADODB.Stream")
    S.Open
    S.LoadFromFile FileName
    LoadText = S.ReadText
    S.Close
    Exit Function
errHandler:
    ErrorIn "frmMain.LoadText(FileName)", FileName
End Function

Private Sub cmdNextPage_Click()
    On Error GoTo errHandler
    With sstMain
        .Tab = IIf(.Tab = .Tabs - 1, 0, .Tab + 1)
    End With
    Exit Sub
errHandler:
    ErrorIn "frmMain.cmdNextPage_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrevPage_Click()
    On Error GoTo errHandler
    With sstMain
        .Tab = IIf(.Tab = 0, .Tabs - 1, .Tab - 1)
    End With
    Exit Sub
errHandler:
    ErrorIn "frmMain.cmdPrevPage_Click", , EA_NORERAISE
    HandleError
End Sub


'==============================================================
Public Sub HandleError()
    On Error Resume Next 'no more errors or exceptions!!!!
    Screen.MousePointer = vbDefault 'If mouse pointer was set to hour glass, then put it back to default!
    If InException Then
        Select Case ErrNumber
            Case EXC_GENERAL:    MsgBox ErrDescription, vbOKOnly, "Exception"
            Case EXC_CANCELLED:  'nothing to do - it is silent exception.
            Case EXC_MULTIPLE:   'In this case we may have a special form to show multiple input errors
                                 MsgBox ErrDescription, vbOKOnly, "Multiple Input Errors"
            Case EXC_VALIDATION: MsgBox ErrDescription, vbOKOnly, "User input error"
        End Select
    Else
        If chkLogToEventLog.Value = vbChecked Then ErrSaveToEventLog
        If chkLogToDB.Value = vbChecked Then ErrSaveToDB txtConnectString.Text
        If chkLogToFile.Value = vbChecked Then ErrSaveToFile
        frmShowError.ErrorReport = ErrReport
    End If
    'We use Err.HelpFile in some demos to send name of control to set focus to
    'We must check if it is not system-provided help file name, so it doesn't include "\"
    If InStr(1, ErrHelpFile, "\") = 0 Then SetFocusTo ErrHelpFile
End Sub

Public Sub SetFocusTo(ByVal ControlName As String)
    On Error Resume Next
    Me.Controls(ControlName).SetFocus
End Sub



