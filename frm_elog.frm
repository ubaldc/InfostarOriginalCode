VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_elog 
   Caption         =   "System Log"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11265
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_elog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin FPUSpreadADO.fpSpread ss_elog 
      Height          =   7935
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   10815
      _Version        =   458752
      _ExtentX        =   19076
      _ExtentY        =   13996
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   -2147483633
      MaxCols         =   6
      MaxRows         =   500000
      OperationMode   =   1
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   0
      SpreadDesigner  =   "frm_elog.frx":0442
      UserResize      =   1
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9840
      Top             =   8400
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9315
      FormDesignWidth =   11265
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_clear 
      Height          =   650
      Left            =   2520
      TabIndex        =   1
      Top             =   8280
      Width           =   2000
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_elog.frx":06B6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_elog.frx":06E8
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_elog.frx":0708
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_exit 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   6720
      TabIndex        =   0
      Top             =   8280
      Width           =   2000
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_elog.frx":0724
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_elog.frx":074C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_elog.frx":076C
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   10560
      Top             =   8400
      _ExtentX        =   794
      _ExtentY        =   794
      PersistentTip   =   0   'False
      ActivateTips    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TipBgColor      =   -2147483624
      TipFgColor      =   -2147483625
      Style           =   -1
      Transparency    =   0
   End
   Begin HexUniControls.ctlUniFormCaption uniTitle 
      Left            =   9240
      Top             =   8280
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_elog.frx":0788
   End
End
Attribute VB_Name = "frm_elog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_exit_Click()
  
  unity_main.errorstring = "System Log screen Exit button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Unload frm_elog
End Sub

Private Sub cmd_clear_Click()
  Dim ftokill As String
  Dim optVal As Integer
  Dim errMsg As String
  Dim uniFile As New clsUniFile
  Dim uniMsg As String

  unity_main.errorstring = "System Log screen Clear Log button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  On Error GoTo NO_DELETE
  ftokill = (LOGFILE_DIR & SYSTEM_LOG_FILE)
  uniMsg = MLSupport.GGS_Params("fileMsg1", "Are you sure you want to delete %1?", ftokill)
  optVal = CWrap.ShowMessageBoxW(uniMsg, vbYesNo)
  
  If (optVal = vbYes) Then
    uniFile.st_RmFile ftokill
    Unload frm_elog
  End If
  
  unity_main.checklogfile

  unity_main.errorstring = "User cleared system log file"
  unity_main.write_error (LOG_DBG_LEVEL1)
  Exit Sub
  
NO_DELETE:
  errMsg = (ftokill & " file delete error. " & Error$)
  unity_main.errorstring = errMsg
  unity_main.write_error (LOG_DBG_LEVEL1)
  uniMsg = MLSupport.GGS_Params("fileErrMsg8", "%1 file delete error. %2", ftokill, Error$)
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub








