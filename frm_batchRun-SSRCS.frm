VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "resize32.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "Fpspru70.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_batchRun 
   Caption         =   "Batch Run Control"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13350
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   13350
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniListBoxExXP lst_towersInUse 
      Height          =   495
      Left            =   11160
      TabIndex        =   4
      Top             =   7680
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      IconDim         =   16
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TrapTab         =   0   'False
      Tip             =   "frm_batchRun-SSRCS.frx":0000
      MultiSelect     =   0
      Sorted          =   -1  'True
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_batchRun-SSRCS.frx":0020
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_abortBatch 
      Height          =   705
      Left            =   8160
      TabIndex        =   3
      Top             =   5040
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   1244
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
      Caption         =   "frm_batchRun-SSRCS.frx":003C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_batchRun-SSRCS.frx":0072
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_batchRun-SSRCS.frx":0092
   End
   Begin VB.Timer tmr_pollStatus 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   1680
      Top             =   7440
   End
   Begin VB.Timer tmr_batch 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1080
      Top             =   7440
   End
   Begin FPUSpreadADO.fpSpread ss_batchInfo 
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   13120
      _Version        =   458752
      _ExtentX        =   23142
      _ExtentY        =   6800
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      OperationMode   =   1
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   0
      SpreadDesigner  =   "frm_batchRun-SSRCS.frx":00AE
      UserResize      =   1
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   480
      Top             =   7440
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7920
      FormDesignWidth =   13350
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_exit 
      Cancel          =   -1  'True
      Height          =   705
      Left            =   8160
      TabIndex        =   0
      Top             =   5040
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   1244
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
      Caption         =   "frm_batchRun-SSRCS.frx":0390
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_batchRun-SSRCS.frx":03B8
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_batchRun-SSRCS.frx":03D8
   End
   Begin HexUniControls.ctlUniLabel lbl_scanProgress 
      Height          =   255
      Left            =   4380
      Top             =   6960
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_batchRun-SSRCS.frx":03F4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_batchRun-SSRCS.frx":042E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_batchRun-SSRCS.frx":044E
   End
   Begin HexUniControls.ctlProgressXP scanProgress 
      Height          =   375
      Left            =   4320
      Top             =   7320
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BackColor       =   16777215
      ForeColor       =   49152
      Border          =   -1  'True
      BorderSpace     =   -1  'True
      Spaces          =   -1  'True
      Tip             =   "frm_batchRun-SSRCS.frx":046A
      Style           =   -1
      BackStyle       =   -1
      RoundedBorders  =   -1  'True
      MousePointer    =   0
      MouseIcon       =   "frm_batchRun-SSRCS.frx":048A
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_batchReport 
      Height          =   705
      Left            =   2970
      TabIndex        =   1
      Top             =   5040
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   1244
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
      Caption         =   "frm_batchRun-SSRCS.frx":04A6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_batchRun-SSRCS.frx":04DE
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_batchRun-SSRCS.frx":04FE
   End
   Begin HexUniControls.ctlUniLabel lbl_batchProgress 
      Height          =   255
      Left            =   5205
      Top             =   5880
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_batchRun-SSRCS.frx":051A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_batchRun-SSRCS.frx":0556
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_batchRun-SSRCS.frx":0576
   End
   Begin HexUniControls.ctlUniLabel lbl_batchProgress2 
      Height          =   495
      Left            =   1485
      Top             =   6240
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_batchRun-SSRCS.frx":0592
      BackColor       =   -2147483643
      ForeColor       =   12583104
      Alignment       =   2
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_batchRun-SSRCS.frx":05B2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_batchRun-SSRCS.frx":05D2
   End
   Begin HexUniControls.ctlUniLabel lbl_batchName 
      Height          =   405
      Left            =   3368
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_batchRun-SSRCS.frx":05EE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_batchRun-SSRCS.frx":0622
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_batchRun-SSRCS.frx":0642
   End
   Begin HexUniControls.ctlUniLabel lbl_batchNameFile 
      Height          =   405
      Left            =   6368
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_batchRun-SSRCS.frx":065E
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frm_batchRun-SSRCS.frx":067E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_batchRun-SSRCS.frx":069E
   End
   Begin HexUniControls.ctlUniLabel lbl_autoSamplrVer 
      Height          =   405
      Left            =   120
      Top             =   4560
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_batchRun-SSRCS.frx":06BA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_batchRun-SSRCS.frx":06DA
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_batchRun-SSRCS.frx":06FA
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_emergencyStop 
      Height          =   705
      Left            =   2970
      TabIndex        =   5
      Top             =   5040
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   1244
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
      Caption         =   "frm_batchRun-SSRCS.frx":0716
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_batchRun-SSRCS.frx":0752
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_batchRun-SSRCS.frx":0772
   End
   Begin HexUniControls.ctlUniFormCaption uniTitle 
      Left            =   0
      Top             =   7440
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_batchRun-SSRCS.frx":078E
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   0
      Top             =   6840
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
End
Attribute VB_Name = "frm_batchRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_autoSamplrCmdState As AUTO_SAMPLR_CMD_STATES
Private m_autoSmplrDestCupPos As Integer
Private m_autoSmplrError As Integer
Private m_autoSamplrLastCommErr As String
Private m_autoSamplrLastCommErrCode As Integer
Private m_autoSamplrLastCommStat As AutoSamplerCommCtrl.eASCommStats
Private m_autoSamplrLastOperErr As String
Private m_autoSamplrOperState As AUTO_SAMPLR_OPER_STATES
Private m_batchAbortFlg As Integer
Private m_batchTmrState As BATCH_TMR_STATES
Private m_currentBatchSmplNum As Integer
Private m_currentEntryZoneSmplInfo As SampleCupInfo
Private m_currentExitZoneSmplInfo As SampleCupInfo
Private m_currentSmplWinSmplInfo As SampleCupInfo
Private m_dsrState As Boolean
Private m_dumpPosIndx As Integer
Private m_fetchCupNum As Integer
Private m_fetchPosIndx As Integer
Private m_numSamplesScan As Integer
Private m_pollStatusState As POLL_STATUS_STATES
Private m_prodIni As String
Private m_scanAttempts As Integer
Private m_scanFlg As Integer
Private m_scanState As SCAN_TMR_STATES
Private m_smplData As SampleScanInfo

Private Const NUM_ROWS_DISP = 9

Public Sub init_batch()
  
  ' Setup column names
  build_column_names
  
  ' Populate spreadsheet with configured info
  display_batch_info
  
  ' Reset scan progress bar
  scanProgress.percent = 0
    
  ' Show/hide buttons for now
  cmd_abortBatch.Visible = True
  cmd_batchReport.Visible = False
  cmd_emergencyStop.Visible = True
  cmd_exit.Visible = False
  
  unity_main.img_batchRpt.Visible = False
  unity_main.m_batchRptFile = ""
  unity_main.m_batchRunFlg = True
  unity_main.m_scanTmrState = STS_COMPLETED

  lst_towersInUse.Clear
  
  m_autoSmplrDestCupPos = AUTO_SMPLR_UNKNOWN_POS
  m_autoSamplrLastCommErr = ""
  m_autoSamplrLastCommErrCode = 0
  m_autoSamplrLastOperErr = ""
  m_autoSamplrOperState = ASOS_IDLE
  m_batchAbortFlg = 0
  m_currentBatchSmplNum = 0
  m_numSamplesScan = 0
  m_currentEntryZoneSmplInfo.fetchCupNum = 0
  m_currentEntryZoneSmplInfo.sampleNum = 0
  m_currentExitZoneSmplInfo.fetchCupNum = 0
  m_currentExitZoneSmplInfo.sampleNum = 0
  m_currentSmplWinSmplInfo.fetchCupNum = 0
  m_currentSmplWinSmplInfo.sampleNum = 0
  m_dumpPosIndx = AUTO_SMPLR_UNKNOWN_POS
  m_fetchCupNum = 0
  m_fetchPosIndx = AUTO_SMPLR_UNKNOWN_POS
  m_numSamplesScan = 0
  m_prodIni = ""
  m_scanFlg = 0
  m_smplData.scanTime = ""
  
  NumAutoSmplrErrors = 1
  ReDim AutoSmplrErrors(NumAutoSmplrErrors - 1)
  
  AutoSmplrActDumpTower = AUTO_SMPLR_UNKNOWN_POS
  AutoSmplrActFetchTower = AUTO_SMPLR_UNKNOWN_POS
  
  m_batchTmrState = BTS_BATCH_INIT
  tmr_batch.enabled = True
End Sub

Public Sub batch_scan_aborted()

  BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).smplStatus = BSS_BATCH_ABORT

  ' Check if aborted reference scan
  If (m_scanFlg = 1) Then
    unity_main.tmr_ref.enabled = False
    report_err_status m_currentSmplWinSmplInfo.sampleNum, MLSupport.GSS("frm_batchRun", "errMsg3", "Reference Scan Aborted")
  Else
    unity_main.tmr_sample.enabled = False
    report_err_status m_currentSmplWinSmplInfo.sampleNum, MLSupport.GSS("frm_batchRun", "errMsg4", "Sample Scan Aborted")
  End If
  
  m_scanFlg = 0
End Sub

Public Sub process_auto_sampler_response(ByVal cmdId As AutoSamplerCommCtrl.eASCmdIds, cmdIdTxt As String, ByVal commStat As AutoSamplerCommCtrl.eASCommStats, rspData As String)
  Dim varStr As Variant
  Dim ii As Integer
  Dim errMsg As String
  Dim rspLen As Integer

BAD_RESP:
  If (commStat <> eASCommStats.RSP_GOOD) Then
    errMsg = MLSupport.GSS("frm_batchRun", "errMsg10", "Auto-Sampler Communication Failure") & ". "
  
    Select Case (commStat)
      Case eASCommStats.COMMS_NOT_INIT
        errMsg = errMsg & MLSupport.GGS_Params("frm_batchRun.comErrMsg1", "Cannot open or initialize serial port %1", CStr(unity_main.m_autoSmplrPort))
    
      Case eASCommStats.RSP_TIMEOUT
        errMsg = errMsg & MLSupport.GGS_Params("frm_batchRun.comErrMsg2", "Did not receive response for %1 command", cmdIdTxt)
    
      Case eASCommStats.RSP_ERR
        ' Get error code from response
        varStr = Split(rspData, " ")
        m_autoSamplrLastCommErrCode = varStr(2)
        errMsg = errMsg & MLSupport.GGS_Params("frm_batchRun.comErrMsg3", "Received error code %1 for %2 command", Hex(m_autoSamplrLastCommErrCode), cmdIdTxt)
      
      Case eASCommStats.RSP_INV_SIZE
        errMsg = errMsg & MLSupport.GGS_Params("frm_batchRun.comErrMsg4", "Received response with invalid number of characters for %1 command", cmdIdTxt)
    
      Case eASCommStats.RSP_INV
        errMsg = errMsg & MLSupport.GGS_Params("frm_batchRun.comErrMsg5", "Received invalid response for %1 command", cmdIdTxt)
    
      Case eASCommStats.ERR_CRC
        errMsg = errMsg & MLSupport.GGS_Params("frm_batchRun.comErrMsg6", "Received response with CRC error for %1 command", cmdIdTxt)
    
      Case eASCommStats.ERR_BREAK
        errMsg = errMsg & MLSupport.GGS_Params("frm_batchRun.comErrMsg7", "Received response with Break character for %1 command", cmdIdTxt)
    
      Case eASCommStats.ERR_FRAME
        errMsg = errMsg & MLSupport.GGS_Params("frm_batchRun.comErrMsg8", "Received response with framing error for %1 command", cmdIdTxt)
        
      Case eASCommStats.ERR_OVERRUN
        errMsg = errMsg & MLSupport.GGS_Params("frm_batchRun.comErrMsg9", "Received response with port overrun error for %1 command", cmdIdTxt)
        
      Case eASCommStats.ERR_OVERFLOW
        errMsg = errMsg & MLSupport.GGS_Params("frm_batchRun.comErrMsg10", "Received response with buffer overflow error for %1 command", cmdIdTxt)
    
      Case eASCommStats.ERR_PARITY
        errMsg = errMsg & MLSupport.GGS_Params("frm_batchRun.comErrMsg11", "Received response with parity error for %1 command", cmdIdTxt)
    End Select
    
    ' Show error message to user if not running batch and not scanning
    If (unity_main.m_batchRunFlg = False) And (unity_main.m_scanTmrState = STS_COMPLETED) Then
      unity_main.errorstring = errMsg
      unity_main.write_error
      CWrap.ShowMessageBoxW errMsg, vbCritical
    End If
    
    m_autoSamplrLastCommErr = errMsg
    m_autoSamplrLastCommStat = commStat
    m_autoSamplrCmdState = ASCS_COMPLETED_ERR
  Else
    ' Process response based on command ID
    varStr = Split(rspData, " ")
  
    Select Case (cmdId)
      Case eASCmdIds.ACK_ST
      
      Case eASCmdIds.ABT_OPS
      
      Case eASCmdIds.CLR_ERP
        For ii = 0 To NumAutoSmplrErrors - 1
          AutoSmplrErrors(ii) = 0
        Next ii
      
      Case eASCmdIds.NONE
      
      Case eASCmdIds.PRK_ARM
        ' Check if arm is moving to park position
        If (varStr(1) = "?") Then
          AutoSmplrCupPosition = AUTO_SMPLR_UNKNOWN_POS
        Else
          ' Check if invalid cup position value
          If (IsNumeric(varStr(1)) = False) Then
            commStat = RSP_INV
            GoTo BAD_RESP
          Else
            AutoSmplrCupPosition = varStr(1)
          End If
        End If
      
      Case eASCmdIds.RD_CM
      
      Case eASCmdIds.RD_CTL
        ii = varStr(1)
        
        If ((ii >= AUTO_SMPLR_CUP_CNT_CTRL_WORD_OFFSET) And (ii <= (AUTO_SMPLR_CUP_CNT_CTRL_WORD_OFFSET + AUTO_SMPLR_OUT_WAY_POS))) Then
          ii = ii - AUTO_SMPLR_CUP_CNT_CTRL_WORD_OFFSET
          AutoSmplrTowerNumCups(ii) = "&H" & varStr(2)
        End If
      
      Case eASCmdIds.RD_CUP
        ' Check if arm is moving to cup position
        If (varStr(1) = "?") Then
          AutoSmplrCupPosition = AUTO_SMPLR_UNKNOWN_POS
          
          ' Check if invalid switch or jaw status value
          If (IsNumeric(varStr(2)) = False) Or (IsNumeric(varStr(3)) = False) Then
            commStat = RSP_INV
            GoTo BAD_RESP
          End If
        Else
          ' Check if invalid cup position, switch or jaw status value
          If (IsNumeric(varStr(1)) = False) Or (IsNumeric(varStr(2)) = False) Or (IsNumeric(varStr(3)) = False) Then
            commStat = RSP_INV
            GoTo BAD_RESP
          Else
            AutoSmplrCupPosition = varStr(1)
          End If
        End If
        
        AutoSmplrSwitchStatus = varStr(2)
        AutoSmplrJawStatus = varStr(3)
        
      Case eASCmdIds.RD_ERP
        ' Check if invalid number of errors value
        If (IsNumeric(varStr(1)) = False) Then
          commStat = RSP_INV
          GoTo BAD_RESP
        Else
          NumAutoSmplrErrors = varStr(1)
          
          If (NumAutoSmplrErrors > 0) Then
            ReDim AutoSmplrErrors(NumAutoSmplrErrors - 1)
        
            For ii = 0 To NumAutoSmplrErrors - 1
              AutoSmplrErrors(ii) = Val("&H" & varStr(ii + 2))
            Next ii
          End If
        End If
        
      Case eASCmdIds.RD_ST
        ' Check if invalid status value
        If (IsNumeric("&H" & varStr(2)) = False) Then
          commStat = RSP_INV
          GoTo BAD_RESP
        Else
          AutoSmplrStatus = Val("&H" & varStr(2))
        End If
        
      Case eASCmdIds.RD_TUB
        ii = varStr(1)
        
        ' Check if arm is moving to tube to check cup count
        If (varStr(2) = "?") Then
          AutoSmplrTowerState(ii) = UNKNOWN_TUBE
          AutoSmplrTowerNumCups(ii) = AUTO_SMPLR_UNKNOWN_CUPS
        Else
          ' Check if invalid tube state or number of cup value
          If (IsNumeric(varStr(2)) = False) Or (IsNumeric(varStr(3)) = False) Then
            commStat = RSP_INV
            GoTo BAD_RESP
          Else
            ' Save tube state and number of cups
            AutoSmplrTowerState(ii) = varStr(2)
            AutoSmplrTowerNumCups(ii) = varStr(3)
            
            Select Case (AutoSmplrTowerState(ii))
              Case AutoSamplerCommCtrl.eASTubeStates.ACTIVE_DUMP_TUBE
                AutoSmplrActDumpTower = ii
                
              Case AutoSamplerCommCtrl.eASTubeStates.ACTIVE_FETCH_TUBE
                AutoSmplrActFetchTower = ii
            End Select
          End If
        End If
      
      Case eASCmdIds.RD_TUBS
        ' Check if invalid number of tubes value
        If (IsNumeric(varStr(1)) = False) Then
          commStat = RSP_INV
          GoTo BAD_RESP
        Else
          NumAutoSmplrTowers = varStr(1)
          NumAutoSmplrTowers = 4
          tempNumAutoSmplrTowers = varStr(1)
          For ii = 0 To tempNumAutoSmplrTowers - 1
            ' Check if invalid tube state value
            If (IsNumeric(varStr(ii + 2)) = False) Then
              tempNumAutoSmplrTowers = 0
              commStat = RSP_INV
              GoTo BAD_RESP
            Else
              AutoSmplrTowerState(ii) = varStr(ii + 2)
              AutoSmplrTowerNumCups(ii) = AUTO_SMPLR_UNKNOWN_CUPS
            
              Select Case (AutoSmplrTowerState(ii))
                Case AutoSamplerCommCtrl.eASTubeStates.ACTIVE_DUMP_TUBE
                  AutoSmplrActDumpTower = ii
                
                Case AutoSamplerCommCtrl.eASTubeStates.ACTIVE_FETCH_TUBE
                  AutoSmplrActFetchTower = ii
              End Select
            End If
          Next ii
        End If
        
      Case eASCmdIds.RD_VER
        rspLen = Len(rspData)
        AutoSmplrVersion = Right(rspData, rspLen - 4)
        
      Case eASCmdIds.SVY_LOC
      
      Case eASCmdIds.WR_CM
      
      Case eASCmdIds.WR_CTL
        ii = varStr(1)
        
        If ((ii >= AUTO_SMPLR_CUP_CNT_CTRL_WORD_OFFSET) And (ii <= (AUTO_SMPLR_CUP_CNT_CTRL_WORD_OFFSET + AUTO_SMPLR_OUT_WAY_POS))) Then
          ii = ii - AUTO_SMPLR_CUP_CNT_CTRL_WORD_OFFSET
          AutoSmplrTowerNumCups(ii) = "&H" & varStr(2)
        End If
      
      Case eASCmdIds.WR_CUP
        ' Check if arm is moving to cup position
        If (varStr(1) = "?") Then
          AutoSmplrCupPosition = AUTO_SMPLR_UNKNOWN_POS
          
          ' Check if invalid switch or jaw status value
          If (IsNumeric(varStr(2)) = False) Or (IsNumeric(varStr(3)) = False) Then
            commStat = RSP_INV
            GoTo BAD_RESP
          End If
        Else
          ' Check if invalid cup position, switch or jaw status value
          If (IsNumeric(varStr(1)) = False) Or (IsNumeric(varStr(2)) = False) Or (IsNumeric(varStr(3)) = False) Then
            commStat = RSP_INV
            GoTo BAD_RESP
          Else
            AutoSmplrCupPosition = varStr(1)
          End If
        End If
      
        AutoSmplrSwitchStatus = varStr(2)
        AutoSmplrJawStatus = varStr(3)
      
      Case eASCmdIds.WR_CUPCTRL
        ' Check if arm is moving to cup position
        If (varStr(1) = "?") Then
          AutoSmplrCupPosition = AUTO_SMPLR_UNKNOWN_POS
          
          ' Check if invalid switch or jaw status value
          If (IsNumeric(varStr(2)) = False) Or (IsNumeric(varStr(3)) = False) Then
            commStat = RSP_INV
            GoTo BAD_RESP
          End If
        Else
          ' Check if invalid cup position, switch or jaw status value
          If (IsNumeric(varStr(1)) = False) Or (IsNumeric(varStr(2)) = False) Or (IsNumeric(varStr(3)) = False) Then
            commStat = RSP_INV
            GoTo BAD_RESP
          Else
            AutoSmplrCupPosition = varStr(1)
          End If
        End If
      
        AutoSmplrSwitchStatus = varStr(2)
        AutoSmplrJawStatus = varStr(3)
      
      Case eASCmdIds.WR_TUBSTATE
        ii = varStr(1)
        
        ' Check if arm is moving to tube to check cup count
        If (varStr(2) = "?") Then
          AutoSmplrTowerState(ii) = UNKNOWN_TUBE
          AutoSmplrTowerNumCups(ii) = AUTO_SMPLR_UNKNOWN_CUPS
        Else
          ' Check if invalid tube state or number of cup value
          If (IsNumeric(varStr(2)) = False) Or (IsNumeric(varStr(3)) = False) Then
            commStat = RSP_INV
            GoTo BAD_RESP
          Else
            AutoSmplrTowerState(ii) = varStr(2)
            AutoSmplrTowerNumCups(ii) = varStr(3)
            
            Select Case (AutoSmplrTowerState(ii))
              Case AutoSamplerCommCtrl.eASTubeStates.ACTIVE_DUMP_TUBE
                AutoSmplrActDumpTower = ii
                
              Case AutoSamplerCommCtrl.eASTubeStates.ACTIVE_FETCH_TUBE
                AutoSmplrActFetchTower = ii
            End Select
          End If
        End If
    End Select
    
    m_autoSamplrLastCommErr = ""
    m_autoSamplrLastCommErrCode = 0
    m_autoSamplrCmdState = ASCS_COMPLETED_GOOD
  End If
End Sub

Private Sub build_column_names()
  Dim colNum As Integer

  ' Setup header font info
  ss_batchInfo.Col = 0
  ss_batchInfo.Col2 = ss_batchInfo.MaxCols
  ss_batchInfo.Row = 0
  ss_batchInfo.Row2 = ss_batchInfo.MaxRows
  ss_batchInfo.BlockMode = True
  ss_batchInfo.Font.Name = "Arial Unicode MS"
  ss_batchInfo.Font.Size = 10
  ss_batchInfo.FontBold = True
  ss_batchInfo.BlockMode = False

  ss_batchInfo.Row = 0
 
  ' Setup spreadsheet column names
  colNum = BATCH_PROD_NAME_COL
  ss_batchInfo.Col = colNum
  ss_batchInfo.ColWidth(colNum) = 19
  ss_batchInfo.Text = MLSupport.GSS("Headers", "product", "Product")
  
  colNum = BATCH_SAMPLE_ID_COL
  ss_batchInfo.Col = colNum
  ss_batchInfo.ColWidth(colNum) = 12
  ss_batchInfo.Text = MLSupport.GSS("Headers", "sampleID", "Sample ID")
  
  colNum = BATCH_LOAD_TOWER_COL
  ss_batchInfo.Col = colNum
  ss_batchInfo.ColWidth(colNum) = 11
  ss_batchInfo.Text = MLSupport.GSS("Headers", "loadTower", "Load Tower")
  
  colNum = BATCH_LOAD_CUP_NUM_COL
  ss_batchInfo.Col = colNum
  ss_batchInfo.ColWidth(colNum) = 8
  ss_batchInfo.Text = MLSupport.GSS("Headers", "cupNum", "Cup #")
  
  colNum = BATCH_UNLOAD_TOWER_COL
  ss_batchInfo.Col = colNum
  ss_batchInfo.ColWidth(colNum) = 12
  ss_batchInfo.Text = MLSupport.GSS("Headers", "unloadTower", "Unload Tower")
  
  colNum = BATCH_UNLOAD_CUP_NUM_COL
  ss_batchInfo.Col = colNum
  ss_batchInfo.ColWidth(colNum) = 8
  ss_batchInfo.Text = MLSupport.GSS("Headers", "cupNum", "Cup #")
  
  colNum = BATCH_STATUS_COL
  ss_batchInfo.Col = colNum
  ss_batchInfo.ColWidth(colNum) = 32
  ss_batchInfo.Text = MLSupport.GSS("Headers", "status", "Status")
End Sub

Private Sub display_batch_info()
  Dim ii As Integer

  ' Setup spreadsheet for max # of possible samples
  ss_batchInfo.MaxRows = frm_batchRunCfg.m_numBatchSamples
  
  ' Clear spreadsheet
  ss_batchInfo.Col = 1
  ss_batchInfo.Col2 = ss_batchInfo.MaxCols
  ss_batchInfo.Row = 1
  ss_batchInfo.Row2 = ss_batchInfo.MaxRows
  ss_batchInfo.BlockMode = True
  ss_batchInfo.Action = ActionClearText
  ss_batchInfo.ForeColor = vbBlack
  ss_batchInfo.BlockMode = False
  
  ' Display configured sample data
  For ii = 1 To frm_batchRunCfg.m_numBatchSamples
    ss_batchInfo.Row = ii
    ss_batchInfo.Col = BATCH_PROD_NAME_COL
    ss_batchInfo.Text = BatchRunCfg(ii).prodName
    
    ss_batchInfo.Col = BATCH_LOAD_TOWER_COL
    ss_batchInfo.Text = BatchRunCfg(ii).loadTower
    
    ss_batchInfo.Col = BATCH_UNLOAD_TOWER_COL
    ss_batchInfo.Text = BatchRunCfg(ii).unloadTower
  Next ii

  ' Display batch name
  lbl_batchNameFile.Caption = frm_batchRunCfg.txt_batchName.Text
End Sub

Private Function reset_all_tube_states(setArmAwayFlg As Boolean) As Boolean
  Dim ii As Integer
  
  ' Init active dump and fetch towers to VOID state
  AutoSmplrActDumpTower = AUTO_SMPLR_UNKNOWN_POS
  AutoSmplrActFetchTower = AUTO_SMPLR_UNKNOWN_POS
  
  ' Clear tube cup count for sample window
  reset_all_tube_states = set_auto_sampler_tube_cup_count(AUTO_SMPLR_SMPL_WIN_POS, 0)
  
  If (reset_all_tube_states = True) Then
    ' Clear tube cup count for exit zone
    reset_all_tube_states = set_auto_sampler_tube_cup_count(AUTO_SMPLR_EXIT_ZONE_POS, 0)
  
    If (reset_all_tube_states = True) Then
      ' Clear tube cup count for entry zone
      reset_all_tube_states = set_auto_sampler_tube_cup_count(AUTO_SMPLR_ENTRY_ZONE_POS, 0)
       
      If (reset_all_tube_states = True) Then
        ' Set all towers to VOID state
        ii = 0
  
        While (reset_all_tube_states = True) And (ii < NumAutoSmplrTowers)
          reset_all_tube_states = setup_tube_state(ii, VOID_TUBE)
          ii = ii + 1
        Wend
  
        If (reset_all_tube_states = True) And (setArmAwayFlg = True) Then
          ' Set arm position to out of the way
          reset_all_tube_states = move_arm_away
        End If
      End If
    End If
  End If
  
  If (reset_all_tube_states = False) And (m_batchTmrState <> BTS_ABORT_BATCH) And (m_batchTmrState <> BTS_ESTOP_BATCH) Then
    m_batchTmrState = BTS_BATCH_ERR
  End If
End Function

Private Function survey_towers() As Boolean
  Dim ii As Integer
  Dim loadTower As Integer
  Dim unloadTower As Integer
  Dim nn As Integer

  For ii = 1 To frm_batchRunCfg.m_numBatchSamples
    ' Get fetch and dump towers for current sample
    loadTower = BatchRunCfg(ii).loadTower - 1
    unloadTower = BatchRunCfg(ii).unloadTower - 1
    
    ' Check if load tower in list of used towers
    For nn = 0 To lst_towersInUse.ListCount - 1
      If (loadTower = lst_towersInUse.List(nn)) Then
        Exit For
      End If
    Next nn
    
    ' Add load tower if not in list
    If (nn = lst_towersInUse.ListCount) Then
      lst_towersInUse.AddItem loadTower
    End If
    
    ' Check if unload tower in list of used towers
    For nn = 0 To lst_towersInUse.ListCount - 1
      If (unloadTower = lst_towersInUse.List(nn)) Then
        Exit For
      End If
    Next nn
    
    ' Add unload tower if not in list
    If (nn = lst_towersInUse.ListCount) Then
      lst_towersInUse.AddItem unloadTower
    End If
  Next ii
    
  ' Survey towers based on oreder of highest tower number
  For nn = lst_towersInUse.ListCount - 1 To 0 Step -1
    loadTower = lst_towersInUse.List(nn)
    
    If (AutoSmplrTowerNumCups(loadTower) = AUTO_SMPLR_UNKNOWN_CUPS) Then
      If (survey_tower_cups(loadTower) = False) Then
        Exit Function
      End If
    End If
  Next nn

  survey_towers = True
End Function

Private Function survey_tower_cups(tubeIndx As Integer) As Boolean
 
  m_autoSamplrOperState = ASOS_SURVEY_TUBE
  lbl_batchProgress2.Caption = MLSupport.GGS_Params("frm_batchRun.statMsg2", "Surveying Tower %1 for Number of Sample Cups", CStr(tubeIndx + 1))

  ' Set tower to inactive state
  If (setup_tube_state(tubeIndx, INACTIVE_TUBE)) Then
    ' Survey tuwer for number of cups
    If (survey_auto_sampler_arm_loc(tubeIndx) = True) Then
      ' Wait for arm to complete movement
      If (wait_busy_status = False) Then
        Exit Function
      End If
    
      ' Check if any error getting tube state and count
      If (get_auto_sampler_tube_state(tubeIndx) = False) Then
        Exit Function
      End If
    
      m_autoSamplrOperState = ASOS_IDLE
      survey_tower_cups = True
    End If
  End If
End Function

Private Function survey_entry_zone() As Boolean

  m_autoSamplrOperState = ASOS_SURVEY_TUBE
  lbl_batchProgress2.Caption = MLSupport.GSS("frm_batchRun", "statMsg11", "Surveying Entry Zone for Any Stray Sample Cup")

  ' Check if command was processed by auto-sampler
  If (survey_auto_sampler_arm_loc(AUTO_SMPLR_ENTRY_ZONE_POS) = True) Then
    ' Wait for arm to complete movement
    If (wait_busy_status = False) Then
      Exit Function
    End If
    
    ' Check if any error getting tube state and count
    If (get_auto_sampler_tube_state(AUTO_SMPLR_ENTRY_ZONE_POS) = False) Then
      Exit Function
    End If
    
    m_autoSamplrOperState = ASOS_IDLE
    survey_entry_zone = True
  End If
End Function

Private Function survey_exit_zone() As Boolean

  m_autoSamplrOperState = ASOS_SURVEY_TUBE
  lbl_batchProgress2.Caption = MLSupport.GSS("frm_batchRun", "statMsg12", "Surveying Exit Zone for Any Stray Sample Cup")

  ' Check if command was processed by auto-sampler
  If (survey_auto_sampler_arm_loc(AUTO_SMPLR_EXIT_ZONE_POS) = True) Then
    ' Wait for arm to complete movement
    If (wait_busy_status = False) Then
      Exit Function
    End If
    
    ' Check if any error getting tube state and count
    If (get_auto_sampler_tube_state(AUTO_SMPLR_EXIT_ZONE_POS) = False) Then
      Exit Function
    End If
    
    m_autoSamplrOperState = ASOS_IDLE
    survey_exit_zone = True
  End If
End Function

Private Function survey_sample_win() As Boolean

  m_autoSamplrOperState = ASOS_SURVEY_TUBE
  lbl_batchProgress2.Caption = MLSupport.GSS("frm_batchRun", "statMsg13", "Surveying Sample Window for Any Stray Sample Cup")

  ' Check if command was processed by auto-sampler
  If (survey_auto_sampler_arm_loc(AUTO_SMPLR_SMPL_WIN_POS) = True) Then
    ' Wait for arm to complete movement
    If (wait_busy_status = False) Then
      Exit Function
    End If
    
    ' Check if any error getting tube state and count
    If (get_auto_sampler_tube_state(AUTO_SMPLR_SMPL_WIN_POS) = False) Then
      Exit Function
    End If
    
    m_autoSamplrOperState = ASOS_IDLE
    survey_sample_win = True
  End If
End Function

Private Sub chk_4_any_stray_cups()
  Dim res As VbMsgBoxResult
  Dim rc As Boolean
   
  ' Check if any stray cups at entry zone, exit zone and/or sample window
  If (AutoSmplrTowerNumCups(AUTO_SMPLR_ENTRY_ZONE_POS) > 0) Or _
     (AutoSmplrTowerNumCups(AUTO_SMPLR_EXIT_ZONE_POS) > 0) Or _
     (AutoSmplrTowerNumCups(AUTO_SMPLR_SMPL_WIN_POS) > 0) Then
    res = CWrap.ShowMessageBoxW(MLSupport.GSS("frm_batchRun", "promptMsg1", "Please remove sample cups from Entry Zone, Exit Zone and/or Sample Window and press OK to continue. Otherwise press CANCEL to abort batch run!"), vbOKCancel)
    
    ' Check if user wants to abort batch run
    If (res = vbCancel) Then
      cmd_abortBatch.Visible = False
      m_batchTmrState = BTS_ABORT_BATCH
      Exit Sub
    Else
      rc = True
      
      ' Clear tube cup count for any stray cups at entry zone
      If (AutoSmplrTowerNumCups(AUTO_SMPLR_ENTRY_ZONE_POS) > 0) Then
        rc = set_auto_sampler_tube_cup_count(AUTO_SMPLR_ENTRY_ZONE_POS, 0)
      End If
        
      If (rc = True) Then
        ' Clear tube cup count for any stray cups at exit zone
        If (AutoSmplrTowerNumCups(AUTO_SMPLR_EXIT_ZONE_POS) > 0) Then
          rc = set_auto_sampler_tube_cup_count(AUTO_SMPLR_EXIT_ZONE_POS, 0)
        End If
      End If
          
      If (rc = True) Then
        ' Clear tube cup count for any stray cups at sample window
        If (AutoSmplrTowerNumCups(AUTO_SMPLR_SMPL_WIN_POS) > 0) Then
          rc = set_auto_sampler_tube_cup_count(AUTO_SMPLR_SMPL_WIN_POS, 0)
        End If
      End If
      
      If (rc = False) Then
        m_batchTmrState = BTS_BATCH_ERR
        Exit Sub
      End If
    End If
  End If
  
  m_batchTmrState = BTS_SMPL_INIT
End Sub

Private Sub setup_sample()
  Dim ii As Integer
  Dim inputStrg As String

  m_currentBatchSmplNum = m_currentSmplWinSmplInfo.sampleNum
  m_prodIni = BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).prodIni
     
  ' Check if need to load product ini file
  If (m_prodIni <> unity_main.current_ini) Then
    unity_main.load_prod_file m_prodIni, True
          
    ' Check if any error loading product file
    If (unity_main.lblProd1.Caption = "") Then
      BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).smplStatus = BSS_PROD_ERROR
      report_err_status m_currentSmplWinSmplInfo.sampleNum, MLSupport.GSS("frm_batchRun", "errMsg2", "Cannot load product file")
      unload_sample m_currentSmplWinSmplInfo.sampleNum
      Exit Sub
    End If
  End If
  
  ' Check if current product has PRD model type
  If (unity_main.m_prdModelType = True) Then
    ' Check if .NET Framework 2.0 is not installed
    If (unity_main.m_netFWInstalled = False) Then
      unity_main.errorstring = "Problem with product trying to save spectra data to UCal SVF file without MS .NET Framework 2.0 installed"
      unity_main.write_error
      BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).smplStatus = BSS_PRED_ERROR
      report_err_status m_currentSmplWinSmplInfo.sampleNum, MLSupport.GSS("frm_batchRun", "errMsg1", ".NET Framework Required")
      unload_sample m_currentSmplWinSmplInfo.sampleNum
      Exit Sub
    End If
  End If
    
  ' Check if need to perform reference before each sample
  If (unity_main.m_backFreq = REF_FREQ_ALL_SMPLS) Then
    ' Start reference scan
    setup_ref_scan
  Else
    ' Check if need to perform reference scan
    If (unity_main.m_intRefAutoScan = True) Then
      ' Start reference scan
      setup_ref_scan
    Else
      ' Setup for sample scan
      setup_sample_scan
    End If
  End If
End Sub

Private Sub setup_ref_scan()
    
  lbl_scanProgress.Caption = MLSupport.GSS("frm_batchRun", "RefScanProgress", "Reference Scan Progress")
  display_sample_info m_currentSmplWinSmplInfo.sampleNum, BATCH_STATUS_COL, MLSupport.GSS("frm_batchRun", "statMsg7", "Scanning Reference")
    
  ' Start reference scan
  m_scanFlg = 1
  m_scanAttempts = 2
  unity_main.m_scanDataType = SDT_PRODINTREF
  unity_main.m_scanTmrState = STS_SETUP
  unity_main.tmr_ref.enabled = True

  m_batchTmrState = BTS_UNLOAD_QUE_SMPL
End Sub

Private Sub setup_sample_scan()
  Dim ii As Integer
  Dim inputStrg As String

  m_scanFlg = 0
  m_scanAttempts = 2
  unity_main.prepscan
  unity_main.clearpredtable
    
  ' Check if sample ID defined for current sample
  If (BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).sampleId <> "") Then
    unity_main.txtsamplename.Text = BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).sampleId
  Else
    If (frm_collect.get_samp_name = False) Then
      BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).smplStatus = BSS_SMPL_ID_ERROR
      report_err_status m_currentSmplWinSmplInfo.sampleNum, MLSupport.GSS("frm_batchRun", "errMsg5", "Cannot obtain sample ID")
      unload_sample m_currentSmplWinSmplInfo.sampleNum
      Exit Sub
    End If
  End If
    
  display_sample_info m_currentSmplWinSmplInfo.sampleNum, BATCH_SAMPLE_ID_COL, unity_main.txtsamplename.Text
  
  ' Setup sample comment
  unity_main.txtsampcomment.Text = BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).Comment
  
  ' Check if user inputs configured for product
  If (BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).usrInputFlg = True) Then
    ' Setup each user input
    For ii = 1 To MAX_MAN_INPUTS
      ' Setup for input enable field
      frm_buttoncfg.ss_buttonconfig.Col = ii
      frm_buttoncfg.ss_buttonconfig.Row = 1
      
      ' Check if input enabled
      If (frm_buttoncfg.ss_buttonconfig.Value = 1) Then
        inputStrg = BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).usrInputs(ii)
      Else
        inputStrg = ""
      End If
      
      ' Setup for text entry/list box selection field
      frm_buttoncfg.ss_buttonconfig.Col = ii
      frm_buttoncfg.ss_buttonconfig.Row = 2
  
      ' Check if using text entry
      If (frm_buttoncfg.ss_buttonconfig.Value = 0) Then
        frm_scanname.txtbx(ii).Text = inputStrg
      Else    ' Using list
        frm_scanname.combo(ii).Text = inputStrg
      End If
    Next ii
  End If
    
  lbl_scanProgress.Caption = MLSupport.GSS("frm_batchRun", "SmplScanProgress", "Sample Scan Progress")
  display_sample_info m_currentSmplWinSmplInfo.sampleNum, BATCH_STATUS_COL, MLSupport.GSS("frm_batchRun", "statMsg8", "Scanning Sample")
  
  ' Start sample scan
  m_scanFlg = 2
  unity_main.m_smplRepacks = 1
  unity_main.m_scanDataType = SDT_PRODSMPL
  unity_main.m_scanTmrState = STS_SETUP
  unity_main.tmr_sample.enabled = True

  m_batchTmrState = BTS_UNLOAD_QUE_SMPL
End Sub

Private Sub abort_scan()
  Dim nn As Integer
  Dim StartTime As Single
  Dim errMsg As String
  Dim uniMsg As String

  ' Exit if not scanning
  If (m_scanFlg = 0) Then
    Exit Sub
  End If
  
  ' Wait if scan is in setup mode
  While (unity_main.m_scanTmrState = STS_SETUP)
    DoEvents
  Wend
  
  ' Check if scan is running
  If (unity_main.m_scanTmrState = STS_WAIT_CMP) Then
    DoEvents
  
    ' Clear any previous errors
    Clear_MS11_Error_Codes
  
#If SSRCS Then
    SSRCSClientError = unity_main.SSRCSClient.ScanStop()
    
    If (SSRCSClientError = 0) Then
#Else
    If (unity_main.MS11srv.ScanStop() = True) Then
#End If
      unity_main.m_scanState = SS_ABORT
      uniMsg = MLSupport.GSS("frm_status", "statMsg1", "User aborted scan")
      Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "User aborted scan", uniMsg)
      
      StartTime = Timer
      
      ' Delay to allow instrument to actually stop scan
      Do While (True)
        ' Check scan state to see if not actively scanning
#If SSRCS Then
        Dim scanState As Long
        SSRCSClientError = unity_main.SSRCSClient.GetScanState(scanState)
        
        If (scanState <> 5) Then GoTo CLEAR_SCAN

        If (unity_main.chk_timeout(StartTime, 8) = True) Then
#Else
        If (unity_main.MS11srv.scanState <> 5) Then GoTo CLEAR_SCAN

        If (unity_main.chk_timeout(StartTime, 5) = True) Then
#End If
          Exit Do
        End If
    
        DoEvents
      Loop
      
CLEAR_SCAN:
#If SSRCS Then
      SSRCSClientError = unity_main.SSRCSClient.ScanDataClr()
    
      If (SSRCSClientError <> 0) Then
#Else
      If (unity_main.MS11srv.ScanDataClr() = False) Then
#End If
        ' Report error codes
        uniMsg = MLSupport.GSS("OperStatus", "status27", "Error clearing scan data")
        Call frm_status.report_error_codes("Error clearing scan data", uniMsg)
      End If
    Else
      
    End If
  End If
  
  ' Wait until scan is fully completed or stopped due to error
  While ((unity_main.m_scanTmrState <> STS_COMPLETED) And (unity_main.m_scanTmrState <> STS_COMPLETED_AS) And (unity_main.m_scanTmrState <> STS_ABORT))
    DoEvents
  Wend
    
  If (unity_main.m_scanTmrState = STS_COMPLETED) Then
    ' Save sample scan result data
    save_scan_data
  End If
  
  m_scanFlg = 0
End Sub

Private Function abort_clean_up() As Boolean
  Dim fetchCupNum As Integer
  Dim cupPos As Integer
  
  ' Get current cup count for sample window, entry and exit zones
  If (get_auto_sampler_tube_cup_count(AUTO_SMPLR_SMPL_WIN_POS) = True) Then
    If (get_auto_sampler_tube_cup_count(AUTO_SMPLR_ENTRY_ZONE_POS) = True) Then
      If (get_auto_sampler_tube_cup_count(AUTO_SMPLR_EXIT_ZONE_POS) = True) Then
        ' Check if arm has sample cup in jaws
        If (get_auto_sampler_cup_pos = True) Then
          ' Check if switch is activated
          If (AutoSmplrSwitchStatus = 1) Then
            ' Check if jaws are activated
            If (AutoSmplrJawStatus = 1) Then
              ' Determine where robot was going to dump cup
              Select Case (m_dumpPosIndx)
                Case AUTO_SMPLR_SMPL_WIN_POS            ' sample window
                  ' Put sample cup back in loading tower
                  If (restack_sample(m_currentBatchSmplNum, AUTO_SMPLR_UNKNOWN_POS) = False) Then
                    Exit Function
                  End If

                Case AUTO_SMPLR_ENTRY_ZONE_POS          ' entry zone
                  ' Put sample cup back in loading tower
                  If (restack_sample(m_currentBatchSmplNum, AUTO_SMPLR_UNKNOWN_POS) = False) Then
                    Exit Function
                  End If
              
                Case AUTO_SMPLR_EXIT_ZONE_POS           ' exit zone
                  ' Unload sample cup
                  If (unload_aborted_sample(m_currentBatchSmplNum, AUTO_SMPLR_UNKNOWN_POS, m_currentSmplWinSmplInfo.fetchCupNum) = False) Then
                    Exit Function
                  End If
                  
                Case Else                               ' unload tower
                  ' Check if unloading from exit zone
                  If (m_fetchPosIndx = AUTO_SMPLR_EXIT_ZONE_POS) Then
                    fetchCupNum = m_currentExitZoneSmplInfo.fetchCupNum
                  Else
                    fetchCupNum = m_currentSmplWinSmplInfo.fetchCupNum
                  End If
              
                  ' Unload sample cup
                  If (unload_aborted_sample(m_currentBatchSmplNum, AUTO_SMPLR_UNKNOWN_POS, fetchCupNum) = False) Then
                    Exit Function
                  End If
              End Select
            Else    ' Reposition/raise arm to deactivate switch
              cupPos = AutoSmplrCupPosition
              
              If (setup_cup_pos(cupPos, UNKNOWN_CTRL) = False) Then
                Exit Function
              End If
            End If
          End If

          ' Check if sample cup left on exit zone
          If (AutoSmplrTowerNumCups(AUTO_SMPLR_EXIT_ZONE_POS) > 0) Then
            ' Unload sample cup
            If (unload_aborted_sample(m_currentExitZoneSmplInfo.sampleNum, AUTO_SMPLR_EXIT_ZONE_POS, m_currentExitZoneSmplInfo.fetchCupNum) = False) Then
              Exit Function
            End If
          End If
        
          ' Check if any sample cup left on sample window
          If (AutoSmplrTowerNumCups(AUTO_SMPLR_SMPL_WIN_POS) > 0) Then
            If ((BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).smplStatus = BSS_IDLE) Or (BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).smplStatus = BSS_BATCH_ABORT)) Then
              ss_batchInfo.Row = m_currentSmplWinSmplInfo.sampleNum
              ss_batchInfo.Col = BATCH_SAMPLE_ID_COL

              ' Check if sample ID has been assigned via auto naming convention
              If ((ss_batchInfo.Text <> "") And (BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).sampleId = "")) Then
                frm_collect.roll_back_name_ctr
              End If

              display_sample_info m_currentSmplWinSmplInfo.sampleNum, BATCH_SAMPLE_ID_COL, ""
              BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).smplStatus = BSS_IDLE
              clear_status m_currentSmplWinSmplInfo.sampleNum

              ' Put sample cup back in loading tower
              If (restack_sample(m_currentSmplWinSmplInfo.sampleNum, AUTO_SMPLR_SMPL_WIN_POS) = False) Then
                Exit Function
              End If
            Else
              ' Unload sample cup
              If (unload_aborted_sample(m_currentSmplWinSmplInfo.sampleNum, AUTO_SMPLR_SMPL_WIN_POS, m_currentSmplWinSmplInfo.fetchCupNum) = False) Then
                Exit Function
              End If
            End If
          End If
        
          ' Check if sample cup left on entry zone
          If (AutoSmplrTowerNumCups(AUTO_SMPLR_ENTRY_ZONE_POS) > 0) Then
            ' Put sample cup back in loading tower
            If (restack_sample(m_currentEntryZoneSmplInfo.sampleNum, AUTO_SMPLR_ENTRY_ZONE_POS) = False) Then
              Exit Function
            End If
          End If
        
          If (reset_all_tube_states(True) = True) Then
            abort_clean_up = True
          End If
        End If
      End If
    End If
  End If
End Function

Private Sub load_sample(smplNum As Integer)
  Dim cupNumStrg As String

  m_currentBatchSmplNum = smplNum
  m_fetchPosIndx = BatchRunCfg(smplNum).loadTower - 1
  m_fetchCupNum = AutoSmplrTowerNumCups(m_fetchPosIndx)
  
  ' Check if sample cup queued on entry zone
  If (AutoSmplrTowerNumCups(AUTO_SMPLR_ENTRY_ZONE_POS) > 0) Then
    ' Place cup on sample window
    unity_main.m_scanTmrState = STS_COMPLETED_AS
    m_fetchPosIndx = AUTO_SMPLR_ENTRY_ZONE_POS
    m_dumpPosIndx = AUTO_SMPLR_SMPL_WIN_POS
    m_fetchCupNum = m_currentEntryZoneSmplInfo.fetchCupNum
    m_currentSmplWinSmplInfo.sampleNum = smplNum
    m_currentSmplWinSmplInfo.fetchCupNum = m_fetchCupNum
    cupNumStrg = m_fetchCupNum
    lbl_batchProgress2.Caption = MLSupport.GSS("frm_batchRun", "statMsg2", "Loading Sample Cup")
  Else   ' fetch sample cup from tower
    ' Check if no sample cup on sample window
    If (AutoSmplrTowerNumCups(AUTO_SMPLR_SMPL_WIN_POS) = 0) Then
      ' Place cup on sample window
      unity_main.m_scanTmrState = STS_COMPLETED_AS
      m_currentSmplWinSmplInfo.sampleNum = smplNum
      m_currentSmplWinSmplInfo.fetchCupNum = m_fetchCupNum
      m_dumpPosIndx = AUTO_SMPLR_SMPL_WIN_POS
      cupNumStrg = m_fetchCupNum
      lbl_batchProgress2.Caption = MLSupport.GSS("frm_batchRun", "statMsg2", "Loading Sample Cup")
    Else
      ' Place cup on entry zone
      m_currentEntryZoneSmplInfo.sampleNum = smplNum
      m_currentEntryZoneSmplInfo.fetchCupNum = m_fetchCupNum
      m_dumpPosIndx = AUTO_SMPLR_ENTRY_ZONE_POS
      cupNumStrg = MLSupport.GSS("frm_batchRun", "statMsg1", "Queued")
      lbl_batchProgress2.Caption = MLSupport.GSS("frm_batchRun", "statMsg17", "Queueing Entry Sample Cup")
    End If
    
    ' Setup tube for active fetch
    If (check_tube_state_setup(m_fetchPosIndx, ACTIVE_FETCH_TUBE) = False) Then
      Exit Sub
    End If
  End If
  
  ' Fetch cup
  If (setup_cup_pos(m_fetchPosIndx, FETCH_CUP) = True) Then
    ' Release cup
    If (setup_cup_pos(m_dumpPosIndx, RELEASE_CUP) = True) Then
      ' Update load cup number on screen
      display_sample_info smplNum, BATCH_LOAD_CUP_NUM_COL, cupNumStrg
      
      ' Check if new sample cup to scan
      If (m_dumpPosIndx = AUTO_SMPLR_SMPL_WIN_POS) Then
        m_batchTmrState = BTS_START_SCAN
      Else    ' queued on entry zone
        m_batchTmrState = BTS_UNLOAD_QUE_SMPL
      End If
    End If
  End If
End Sub

Private Sub unload_sample(smplNum As Integer)
  Dim dumpCupNum As Integer
  Dim fetchCupNum As Integer
  Dim cupNumStrg As String

  m_currentBatchSmplNum = smplNum
  m_dumpPosIndx = BatchRunCfg(smplNum).unloadTower - 1
  dumpCupNum = AutoSmplrTowerNumCups(m_dumpPosIndx) + 1
  
  If (dumpCupNum > MAX_AUTO_SMPLR_TOWER_CUPS) Then
    BatchRunCfg(smplNum).smplStatus = BSS_SMPL_UNLOAD_ERROR
    report_err_status smplNum, MLSupport.GSS("frm_batchRun", "errMsg7", "No Room to Unload Sample Cup")
    m_autoSamplrLastOperErr = MLSupport.GGS_Params("frm_batchRun.operErrMsg3", "Tower %1 has no room to unload sample cup", CStr(m_dumpPosIndx + 1))
    m_batchTmrState = BTS_BATCH_ERR
    Exit Sub
  End If
  
  ' Check if sample cup queued on exit zone
  If (AutoSmplrTowerNumCups(AUTO_SMPLR_EXIT_ZONE_POS) > 0) Then
    m_fetchPosIndx = AUTO_SMPLR_EXIT_ZONE_POS
    cupNumStrg = dumpCupNum
    lbl_batchProgress2.Caption = MLSupport.GSS("frm_batchRun", "statMsg3", "Unloading Sample Cup")
  Else   ' fetch sample cup from sample window
    m_fetchPosIndx = AUTO_SMPLR_SMPL_WIN_POS
    fetchCupNum = m_currentSmplWinSmplInfo.fetchCupNum
    
    ' Check if sample cup already queued on entry zone to be scanned
    If (AutoSmplrTowerNumCups(AUTO_SMPLR_ENTRY_ZONE_POS) > 0) Then
      m_dumpPosIndx = AUTO_SMPLR_EXIT_ZONE_POS
      cupNumStrg = MLSupport.GSS("frm_batchRun", "statMsg1", "Queued")
      lbl_batchProgress2.Caption = MLSupport.GSS("frm_batchRun", "statMsg18", "Queueing Exit Sample Cup")
    Else
      cupNumStrg = dumpCupNum
      lbl_batchProgress2.Caption = MLSupport.GSS("frm_batchRun", "statMsg3", "Unloading Sample Cup")
    End If
  End If
  
  ' Check if unloading sample to a dump tower
  If (m_dumpPosIndx <> AUTO_SMPLR_EXIT_ZONE_POS) Then
    ' Setup tube for active dump
    If (check_tube_state_setup(m_dumpPosIndx, ACTIVE_DUMP_TUBE) = False) Then
      Exit Sub
    End If
  Else
    m_currentExitZoneSmplInfo.sampleNum = smplNum
    m_currentExitZoneSmplInfo.fetchCupNum = fetchCupNum
  End If
  
  ' Fetch cup
  If (setup_cup_pos(m_fetchPosIndx, FETCH_CUP) = True) Then
    ' Release cup
    If (setup_cup_pos(m_dumpPosIndx, RELEASE_CUP) = True) Then
      ' Update unload cup number on screen
      display_sample_info smplNum, BATCH_UNLOAD_CUP_NUM_COL, cupNumStrg
      
      ' Check if sample cup unloaded into tower
      If (m_dumpPosIndx <> AUTO_SMPLR_EXIT_ZONE_POS) Then
        m_batchTmrState = BTS_SMPL_CMPL
      Else    ' queued on exit zone
        m_batchTmrState = BTS_LOAD_QUE_SMPL
      End If
    End If
  End If
End Sub

Private Function unload_aborted_sample(smplNum As Integer, fetchPosIndx As Integer, fetchCupNum As Integer) As Boolean
  Dim dumpCupNum As Integer
  
  m_currentBatchSmplNum = smplNum
  m_dumpPosIndx = BatchRunCfg(smplNum).unloadTower - 1
  dumpCupNum = AutoSmplrTowerNumCups(m_dumpPosIndx) + 1
  
  ' Setup tube for active dump
  If (check_tube_state_setup(m_dumpPosIndx, ACTIVE_DUMP_TUBE) = False) Then
    Exit Function
  End If
  
  lbl_batchProgress2.Caption = MLSupport.GSS("frm_batchRun", "statMsg3", "Unloading Sample Cup")
  
  ' Check if sample cup needs to be fetched
  If (fetchPosIndx <> AUTO_SMPLR_UNKNOWN_POS) Then
    m_fetchPosIndx = fetchPosIndx
    
    ' Fetch cup
    If (setup_cup_pos(m_fetchPosIndx, FETCH_CUP) = False) Then
      Exit Function
    End If
  End If
  
  If ((m_fetchPosIndx <> AUTO_SMPLR_SMPL_WIN_POS) Or (m_fetchPosIndx <> AUTO_SMPLR_EXIT_ZONE_POS)) Then
    ' Update load cup number on screen
    display_sample_info smplNum, BATCH_LOAD_CUP_NUM_COL, CStr(fetchCupNum)
  End If
    
  ' Release cup
  If (setup_cup_pos(m_dumpPosIndx, RELEASE_CUP) = True) Then
    ' Update unload cup number on screen
    display_sample_info smplNum, BATCH_UNLOAD_CUP_NUM_COL, CStr(dumpCupNum)
      
    If (BatchRunCfg(smplNum).smplStatus = BSS_SCAN_COMPLETE) Then
      BatchRunCfg(smplNum).smplStatus = BSS_SAMPLE_COMPLETE
      report_cmpl_status smplNum
    End If
    
    write_batch_report smplNum, frm_batchRunCfg.m_batchRptFile
    unload_aborted_sample = True
  End If
End Function

Private Function restack_sample(smplNum As Integer, fetchPosIndx As Integer) As Boolean
  
  m_currentBatchSmplNum = smplNum
  m_dumpPosIndx = BatchRunCfg(smplNum).loadTower - 1
  
  ' Setup tube for active dump
  If (check_tube_state_setup(m_dumpPosIndx, ACTIVE_DUMP_TUBE) = False) Then
    Exit Function
  End If
  
  lbl_batchProgress2.Caption = MLSupport.GSS("frm_batchRun", "statMsg19", "Restacking Sample Cup")
  
  ' Check if sample cup needs to be fetched
  If (fetchPosIndx <> AUTO_SMPLR_UNKNOWN_POS) Then
    m_fetchPosIndx = fetchPosIndx
    
    ' Fetch cup
    If (setup_cup_pos(m_fetchPosIndx, FETCH_CUP) = False) Then
      Exit Function
    End If
  End If
  
  ' Release cup
  If (setup_cup_pos(m_dumpPosIndx, RELEASE_CUP) = True) Then
    ' Clear load cup number on screen
    display_sample_info smplNum, BATCH_LOAD_CUP_NUM_COL, ""
    restack_sample = True
  End If
End Function

Private Function check_tube_state_setup(tubeIndx As Integer, tubeState As AutoSamplerCommCtrl.eASTubeStates) As Boolean

  ' Check if tube state active dump
  If (tubeState = ACTIVE_DUMP_TUBE) Then
    ' Check if tube has no room to unload cup
    If (AutoSmplrTowerNumCups(tubeIndx) >= MAX_AUTO_SMPLR_TOWER_CUPS) Then
      BatchRunCfg(m_currentBatchSmplNum).smplStatus = BSS_SMPL_UNLOAD_ERROR
      report_err_status m_currentBatchSmplNum, MLSupport.GSS("frm_batchRun", "errMsg7", "No Room to Unload Sample Cup")
      m_autoSamplrLastOperErr = MLSupport.GGS_Params("frm_batchRun.operErrMsg3", "Tower %1 has no room to unload sample cup", CStr(tubeIndx + 1))
      m_batchTmrState = BTS_BATCH_ERR
      Exit Function
    End If
  
    ' Check if tower is already active fetch
    If (AutoSmplrActFetchTower = tubeIndx) Then
      AutoSmplrActFetchTower = AUTO_SMPLR_UNKNOWN_POS
    End If
    
    ' Check if tower is already active dump
    If (AutoSmplrActDumpTower = tubeIndx) Then
      check_tube_state_setup = True
    Else
      ' Set new active dump tower
      check_tube_state_setup = setup_tube_state(tubeIndx, ACTIVE_DUMP_TUBE)
    End If
  Else        ' tube state active fetch
    ' Check if no cups to fetch
    If (AutoSmplrTowerNumCups(tubeIndx) = 0) Then
      BatchRunCfg(m_currentBatchSmplNum).smplStatus = BSS_SMPL_LOAD_ERROR
      report_err_status m_currentBatchSmplNum, MLSupport.GSS("frm_batchRun", "errMsg6", "No Sample Cup to Load")
      m_autoSamplrLastOperErr = MLSupport.GGS_Params("frm_batchRun.operErrMsg2", "Tower %1 has no sample cup to load", CStr(tubeIndx + 1))
      m_batchTmrState = BTS_BATCH_ERR
      Exit Function
    End If

    ' Check if tower is already active dump
    If (AutoSmplrActDumpTower = tubeIndx) Then
      AutoSmplrActDumpTower = AUTO_SMPLR_UNKNOWN_POS
    End If
    
    ' Check if tower is already active fetch
    If (AutoSmplrActFetchTower = tubeIndx) Then
      check_tube_state_setup = True
    Else
      ' Set new active fetch tower
      check_tube_state_setup = setup_tube_state(tubeIndx, ACTIVE_FETCH_TUBE)
    End If
  End If
End Function

Private Function setup_tube_state(tubeIndx As Integer, tubeState As AutoSamplerCommCtrl.eASTubeStates) As Boolean

  m_autoSamplrOperState = ASOS_SETUP_TUBE

  ' Check if command was processed by auto-sampler
  If (set_auto_sampler_tube_state(tubeIndx, tubeState) = True) Then
    ' Check if tube state is not what was requested
    If (AutoSmplrTowerState(tubeIndx) <> tubeState) Then
      ' Error if arm is not moving to tube
      If (AutoSmplrTowerState(tubeIndx) <> UNKNOWN_TUBE) Then
        m_autoSamplrLastOperErr = MLSupport.GGS_Params("frm_batchRun.operErrMsg1", "Tower %1 wrong state; received %2, expected %3", CStr(tubeIndx + 1), CStr(tubeState), CStr(AutoSmplrTowerState(tubeIndx)))
        Exit Function
      End If
      
      ' Wait for arm to complete movement and tube is setup
      If (wait_busy_status = False) Then
        Exit Function
      End If
      
      ' Check if any error getting tube state and count
      If (get_auto_sampler_tube_state(tubeIndx) = False) Then
        Exit Function
      End If
     
      ' Check if tube not set for proper state
      If (AutoSmplrTowerState(tubeIndx) <> tubeState) Then
        m_autoSamplrLastOperErr = MLSupport.GGS_Params("frm_batchRun.operErrMsg1", "Tower %1 wrong state; received %2, expected %3", CStr(tubeIndx + 1), CStr(tubeState), CStr(AutoSmplrTowerState(tubeIndx)))
        Exit Function
      End If
    End If
    
    m_autoSamplrOperState = ASOS_IDLE
    setup_tube_state = True
  End If
End Function

Private Function setup_cup_pos(cupPos As Integer, actCtrl As AutoSamplerCommCtrl.eASActCtrls) As Boolean
  Dim rc As Boolean

  ' Check if no cup control
  If (actCtrl = UNKNOWN_CTRL) Then
    m_autoSamplrOperState = ASOS_MOVE_ARM
  
    ' Send cup position command to auto-sampler
    rc = set_auto_sampler_cup_pos(cupPos, UNKNOWN_CTRL)
  Else
    If (actCtrl = FETCH_CUP) Then
      m_autoSamplrOperState = ASOS_FETCH_CUP
    Else
      m_autoSamplrOperState = ASOS_RELEASE_CUP
    End If
    
    ' Send cup position & cup control command to auto-sampler
    rc = set_auto_sampler_cup_pos(cupPos, actCtrl)
  End If
  
  ' Check if command was processed by auto-sampler
  If (rc = True) Then
    ' Check if cup position is not what was requested
    If (AutoSmplrCupPosition <> cupPos) Then
      ' Error if arm is not moving to cup position
      If (AutoSmplrCupPosition <> AUTO_SMPLR_UNKNOWN_POS) Then
        m_autoSamplrLastOperErr = MLSupport.GGS_Params("frm_batchRun.operErrMsg4", "Arm not at proper cup position; received %1, expected %2", CStr(AutoSmplrCupPosition), CStr(cupPos))
        Exit Function
      End If
     
      ' Wait for arm to complete movement to cup position/control
      If (wait_busy_status = False) Then
        Exit Function
      End If
      
      ' Check if any error getting cup position
      If (get_auto_sampler_cup_pos() = False) Then
        Exit Function
      End If
     
      ' Check if arm not at proper cup position
      If (AutoSmplrCupPosition <> cupPos) Then
        m_autoSamplrLastOperErr = MLSupport.GGS_Params("frm_batchRun.operErrMsg4", "Arm not at proper cup position; received %1, expected %2", CStr(AutoSmplrCupPosition), CStr(cupPos))
        Exit Function
      End If
      
      ' Get update of cup count for position
      If (get_auto_sampler_tube_state(cupPos) = False) Then
        Exit Function
      End If
    End If
    
    m_autoSamplrOperState = ASOS_IDLE
    setup_cup_pos = True
  End If
End Function

Private Function move_arm_away() As Boolean
  
  lbl_batchProgress2.Caption = MLSupport.GSS("frm_batchRun", "statMsg20", "Moving Arm Out of the Way")
  move_arm_away = setup_cup_pos(AUTO_SMPLR_OUT_WAY_POS, UNKNOWN_CTRL)
End Function

Private Function park_arm_away() As Boolean
  
  m_autoSamplrOperState = ASOS_PARK_ARM
  lbl_batchProgress2.Caption = MLSupport.GSS("frm_batchRun", "statMsg10", "Parking Arm Out of Away")
  
  ' Check if command was processed by auto-sampler
  If (park_auto_sampler_arm = True) Then
    ' Check if cup position is not what was requested
    If (AutoSmplrCupPosition <> AUTO_SMPLR_OUT_WAY_POS) Then
      ' Error if arm is not moving to cup position
      If (AutoSmplrCupPosition <> AUTO_SMPLR_UNKNOWN_POS) Then
        m_autoSamplrLastOperErr = MLSupport.GGS_Params("frm_batchRun.operErrMsg4", "Arm not at proper cup position; received %1, expected %2", CStr(AutoSmplrCupPosition), CStr(AUTO_SMPLR_UNKNOWN_POS))
        Exit Function
      End If
      
      ' Wait for arm to complete movement to cup position/control
      If (wait_busy_status = False) Then
        Exit Function
      End If
      
      ' Check if any error getting cup position
      If (get_auto_sampler_cup_pos() = False) Then
        Exit Function
      End If
     
      ' Check if arm not at proper cup position
      If (AutoSmplrCupPosition <> AUTO_SMPLR_OUT_WAY_POS) Then
        m_autoSamplrLastOperErr = MLSupport.GGS_Params("frm_batchRun.operErrMsg4", "Arm not at proper cup position; received %1, expected %2", CStr(AutoSmplrCupPosition), CStr(AUTO_SMPLR_UNKNOWN_POS))
        Exit Function
      End If
    End If
    
    m_autoSamplrOperState = ASOS_IDLE
    park_arm_away = True
  End If
End Function

Private Function wait_busy_status() As Boolean
  Dim uniMsg As String

  m_pollStatusState = ASPS_IN_PROGRESS

  ' Enable poll status timer
  tmr_pollStatus.enabled = True
      
  ' Wait for arm movement to complete
  While (m_pollStatusState = ASPS_IN_PROGRESS)
    ' Check if scanning
    If (m_scanFlg <> 0) Then
      Select Case (unity_main.m_scanTmrState)
        Case STS_COMPLETED          ' reference/sample scan completed
          clear_status m_currentSmplWinSmplInfo.sampleNum

          ' Check if completed reference scan
          If (m_scanFlg = 1) Then
            BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).smplStatus = BSS_IDLE
            uniMsg = MLSupport.GSS("frm_batchRun", "statMsg5", "Reference Scan Completed")
            lbl_batchProgress2.Caption = uniMsg
            display_sample_info m_currentSmplWinSmplInfo.sampleNum, BATCH_STATUS_COL, uniMsg
          Else
            If (unity_main.pukedonpred = True) Then
              BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).smplStatus = BSS_PRED_ERROR
              report_err_status m_currentSmplWinSmplInfo.sampleNum, MLSupport.GSS("OperStatus", "status21", "Model Failure")
            Else
              uniMsg = MLSupport.GSS("frm_batchRun", "statMsg6", "Sample Scan Completed")
              lbl_batchProgress2.Caption = uniMsg
              display_sample_info m_currentSmplWinSmplInfo.sampleNum, BATCH_STATUS_COL, uniMsg
              
              ' Save sample scan result data
              save_scan_data
            End If
          End If
        
          m_scanState = unity_main.m_scanTmrState
          unity_main.m_scanTmrState = STS_COMPLETED_AS
          
        Case STS_ABORT              ' scan was aborted due to user request or error
          ' Check if scan aborted due to SpectraStar error
          If (m_batchAbortFlg = 0) Then
            BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).smplStatus = BSS_SCAN_ERROR
            
            ' Check if reference scan failed
            If (m_scanFlg = 1) Then
              report_err_status m_currentSmplWinSmplInfo.sampleNum, MLSupport.GSS("frm_batchRun", "errMsg8", "Reference Scan Error")
            Else
              report_err_status m_currentSmplWinSmplInfo.sampleNum, MLSupport.GSS("frm_batchRun", "errMsg9", "Sample Scan Error")
            End If
          
            m_scanState = unity_main.m_scanTmrState
            unity_main.m_scanTmrState = STS_COMPLETED_AS
          End If
      End Select
    End If

    DoEvents
  Wend
  
  Select Case (m_pollStatusState)
    Case ASPS_COMPLETED_ESTOP       ' poll status word completed due to emergency stop
      m_batchTmrState = BTS_ESTOP_BATCH
    
    Case ASPS_COMPLETED_ABORT       ' poll status word completed due to abort
      m_batchTmrState = BTS_ABORT_BATCH
    
    Case ASPS_COMPLETED_ERR         ' poll status word completed with error
      ' Check if auto-sampler operational error
      If ((AutoSmplrStatus And &H8000) = &H8000) Then
        If (get_auto_sampler_err_report = True) Then
          m_autoSmplrError = AutoSmplrErrors(0)
          m_autoSamplrLastOperErr = "Error code: 0x" & Format(Hex(AutoSmplrErrors(0)), "0000")
          clear_auto_sampler_errs
        End If
      End If
          
      m_batchTmrState = BTS_BATCH_ERR
      
    Case ASPS_COMPLETED_GOOD        ' poll status word completed successfully
      wait_busy_status = True
  End Select
End Function

Private Function abort_auto_sampler_operations() As Boolean

  ' Wait for any command that is still in progress
  check_auto_sampler_comm_in_progress
  
  ' Send abort operations command to auto-sampler
  m_autoSamplrCmdState = ASCS_IN_PROGRESS
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.AbortASOps
#Else
  autoSmplr.abort_operations AUTO_SMPLR_CRC_USAGE
#End If
  
  ' Wait for auto-sampler response
  wait_auto_sampler_response
  
  ' Check if command was processed by auto-sampler
  If (m_autoSamplrCmdState = ASCS_COMPLETED_GOOD) Then
    abort_auto_sampler_operations = True
  Else
    m_batchTmrState = BTS_BATCH_ERR
  End If
End Function

Private Function clear_auto_sampler_errs() As Boolean

  ' Wait for any command that is still in progress
  check_auto_sampler_comm_in_progress
  
  ' Send clear error report command to auto-sampler
  m_autoSamplrCmdState = ASCS_IN_PROGRESS
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.ClrASErrors
#Else
  autoSmplr.clear_errors AUTO_SMPLR_CRC_USAGE
#End If

  ' Wait for auto-sampler response
  wait_auto_sampler_response
  
  ' Check if command was processed by auto-sampler
  If (m_autoSamplrCmdState = ASCS_COMPLETED_GOOD) Then
    clear_auto_sampler_errs = True
  Else
    m_batchTmrState = BTS_BATCH_ERR
  End If
End Function

Private Function get_auto_sampler_all_tube_states() As Boolean

  ' Wait for any command that is still in progress
  check_auto_sampler_comm_in_progress
  
  ' Send get number of tubes and tube states count command to auto-sampler
  m_autoSamplrCmdState = ASCS_IN_PROGRESS
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.GetAllASTubesState
#Else
  autoSmplr.get_all_tubes_state AUTO_SMPLR_CRC_USAGE
#End If
  
  ' Wait for auto-sampler response
  wait_auto_sampler_response
  
  ' Check if command was processed by auto-sampler
  If (m_autoSamplrCmdState = ASCS_COMPLETED_GOOD) Then
    get_auto_sampler_all_tube_states = True
  Else
    m_batchTmrState = BTS_BATCH_ERR
  End If
End Function

Private Function get_auto_sampler_cup_pos() As Boolean

  ' Wait for any command that is still in progress
  check_auto_sampler_comm_in_progress
  
  ' Send get cup position command to auto-sampler
  m_autoSamplrCmdState = ASCS_IN_PROGRESS
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.GetASCupPos
#Else
  autoSmplr.get_cup_pos AUTO_SMPLR_CRC_USAGE
#End If
  
  ' Wait for auto-sampler response
  wait_auto_sampler_response
  
  ' Check if command was processed by auto-sampler
  If (m_autoSamplrCmdState = ASCS_COMPLETED_GOOD) Then
    get_auto_sampler_cup_pos = True
  Else
    m_batchTmrState = BTS_BATCH_ERR
  End If
End Function

Private Function get_auto_sampler_dsr_state() As Boolean
  
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.GetASDSRState(m_dsrState)
  
  If (SSRCSClientError = 0) Then
    get_auto_sampler_dsr_state = True
  Else
    m_autoSamplrLastCommErr = MLSupport.GSS("frm_ssrcsConnect", "errMsg3", "Not connected to a SpectraStar RCS server")
  End If
  
#Else
  get_auto_sampler_dsr_state = autoSmplr.get_dsr_state(m_dsrState)
#End If
End Function

Private Function get_auto_sampler_err_report() As Boolean

  ' Wait for any command that is still in progress
  check_auto_sampler_comm_in_progress
  
  ' Send get error report command to auto-sampler
  m_autoSamplrCmdState = ASCS_IN_PROGRESS
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.GetASErrors
#Else
  autoSmplr.get_errors AUTO_SMPLR_CRC_USAGE
#End If
  
  ' Wait for auto-sampler response
  wait_auto_sampler_response
  
  ' Check if command was processed by auto-sampler
  If (m_autoSamplrCmdState = ASCS_COMPLETED_GOOD) Then
    get_auto_sampler_err_report = True
  Else
    m_batchTmrState = BTS_BATCH_ERR
  End If
End Function

Private Function get_auto_sampler_status(statWord As Integer) As Boolean

  ' Wait for any command that is still in progress
  check_auto_sampler_comm_in_progress
  
  ' Send get status word report command to auto-sampler
  m_autoSamplrCmdState = ASCS_IN_PROGRESS
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.GetASStatusWord(statWord)
#Else
  autoSmplr.get_status_word statWord, AUTO_SMPLR_CRC_USAGE
#End If
  
  ' Wait for auto-sampler response
  wait_auto_sampler_response
  
  ' Check if command was processed by auto-sampler
  If (m_autoSamplrCmdState = ASCS_COMPLETED_GOOD) Then
    get_auto_sampler_status = True
  Else
    m_batchTmrState = BTS_BATCH_ERR
  End If
End Function

Private Function get_auto_sampler_tube_cup_count(tubeIndx As Integer) As Boolean
  Dim ctrlWordIndx As Integer
  
  ' Wait for any command that is still in progress
  check_auto_sampler_comm_in_progress
  
  ' Send get tube cup count command to auto-sampler
  m_autoSmplrDestCupPos = AUTO_SMPLR_UNKNOWN_POS
  m_autoSamplrCmdState = ASCS_IN_PROGRESS
  ctrlWordIndx = AUTO_SMPLR_CUP_CNT_CTRL_WORD_OFFSET + tubeIndx
  
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.GetASCtrlWord(ctrlWordIndx)
#Else
  autoSmplr.get_ctrl_word ctrlWordIndx, AUTO_SMPLR_CRC_USAGE
#End If

  ' Wait for auto-sampler response
  wait_auto_sampler_response
  
  ' Check if command was processed by auto-sampler
  If (m_autoSamplrCmdState = ASCS_COMPLETED_GOOD) Then
    get_auto_sampler_tube_cup_count = True
  Else
    m_batchTmrState = BTS_BATCH_ERR
  End If
End Function

Private Function get_auto_sampler_tube_state(tubeIndx As Integer) As Boolean

  ' Wait for any command that is still in progress
  check_auto_sampler_comm_in_progress
  
  ' Send get tube state/cup count command to auto-sampler
  m_autoSamplrCmdState = ASCS_IN_PROGRESS
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.GetASTubeState(tubeIndx)
#Else
  autoSmplr.get_tube_state tubeIndx, AUTO_SMPLR_CRC_USAGE
#End If
  
  ' Wait for auto-sampler response
  wait_auto_sampler_response
  
  ' Check if command was processed by auto-sampler
  If (m_autoSamplrCmdState = ASCS_COMPLETED_GOOD) Then
    get_auto_sampler_tube_state = True
  Else
    m_batchTmrState = BTS_BATCH_ERR
  End If
End Function

Private Function get_auto_sampler_version() As Boolean

  ' Wait for any command that is still in progress
  check_auto_sampler_comm_in_progress
  
  ' Send get version info command to auto-sampler
  m_autoSamplrCmdState = ASCS_IN_PROGRESS
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.GetASVer
#Else
  autoSmplr.get_version_info AUTO_SMPLR_CRC_USAGE
#End If
  
  ' Wait for auto-sampler response
  wait_auto_sampler_response
  
  ' Check if command was processed by auto-sampler
  If (m_autoSamplrCmdState = ASCS_COMPLETED_GOOD) Then
    get_auto_sampler_version = True
  Else
    m_batchTmrState = BTS_BATCH_ERR
  End If
End Function

Private Function park_auto_sampler_arm() As Boolean

  ' Wait for any command that is still in progress
  check_auto_sampler_comm_in_progress
  
  ' Send park arm out of away command to auto-sampler
  m_autoSmplrDestCupPos = AUTO_SMPLR_UNKNOWN_POS
  m_autoSamplrCmdState = ASCS_IN_PROGRESS
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.ParkASArm
#Else
  autoSmplr.park_arm AUTO_SMPLR_CRC_USAGE
#End If
  
  ' Wait for auto-sampler response
  wait_auto_sampler_response
  
  ' Check if command was processed by auto-sampler
  If (m_autoSamplrCmdState = ASCS_COMPLETED_GOOD) Then
    m_autoSmplrDestCupPos = AUTO_SMPLR_OUT_WAY_POS
    park_auto_sampler_arm = True
  Else
    m_batchTmrState = BTS_BATCH_ERR
  End If
End Function

Private Function reboot_auto_sampler_system() As Boolean

  lbl_batchProgress2.Caption = MLSupport.GSS("frm_batchRun", "statMsg23", "Rebooting Auto-Sampler")

  ' Wait for any command that is still in progress
  check_auto_sampler_comm_in_progress
  
  ' Send reboot system command to auto-sampler
  m_autoSamplrCmdState = ASCS_IN_PROGRESS
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.RebootASSys
#Else
  autoSmplr.reboot_system AUTO_SMPLR_CRC_USAGE
#End If
  
  ' Wait for auto-sampler response
  wait_auto_sampler_response
  
  ' Check if command was processed by auto-sampler
  If (m_autoSamplrCmdState = ASCS_COMPLETED_GOOD) Then
    reboot_auto_sampler_system = True
  End If
End Function

Private Function set_auto_sampler_cup_pos(cupPos As Integer, actCtrl As AutoSamplerCommCtrl.eASActCtrls) As Boolean

  ' Wait for any command that is still in progress
  check_auto_sampler_comm_in_progress
  
  m_autoSmplrDestCupPos = AUTO_SMPLR_UNKNOWN_POS
  m_autoSamplrCmdState = ASCS_IN_PROGRESS
  
  ' Check if no cup control
  If (actCtrl = UNKNOWN_CTRL) Then
    ' Send set cup position command to auto-sampler
#If SSRCS Then
    SSRCSClientError = unity_main.SSRCSClient.SetASCupPos(cupPos)
#Else
    autoSmplr.set_cup_pos cupPos, AUTO_SMPLR_CRC_USAGE
#End If
  Else
    ' Send set cup position & control action command to auto-sampler
#If SSRCS Then
    SSRCSClientError = unity_main.SSRCSClient.SetASCupPosCtrl(cupPos, actCtrl)
#Else
    autoSmplr.set_cup_pos_ctrl cupPos, actCtrl, AUTO_SMPLR_CRC_USAGE
#End If
  End If
  
  ' Wait for auto-sampler response
  wait_auto_sampler_response
  
  ' Check if command was processed by auto-sampler
  If (m_autoSamplrCmdState = ASCS_COMPLETED_GOOD) Then
    m_autoSmplrDestCupPos = cupPos
    set_auto_sampler_cup_pos = True
  Else
    m_batchTmrState = BTS_BATCH_ERR
  End If
End Function

Private Function set_auto_sampler_tube_cup_count(tubeIndx As Integer, cupCount As Integer) As Boolean
  Dim ctrlWordIndx As Integer
  
  ' Wait for any command that is still in progress
  check_auto_sampler_comm_in_progress
  
  ' Send set tube cup count command to auto-sampler
  m_autoSmplrDestCupPos = AUTO_SMPLR_UNKNOWN_POS
  m_autoSamplrCmdState = ASCS_IN_PROGRESS
  ctrlWordIndx = AUTO_SMPLR_CUP_CNT_CTRL_WORD_OFFSET + tubeIndx
  
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.SetASCtrlWord(ctrlWordIndx, cupCount, False, 0)
#Else
  autoSmplr.set_ctrl_word ctrlWordIndx, cupCount, False, 0, AUTO_SMPLR_CRC_USAGE
#End If

  ' Wait for auto-sampler response
  wait_auto_sampler_response
  
  ' Check if command was processed by auto-sampler
  If (m_autoSamplrCmdState = ASCS_COMPLETED_GOOD) Then
    set_auto_sampler_tube_cup_count = True
  Else
    m_batchTmrState = BTS_BATCH_ERR
  End If
End Function

Private Function set_auto_sampler_tube_state(tubeIndx As Integer, tubeState As AutoSamplerCommCtrl.eASTubeStates) As Boolean

  ' Wait for any command that is still in progress
  check_auto_sampler_comm_in_progress
  
  ' Send set tube state command to auto-sampler
  m_autoSmplrDestCupPos = AUTO_SMPLR_UNKNOWN_POS
  m_autoSamplrCmdState = ASCS_IN_PROGRESS
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.SetASTubeState(tubeIndx, tubeState)
#Else
  autoSmplr.set_tube_state tubeIndx, tubeState, AUTO_SMPLR_CRC_USAGE
#End If
  
  ' Wait for auto-sampler response
  wait_auto_sampler_response
  
  ' Check if command was processed by auto-sampler
  If (m_autoSamplrCmdState = ASCS_COMPLETED_GOOD) Then
    m_autoSmplrDestCupPos = tubeIndx
    set_auto_sampler_tube_state = True
  Else
    m_batchTmrState = BTS_BATCH_ERR
  End If
End Function

Private Function survey_auto_sampler_arm_loc(armLoc As Integer) As Boolean

  ' Wait for any command that is still in progress
  check_auto_sampler_comm_in_progress
  
  ' Send survey sample window command to auto-sampler
  m_autoSmplrDestCupPos = AUTO_SMPLR_UNKNOWN_POS
  m_autoSamplrCmdState = ASCS_IN_PROGRESS
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.SurveyASArmLoc(armLoc)
#Else
  autoSmplr.survey_arm_loc armLoc, AUTO_SMPLR_CRC_USAGE
#End If
  
  ' Wait for auto-sampler response
  wait_auto_sampler_response
  
  ' Check if command was processed by auto-sampler
  If (m_autoSamplrCmdState = ASCS_COMPLETED_GOOD) Then
    m_autoSmplrDestCupPos = armLoc
    survey_auto_sampler_arm_loc = True
  Else
    m_batchTmrState = BTS_BATCH_ERR
  End If
End Function

Private Sub check_auto_sampler_comm_in_progress()

  ' Wait for any command that is still in progress
  While (m_autoSamplrCmdState = ASCS_IN_PROGRESS)
    DoEvents
  Wend

  DoEvents
End Sub

Private Sub wait_auto_sampler_response()

#If SSRCS Then
  While ((m_autoSamplrCmdState = ASCS_IN_PROGRESS) And (unity_main.m_ssrcsConnected = True))
    DoEvents
  Wend
  
  If (unity_main.m_ssrcsConnected = False) Then
    Dim errMsg As String
    errMsg = MLSupport.GSS("frm_batchRun", "errMsg10", "Auto-Sampler Communication Failure") & ". "
    errMsg = errMsg & MLSupport.GSS("frm_ssrcsConnect", "errMsg3", "Not connected to a SpectraStar RCS server")
    
    ' Show error message to user if not running batch
    If (unity_main.m_batchRunFlg = False) Then
      unity_main.errorstring = errMsg
      unity_main.write_error
      CWrap.ShowMessageBoxW errMsg, vbCritical
    End If
    
    m_autoSamplrLastCommErr = errMsg
    m_autoSamplrLastCommStat = COMMS_NOT_INIT
    m_autoSamplrCmdState = ASCS_COMPLETED_ERR
  End If
#Else
  ' Check if command still in progress
  While (m_autoSamplrCmdState = ASCS_IN_PROGRESS)
    DoEvents
  Wend
#End If
End Sub

Private Sub display_sample_info(ssRow As Integer, ssCol As Integer, ssData As String)
  
  ' Check if to move batch window to display currrent row info
  If (ssRow < ss_batchInfo.TopRow) Then
    ss_batchInfo.TopRow = ssRow
  Else
    If (ssRow > (ss_batchInfo.TopRow + NUM_ROWS_DISP)) Then
      ss_batchInfo.TopRow = ssRow - NUM_ROWS_DISP
    End If
  End If

  ss_batchInfo.Row = ssRow
  ss_batchInfo.Col = ssCol
  ss_batchInfo.Text = ssData
End Sub

Private Sub report_cmpl_status(smplNum As Integer)

  ss_batchInfo.Row = smplNum
  ss_batchInfo.Col = BATCH_STATUS_COL
  
  If (ss_batchInfo.ForeColor = vbBlack) Then
    ss_batchInfo.ForeColor = RGB(0, 128, 0)  ' dark green
    display_sample_info smplNum, BATCH_STATUS_COL, MLSupport.GSS("OperStatus", "status25", "Sample Completed")
  End If
  
  scanProgress.percent = 0
End Sub

Private Sub report_err_status(smplNum As Integer, errMsg As String)

  If (smplNum > 0) And (smplNum <= frm_batchRunCfg.m_numBatchSamples) Then
    ss_batchInfo.Row = smplNum
    ss_batchInfo.Col = BATCH_STATUS_COL
    ss_batchInfo.ForeColor = vbRed
    display_sample_info smplNum, BATCH_STATUS_COL, errMsg
  End If
End Sub

Private Sub clear_status(smplNum As Integer)

  ss_batchInfo.Row = smplNum
  ss_batchInfo.Col = BATCH_STATUS_COL
  ss_batchInfo.ForeColor = vbBlack
  display_sample_info smplNum, BATCH_STATUS_COL, ""
End Sub

Private Sub save_scan_data()
  Dim ii As Integer
  Dim tmpStrg As String

  BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).smplStatus = BSS_SCAN_COMPLETE

  ' Save Date-Time (military format)
  m_smplData.scanTime = unity_main.lbl_miltime.Caption
  
  ' Save User Inputs
  For ii = 1 To MAX_MAN_INPUTS
    If (CSVUserInputs(ii) = True) Then
      frm_buttoncfg.ss_buttonconfig.Col = ii
      frm_buttoncfg.ss_buttonconfig.Row = 1
      
      ' Check if input enabled
      If (unity_main.m_useMIV = True) And (frm_buttoncfg.ss_buttonconfig.Value = 1) Then
        ' Setup for text entry/list box selection field
        frm_buttoncfg.ss_buttonconfig.Col = ii
        frm_buttoncfg.ss_buttonconfig.Row = 2
  
        ' Check if using text entry
        If (frm_buttoncfg.ss_buttonconfig.Value = 0) Then
          tmpStrg = Trim(frm_scanname.txtbx(ii).Text)
        Else    ' Using list
          tmpStrg = Trim(frm_scanname.combo(ii).Text)
        End If
      Else
        tmpStrg = MLSupport.GSS("Headers", "na", "NA")
      End If
    Else
      tmpStrg = ""
    End If
    
    m_smplData.usrInputs(ii) = tmpStrg
  Next ii
  
  ' Save info for each product property
  m_smplData.numprops = Trim(frmedmod.numprops.Text)
  
  If (m_smplData.numprops > 0) Then
    For ii = 1 To m_smplData.numprops
      ' Save Property Name
      unity_main.fpspread_pred.Row = ii
      unity_main.fpspread_pred.Col = 1
      m_smplData.results(ii).Name = Trim(unity_main.fpspread_pred.Text)
    
      ' Save Property Value
      unity_main.fpspread_pred.Row = ii
      unity_main.fpspread_pred.Col = 2
      m_smplData.results(ii).predVal = Trim(unity_main.fpspread_pred.Text)
    
      ' Check if no prediction made
      If (unity_main.lstmd.ListCount = 0) Then
        tmpStrg = unity_main.m_noOLVal
      Else
        tmpStrg = unity_main.lstmd.List(ii - 1)
      End If
      
      ' Save Property M-Distance
      m_smplData.results(ii).mDistVal = tmpStrg
    
      ' Check if no prediction made
      If (unity_main.lstresrat.ListCount = 0) Then
        tmpStrg = unity_main.m_noOLVal
      Else
        tmpStrg = unity_main.lstresrat.List(ii - 1)
      End If
      
      ' Save Property S-Residual
      m_smplData.results(ii).sResidVal = tmpStrg
  
      ' Check if no prediction made
      If (unity_main.lst_qual.ListCount = 0) Then
        tmpStrg = unity_main.m_noOLVal
      Else
        tmpStrg = unity_main.lst_qual.List(ii - 1)
      End If
      
      ' Save Property Outlier
      m_smplData.results(ii).outLierVal = tmpStrg
  
      ' Check if no prediction made
      If (unity_main.lst_nd.ListCount = 0) Then
        tmpStrg = unity_main.m_noOLVal
      Else
        tmpStrg = unity_main.lst_nd.List(ii - 1)
      End If
      
      ' Save Property Neighborhood Distance
      m_smplData.results(ii).nDistVal = tmpStrg
    
      ' Check if no prediction made
      If (unity_main.lstint.ListCount = 0) Then
        tmpStrg = unity_main.m_noOLVal
      Else
        tmpStrg = unity_main.lstint.List(ii - 1)
      End If
      
      ' Save Property Intercept
      m_smplData.results(ii).interVal = tmpStrg
  
      ' Check if no prediction made
      If (unity_main.lstslope.ListCount = 0) Then
        tmpStrg = unity_main.m_noOLVal
      Else
        tmpStrg = unity_main.lstslope.List(ii - 1)
      End If
      
      ' Save Property Slope
      m_smplData.results(ii).slopeVal = tmpStrg
    Next ii
  End If
End Sub

Private Sub write_batch_report(smplNum As Integer, batchRptFile As String)
  Dim fileName As String
  Dim printStrg As String
  Dim ii As Integer
  Dim errMsg As String
  Dim uniFile As New clsUniFile
  Dim uniMsg As String

  If (smplNum < 1) Or (smplNum > frm_batchRunCfg.m_numBatchSamples) Then
    Exit Sub
  End If

  ss_batchInfo.Row = smplNum

  ' Save batch report file path/name
  fileName = batchRptFile

  ' Add Date-Time (military format)
  printStrg = (Chr(34) & m_smplData.scanTime & Chr(34))

  ' Add System Serial Number
  printStrg = printStrg & "," & (Chr(34) & unity_main.m_sysSerialNum & Chr(34))

  ' Add Product
  ss_batchInfo.Col = BATCH_PROD_NAME_COL
  printStrg = printStrg & "," & (Chr(34) & ss_batchInfo.Text & Chr(34))
  
  ' Add Sample ID
  ss_batchInfo.Col = BATCH_SAMPLE_ID_COL
  printStrg = printStrg & "," & (Chr(34) & ss_batchInfo.Text & Chr(34))
  
  ' Add Comment
  printStrg = printStrg & "," & (Chr(34) & BatchRunCfg(smplNum).Comment & Chr(34))

  ' Add Status
  ss_batchInfo.Col = BATCH_STATUS_COL
  printStrg = printStrg & "," & (Chr(34) & ss_batchInfo.Text & Chr(34))

  ' Add User Inputs
  For ii = 1 To MAX_MAN_INPUTS
    If (CSVUserInputs(ii) = True) Then
      printStrg = printStrg & "," & (Chr(34) & m_smplData.usrInputs(ii) & Chr(34))
    End If
  Next ii
  
  ' Add load Tower
  ss_batchInfo.Col = BATCH_LOAD_TOWER_COL
  printStrg = printStrg & "," & (Chr(34) & ss_batchInfo.Text & Chr(34))

  ' Add Cup Number
  ss_batchInfo.Col = BATCH_LOAD_CUP_NUM_COL
  printStrg = printStrg & "," & (Chr(34) & ss_batchInfo.Text & Chr(34))

  ' Add Unload Tower
  ss_batchInfo.Col = BATCH_UNLOAD_TOWER_COL
  printStrg = printStrg & "," & (Chr(34) & ss_batchInfo.Text & Chr(34))

  ' Add Cup Number
  ss_batchInfo.Col = BATCH_UNLOAD_CUP_NUM_COL
  printStrg = printStrg & "," & (Chr(34) & ss_batchInfo.Text & Chr(34))

  If (m_smplData.numprops > 0) Then
    ' Add info for each product property
    For ii = 1 To m_smplData.numprops
      ' Add Property Name
      printStrg = printStrg & "," & (Chr(34) & m_smplData.results(ii).Name & Chr(34))
    
      ' Add Property Value
      printStrg = printStrg & "," & (Chr(34) & m_smplData.results(ii).predVal & Chr(34))
    
      ' Add Property M-Distance
      printStrg = printStrg & "," & (Chr(34) & m_smplData.results(ii).mDistVal & Chr(34))
      
      ' Add Property S-Residual
      printStrg = printStrg & "," & (Chr(34) & m_smplData.results(ii).sResidVal & Chr(34))
      
      ' Add Property Outlier
      printStrg = printStrg & "," & (Chr(34) & m_smplData.results(ii).outLierVal & Chr(34))
        
      ' Add Property Neighborhood Distance
      printStrg = printStrg & "," & (Chr(34) & m_smplData.results(ii).nDistVal & Chr(34))
      
      ' Add Property Intercept
      printStrg = printStrg & "," & (Chr(34) & m_smplData.results(ii).interVal & Chr(34))
  
      ' Add Property Slope
      printStrg = printStrg & "," & (Chr(34) & m_smplData.results(ii).slopeVal & Chr(34))
    Next ii
  End If

  m_smplData.scanTime = ""
  m_smplData.numprops = 0
  
  For ii = 1 To MAX_MAN_INPUTS
    m_smplData.usrInputs(ii) = ""
  Next ii
  
  uniMsg = MLSupport.GGS_Params("frm_batchRun.statMsg5", "Writing batch sample results to report file: %1", fileName)
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Writing btach sample results to report file: " & fileName), uniMsg)
  
  On Error GoTo FILE_ERROR
  CreatePath CFile.st_FilePath(fileName)
  
  ' Check if first batch sample
  If (smplNum = 1) Then
    ' Check if to append to existing batch report file
    If (frm_batchRunCfg.m_batchRptAppendFlg = True) Then
      GoTo APPEND_BATCH
    End If
    
    If (frm_batchRpt.create_batch_rpt_file(uniFile, fileName) = False) Then
      GoTo FILE_ERROR
    End If
  Else
APPEND_BATCH:
    If (uniFile.OpenFileAppend(fileName) = False) Then
      GoTo FILE_ERROR
    End If
  End If
  
  uniFile.WriteUnicodeLine printStrg
  uniFile.Flush
  uniFile.CloseFile
  
  unity_main.img_batchRpt.Visible = True
  Exit Sub

FILE_ERROR:
  uniFile.CloseFile
  errMsg = (fileName & " file write error. " & Error$)
  unity_main.errorstring = errMsg
  unity_main.write_error
End Sub

Private Sub batch_done(progStrg As String)

  tmr_batch.enabled = False
  cmd_exit.Visible = True
  lbl_batchProgress2.Caption = progStrg
  
  ' Check if batch report created
  If (CFile.st_FileExist(frm_batchRunCfg.m_batchRptFile) = True) Then
    cmd_batchReport.Visible = True
  End If
End Sub

Private Sub autoSmplr_cmdCompleted(ByVal cmdId As AutoSamplerCommCtrl.eASCmdIds, cmdIdTxt As String, ByVal commStat As AutoSamplerCommCtrl.eASCommStats, rspData As String)

  process_auto_sampler_response cmdId, cmdIdTxt, commStat, rspData
End Sub

Private Sub cmd_abortBatch_Click()
  
  unity_main.errorstring = "Batch Run Control screen Abort Batch button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  cmd_abortBatch.Visible = False
  
  ' Flag batch aborted for use request
  m_batchAbortFlg = 1
  lbl_batchProgress2.Caption = MLSupport.GSS("frm_batchRun", "statMsg15", "Aborting Batch Run")
  
  abort_auto_sampler_operations
  abort_scan
End Sub

Private Sub cmd_batchReport_Click()

  unity_main.errorstring = "Batch Run Control screen Batch Report button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  frm_batchRpt.show_report frm_batchRunCfg.m_batchRptFile
  frm_batchRpt.cmd_delete.Visible = False
  frm_batchRpt.Show 1
End Sub

Private Sub cmd_emergencyStop_Click()

  unity_main.errorstring = "Batch Run Control screen Emergency Stop button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  cmd_emergencyStop.Visible = False
  cmd_abortBatch.Visible = False

  ' Flag batch aborted for emergency stop
  m_batchAbortFlg = 2
  lbl_batchProgress2.Caption = MLSupport.GSS("frm_batchRun", "statMsg21", "Emergency Stop")
  
  abort_auto_sampler_operations
  abort_scan
End Sub

Private Sub cmd_exit_Click()
  Dim ii As Integer
  Dim rowCnt As Long

  unity_main.errorstring = "Batch Run Control screen Exit button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  unity_main.tmr_ref.enabled = False
  unity_main.tmr_sample.enabled = False
  unity_main.m_batchRunFlg = False
  
  For ii = 1 To frm_batchRunCfg.m_numBatchSamples
    rowCnt = rowCnt + 1
    frm_batchRunCfg.ss_batchCfg.Row = rowCnt
    frm_batchRunCfg.ss_batchCfg.Col = BATCH_CFG_PROD_INI_COL
  
    While (frm_batchRunCfg.ss_batchCfg.Text = "")
      rowCnt = rowCnt + 1
      frm_batchRunCfg.ss_batchCfg.Row = rowCnt
    Wend
  
    ' Set required batch config entry cells foreground color to indicate sample completion status
    Select Case (BatchRunCfg(ii).smplStatus)
      Case BSS_IDLE
        frm_batchRunCfg.set_cells_foregnd_color BATCH_CFG_LOAD_TOWER_COL, BATCH_CFG_PROD_NAME_COL, rowCnt, rowCnt, vbBlack
        
      Case BSS_SAMPLE_COMPLETE
        frm_batchRunCfg.set_cells_foregnd_color BATCH_CFG_LOAD_TOWER_COL, BATCH_CFG_PROD_NAME_COL, rowCnt, rowCnt, RGB(0, 128, 0)  ' dark green
      
      Case Else       ' error condition
        frm_batchRunCfg.set_cells_foregnd_color BATCH_CFG_LOAD_TOWER_COL, BATCH_CFG_PROD_NAME_COL, rowCnt, rowCnt, vbRed
    End Select
  Next ii
  
  ' Check if batch report created
  If (CFile.st_FileExist(frm_batchRunCfg.m_batchRptFile) = True) Then
    unity_main.m_batchRptFile = frm_batchRunCfg.m_batchRptFile
  End If
  
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.CloseASComms
#Else
  autoSmplr.close_comms
#End If

  unity_main.clear_GN_eventQ
  Unload frm_batchRun
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub

Private Sub tmr_batch_Timer()
  Dim uniMsg As String
  Dim rc As Boolean

  Select Case (m_batchTmrState)
    Case BTS_BATCH_INIT       ' initialize batch
      ' Initialize serial port communication w/ auto-sampler
#If SSRCS Then
      Dim parity As String
      parity = AUTO_SMPLR_PARITY
      SSRCSClientError = unity_main.SSRCSClient.InitASComms(unity_main.m_autoSmplrPort, AUTO_SMPLR_BAUD, parity, AUTO_SMPLR_RSP_TIMEOUT, AUTO_SMPLR_NUM_RETRIES, AUTO_SMPLR_CRC_USAGE)
      
      If (SSRCSClientError <> 0) Then
        If (unity_main.m_ssrcsConnected = True) Then
          m_autoSamplrLastCommErr = MLSupport.GGS_Params("ASErrMsg1", "Problem with configuring Auto-Sampler comm port %1. Verify no other communication option is configured to use this port", CStr(unity_main.m_autoSmplrPort))
        Else
          m_autoSamplrLastCommErr = MLSupport.GSS("frm_ssrcsConnect", "errMsg3", "Not connected to a SpectraStar RCS server")
        End If
#Else
      If (frm_batchRun.autoSmplr.init_comms(unity_main.m_autoSmplrPort, AUTO_SMPLR_BAUD, AUTO_SMPLR_PARITY, AUTO_SMPLR_RSP_TIMEOUT, AUTO_SMPLR_NUM_RETRIES) = False) Then
        m_autoSamplrLastCommErr = MLSupport.GGS_Params("ASErrMsg1", "Problem with configuring Auto-Sampler comm port %1. Verify no other communication option is configured to use this port", CStr(unity_main.m_autoSmplrPort))
#End If
        m_batchTmrState = BTS_BATCH_ERR
      Else
        ' Check if auto-sampler connected and running
        get_auto_sampler_dsr_state
        
        If (m_dsrState = False) Then
          tmr_batch.enabled = False
          CWrap.ShowMessageBoxW MLSupport.GSS("frm_batchRun", "errMsg14", "Batch run cannot be performed since since auto-sampler's DTR signal is off. Auto-sampler is either rebooting, powered off or not connected!"), vbCritical
          cmd_exit_Click
          Exit Sub
        End If
        
        ' Check if can clear auto-sampler error log
        If (clear_auto_sampler_errs = True) Then
          ' Check if can get auto-sampler version info
          If (get_auto_sampler_version = True) Then
            lbl_autoSamplrVer.Caption = "Firmware Ver: " & AutoSmplrVersion
            m_batchTmrState = BTS_SURVEY_TOWERS
          End If
        End If
      End If
      
    Case BTS_SURVEY_TOWERS    ' survey towers, entry & exit zones and sample window
      If (reset_all_tube_states(False) = True) Then
        ' Survey towers for number of cups
        If (survey_towers = True) Then
          ' Survey exit zone for any stray cup
          If (survey_exit_zone = True) Then
            ' Survey entry zone for any stray cup
            If (survey_entry_zone = True) Then
              ' Survey sample window for any stray cup
              If (survey_sample_win = True) Then
                ' Report any stray cups
                chk_4_any_stray_cups
              End If
            End If
          End If
        End If
      End If

    Case BTS_SMPL_INIT        ' initialize sample run
      m_numSamplesScan = 1
      clear_status m_numSamplesScan
      lbl_batchProgress2.Caption = MLSupport.GGS_Params("frm_batchRun.statMsg1", "Working with Sample %1", CStr(m_numSamplesScan))
      load_sample m_numSamplesScan
      
    Case BTS_START_SCAN       ' start sample scan
      setup_sample
      
    Case BTS_UNLOAD_QUE_SMPL  ' unload any queued completed sample
      ' Check if to unload any queued completed sample
      If (AutoSmplrTowerNumCups(AUTO_SMPLR_EXIT_ZONE_POS) > 0) Then
        unload_sample m_currentExitZoneSmplInfo.sampleNum
      Else
        m_batchTmrState = BTS_QUE_NEXT_SMPL
      End If
  
    Case BTS_QUE_NEXT_SMPL    ' queue next available sample for scanning
      ' Check if to queue next available sample for scanning
      If (AutoSmplrTowerNumCups(AUTO_SMPLR_ENTRY_ZONE_POS) = 0) And (m_numSamplesScan < frm_batchRunCfg.m_numBatchSamples) Then
        load_sample (m_currentSmplWinSmplInfo.sampleNum + 1)
      Else
        m_batchTmrState = BTS_LOAD_QUE_SMPL
      End If
      
    Case BTS_LOAD_QUE_SMPL    ' load any queued sample for scanning
      ' Check if to load any queued sample for scanning
      If (AutoSmplrTowerNumCups(AUTO_SMPLR_ENTRY_ZONE_POS) > 0) And (AutoSmplrTowerNumCups(AUTO_SMPLR_SMPL_WIN_POS) = 0) Then
        load_sample m_currentEntryZoneSmplInfo.sampleNum
      Else
        ' Check if any sample being scanning
        If (AutoSmplrTowerNumCups(AUTO_SMPLR_SMPL_WIN_POS) > 0) Then
          m_batchTmrState = BTS_SCAN_CMP
        Else
          m_batchTmrState = BTS_UNLOAD_QUE_SMPL
        End If
      End If
      
    Case BTS_SCAN_CMP         ' waiting for reference/sample scan to complete
      ' Check status of reference/sample scan
      Select Case (unity_main.m_scanTmrState)
        Case STS_COMPLETED_AS       ' reference/sample scan completed (auto-sampler)
          unity_main.m_scanTmrState = m_scanState
      
        Case STS_COMPLETED          ' reference/sample scan completed
          clear_status m_currentSmplWinSmplInfo.sampleNum

          ' Check if completed reference scan
          If (m_scanFlg = 1) Then
            uniMsg = MLSupport.GSS("frm_batchRun", "statMsg5", "Reference Scan Completed")
            lbl_batchProgress2.Caption = uniMsg
            display_sample_info m_currentSmplWinSmplInfo.sampleNum, BATCH_STATUS_COL, uniMsg

            ' Setup sample scan
            setup_sample_scan
          Else
            m_scanFlg = 0
            
            If (unity_main.pukedonpred = True) Then
              BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).smplStatus = BSS_PRED_ERROR
              report_err_status m_currentSmplWinSmplInfo.sampleNum, MLSupport.GSS("OperStatus", "status21", "Model Failure")
            Else
              uniMsg = MLSupport.GSS("frm_batchRun", "statMsg6", "Sample Scan Completed")
              lbl_batchProgress2.Caption = uniMsg
              display_sample_info m_currentSmplWinSmplInfo.sampleNum, BATCH_STATUS_COL, uniMsg
              
              ' Save sample scan result data
              save_scan_data
            End If
            
            ' Unload sample
            unload_sample m_currentSmplWinSmplInfo.sampleNum
          End If
      
        Case STS_ABORT              ' scan was aborted due to user request or error
          ' Check if scan aborted due to SpectraStar error
          If (m_batchAbortFlg = 0) Then
            BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).smplStatus = BSS_SCAN_ERROR
            
            ' Check if reference scan failed
            If (m_scanFlg = 1) Then
              report_err_status m_currentSmplWinSmplInfo.sampleNum, MLSupport.GSS("frm_batchRun", "errMsg8", "Reference Scan Error")
            Else
              report_err_status m_currentSmplWinSmplInfo.sampleNum, MLSupport.GSS("frm_batchRun", "errMsg9", "Sample Scan Error")
            End If
            
            m_scanFlg = 0
            
            ' Check if have another attempt to scan
            m_scanAttempts = m_scanAttempts - 1
            
            If (m_scanAttempts > 0) Then
              BatchRunCfg(m_currentSmplWinSmplInfo.sampleNum).smplStatus = BSS_IDLE
              clear_status m_currentSmplWinSmplInfo.sampleNum
              unity_main.m_scanTmrState = STS_SETUP
            Else
              unload_sample m_currentSmplWinSmplInfo.sampleNum
            End If
          Else    ' scan aborted due to batch abort request
            m_batchTmrState = BTS_ABORT_BATCH
          End If
      
        Case STS_WAIT_CMP           ' waiting for reference/sample scan to complete
          ' Check if waiting for referecne scan to complete
          If (m_scanFlg = 1) Then
            uniMsg = MLSupport.GSS("frm_batchRun", "statMsg7", "Scanning Reference")
          Else
            uniMsg = MLSupport.GSS("frm_batchRun", "statMsg8", "Scanning Sample")
          End If
          
          lbl_batchProgress2.Caption = uniMsg
          display_sample_info m_currentSmplWinSmplInfo.sampleNum, BATCH_STATUS_COL, uniMsg
          m_batchTmrState = BTS_POS_SMPL_WIN
      
        Case Else           ' all other scanning states
          m_batchTmrState = BTS_POS_SMPL_WIN
      End Select
  
    Case BTS_POS_SMPL_WIN     ' position arm over sample window while scanning
      ' Position arm over sample window if waiting for scan with no queue unload cup and
      ' either last sample or sample queued for scanning
      If ((AutoSmplrCupPosition <> AUTO_SMPLR_SMPL_WIN_POS) And (AutoSmplrTowerNumCups(AUTO_SMPLR_EXIT_ZONE_POS) = 0)) Then
        If ((m_numSamplesScan = frm_batchRunCfg.m_numBatchSamples) Or ((AutoSmplrTowerNumCups(AUTO_SMPLR_ENTRY_ZONE_POS) > 0))) Then
          If (setup_cup_pos(AUTO_SMPLR_SMPL_WIN_POS, UNKNOWN_CTRL) = True) Then
            m_batchTmrState = BTS_SCAN_CMP
          End If
          
          Exit Sub
        End If
      End If
      
      m_batchTmrState = BTS_UNLOAD_QUE_SMPL

    Case BTS_SMPL_CMPL        ' sample completed
      If (BatchRunCfg(m_numSamplesScan).smplStatus = BSS_SCAN_COMPLETE) Then
        BatchRunCfg(m_numSamplesScan).smplStatus = BSS_SAMPLE_COMPLETE
      End If
      
      report_cmpl_status m_numSamplesScan
      write_batch_report m_numSamplesScan, frm_batchRunCfg.m_batchRptFile
      m_numSamplesScan = m_numSamplesScan + 1
      
      ' Check if have completed all samples for batch run
      If (m_numSamplesScan > frm_batchRunCfg.m_numBatchSamples) Then
        ' Hide abort batch/emergency stop buttons
        cmd_abortBatch.Visible = False
        cmd_emergencyStop.Visible = False
        
        If (reset_all_tube_states(True) = False) Then
          Exit Sub
        End If
        
        batch_done MLSupport.GSS("frm_batchRun", "statMsg9", "Batch Completed")
      Else
        m_batchTmrState = BTS_QUE_NEXT_SMPL
      End If
      
    Case BTS_BATCH_ERR        ' auto-sampler operation/communication error
      Dim rebootFlg As Boolean
      
      ' Hide abort batch/emergency stop buttons
      cmd_abortBatch.Visible = False
      cmd_emergencyStop.Visible = False
      abort_scan
      
CLEANUP_ERR:
      unity_main.errorstring = "Batch run aborted due to error"
      unity_main.write_error
      
      ' Check if had auto-sampler communication error
      If (m_autoSamplrLastCommErr <> "") Then
        If (m_currentBatchSmplNum > 0) Then
          BatchRunCfg(m_currentBatchSmplNum).smplStatus = BSS_AS_COMM_ERROR
          report_err_status m_currentBatchSmplNum, MLSupport.GSS("frm_batchRun", "errMsg10", "Auto-Sampler Communication Failure")
        End If
        
        uniMsg = m_autoSamplrLastCommErr
      Else    ' auto-sampler operation error
        uniMsg = MLSupport.GSS("frm_batchRun", "errMsg11", "Auto-Sampler Operation Failure")
        
        If ((m_currentBatchSmplNum > 0) And (m_autoSmplrError <> 0)) Then
          BatchRunCfg(m_currentBatchSmplNum).smplStatus = BSS_AS_OPER_ERROR
          report_err_status m_currentBatchSmplNum, uniMsg
        End If
        
        uniMsg = uniMsg & ". " & m_autoSamplrLastOperErr
        
        ' Check if need to reboot robot
        If ((m_autoSmplrError And &HF) <> 0) Then
          reboot_auto_sampler_system
          uniMsg = uniMsg & ". " & MLSupport.GSS("frm_batchRun", "errMsg13", "Auto-sampler had to be rebooted to clear error")
          rebootFlg = True
        End If
      End If
      
      unity_main.errorstring = uniMsg
      unity_main.write_error
      
      uniMsg = uniMsg & ". " & MLSupport.GSS("frm_batchRun", "errMsg12", "Cannot continue with batch run, problem must be corrected")
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
      
      If (rebootFlg = True) Then
        If (m_currentBatchSmplNum > 0) Then
          write_batch_report m_currentBatchSmplNum, frm_batchRunCfg.m_batchRptFile
        End If
      Else
        If (m_currentBatchSmplNum > 0) Then
          If (BatchRunCfg(m_currentBatchSmplNum).smplStatus = BSS_SMPL_LOAD_ERROR) Then
            abort_clean_up
          Else
            write_batch_report m_currentBatchSmplNum, frm_batchRunCfg.m_batchRptFile
          End If
        End If
      
      reset_all_tube_states True
    End If
    
    batch_done MLSupport.GSS("frm_batchRun", "statMsg14", "Batch Run Aborted Due to Error")
              
    Case BTS_ABORT_BATCH      ' batch was aborted due to user
      ' Reenable auto-sampler movement
      m_batchAbortFlg = 0

      ' Check if batch aborted while running
      If (m_numSamplesScan > 0) Then
        ' Check if any error cleaning up abort
        If (abort_clean_up = False) Then
          ' Check if emergency stop
          If (m_batchAbortFlg = 2) Then
            Exit Sub
          Else
            cmd_emergencyStop.Visible = False
            GoTo CLEANUP_ERR
          End If
        End If
      Else
        ' Check if any error moving arm away
        If (move_arm_away = False) Then
          ' Check if emergency stop
          If (m_batchAbortFlg = 2) Then
            Exit Sub
          Else
            cmd_emergencyStop.Visible = False
            GoTo CLEANUP_ERR
          End If
        End If
      End If
      
      cmd_emergencyStop.Visible = False
      unity_main.errorstring = "Batch run aborted due to user request"
      unity_main.write_error

      batch_done MLSupport.GSS("frm_batchRun", "statMsg16", "Batch Run Aborted Due to User Request")
      
    Case BTS_ESTOP_BATCH        ' batch was stopped due to user emergency
      uniMsg = MLSupport.GSS("frm_batchRun", "promptMsg2", "Batch run was stopped for an emergency situation. Press OK when situation is resolved to allow auto-sampler to move arm arm out of the way.")
      CWrap.ShowMessageBoxW uniMsg, vbCritical
      
      ' Reenable auto-sampler movement
      m_batchAbortFlg = 0
      
      move_arm_away
      
      unity_main.errorstring = "Batch run stopped due to user emergency"
      unity_main.write_error

      batch_done MLSupport.GSS("frm_batchRun", "statMsg22", "Batch Run Stopped Due to User Emergency")
  End Select
End Sub

Private Sub tmr_pollStatus_Timer()
  
  ' Check if user wishes to abort batch run
  If (m_batchAbortFlg <> 0) Then
    If (m_batchAbortFlg = 1) Then
      m_pollStatusState = ASPS_COMPLETED_ABORT
    Else
      m_pollStatusState = ASPS_COMPLETED_ESTOP
    End If
  Else
    If (get_auto_sampler_status(0) = False) Then
      m_pollStatusState = ASPS_COMPLETED_ERR
    Else
      ' Check if auto-sampler still busy and has no operational error
      If (AutoSmplrStatus <> 0) And ((AutoSmplrStatus And &H8000) <> &H8000) Then
        m_pollStatusState = ASPS_IN_PROGRESS
        Exit Sub
      Else
        ' Check if not auto-sampler busy
        If (AutoSmplrStatus = 0) Then
          m_pollStatusState = ASPS_COMPLETED_GOOD
        Else
          ' Error flag set in status
          m_pollStatusState = ASPS_COMPLETED_ERR
        End If
      End If
    End If
  End If
  
  tmr_pollStatus.enabled = False
End Sub






