VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_status 
   Caption         =   "Scan Status"
   ClientHeight    =   5820
   ClientLeft      =   2055
   ClientTop       =   2955
   ClientWidth     =   8070
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
   Icon            =   "frm_status.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniButtonImageXP cmd_startScan 
      Height          =   705
      Left            =   480
      TabIndex        =   7
      Top             =   4080
      Width           =   2205
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
      Caption         =   "frm_status.frx":0442
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_status.frx":046C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_status.frx":048C
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   960
      Top             =   5160
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5820
      FormDesignWidth =   8070
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_TrayStatus 
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   5400
      Width           =   975
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_status.frx":04A8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   0
      MultiLine       =   0   'False
      Alignment       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frm_status.frx":04C8
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_status.frx":04E8
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_GenNotice 
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   4920
      Width           =   975
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_status.frx":0504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   0
      MultiLine       =   0   'False
      Alignment       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frm_status.frx":0524
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_status.frx":0544
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_exitScan 
      Cancel          =   -1  'True
      Height          =   705
      Left            =   5400
      TabIndex        =   2
      Top             =   4080
      Width           =   2205
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
      Caption         =   "frm_status.frx":0560
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_status.frx":0588
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_status.frx":05A8
   End
   Begin HexUniControls.ctlUniLabel lbl_ScanProgress 
      Height          =   255
      Left            =   2520
      Top             =   4920
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
      Caption         =   "frm_status.frx":05C4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_status.frx":05FE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_status.frx":061E
   End
   Begin HexUniControls.ctlProgressXP scanProgress 
      Height          =   375
      Left            =   1680
      Top             =   5280
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BackColor       =   16777215
      ForeColor       =   49152
      Border          =   -1  'True
      BorderSpace     =   -1  'True
      Spaces          =   -1  'True
      Tip             =   "frm_status.frx":063A
      Style           =   -1
      BackStyle       =   -1
      RoundedBorders  =   -1  'True
      MousePointer    =   0
      MouseIcon       =   "frm_status.frx":065A
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_abortScan 
      Height          =   705
      Left            =   2940
      TabIndex        =   1
      Top             =   4080
      Width           =   2205
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
      Caption         =   "frm_status.frx":0676
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_status.frx":06A0
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_status.frx":06C0
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_pauseScan 
      Height          =   705
      Left            =   480
      TabIndex        =   0
      Top             =   4080
      Width           =   2205
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
      Caption         =   "frm_status.frx":06DC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_status.frx":0706
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_status.frx":0726
   End
   Begin HexUniControls.ctlUniLabel lbl_status 
      Height          =   1575
      Left            =   240
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   2778
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_status.frx":0742
      BackColor       =   -2147483633
      ForeColor       =   12582912
      BorderColor     =   0
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_status.frx":0762
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_status.frx":0782
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_resumeScan 
      Height          =   705
      Left            =   480
      TabIndex        =   3
      Top             =   4080
      Width           =   2205
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
      Caption         =   "frm_status.frx":079E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_status.frx":07CA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_status.frx":07EA
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_retryScan 
      Height          =   705
      Left            =   2940
      TabIndex        =   4
      Top             =   4080
      Width           =   2205
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
      Caption         =   "frm_status.frx":0806
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_status.frx":0830
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_status.frx":0850
   End
   Begin HexUniControls.ctlUniLabel lbl_statusCmd 
      Height          =   1935
      Left            =   240
      Top             =   1920
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3413
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_status.frx":086C
      BackColor       =   -2147483633
      ForeColor       =   32768
      BorderColor     =   0
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_status.frx":088C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_status.frx":08AC
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancelScan 
      Height          =   705
      Left            =   2940
      TabIndex        =   8
      Top             =   4080
      Width           =   2205
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
      Caption         =   "frm_status.frx":08C8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_status.frx":08F4
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_status.frx":0914
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   240
      Top             =   4920
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
      Left            =   240
      Top             =   5280
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_status.frx":0930
   End
End
Attribute VB_Name = "frm_status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_scanAborted As Boolean

Public Sub report_error_codes(errMsg As String, uniMsg As String)
  Dim errCodeMsg As String

#If ABBFT Then
  Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
  frm_status.lbl_statusCmd.Caption = uniMsg
  frm_status.lbl_statusCmd.ForeColor = vbRed
#Else
  Call Get_MS11_Errorcodes_Msg(errCodeMsg)
  Call unity_main.log_scan_status(LOG_DBG_LEVEL1, (errMsg & " " & errCodeMsg), (uniMsg & " " & errCodeMsg))

  ' Check if not performing batch scan and not internal reference calibration function
  If ((unity_main.m_batchRunFlg = False) And (unity_main.m_intRefCalFlg = False)) Then
    frm_status.lbl_statusCmd.Caption = (uniMsg & vbCrLf & errCodeMsg)
    frm_status.lbl_statusCmd.ForeColor = vbRed
  End If
#End If
End Sub

Private Sub cmd_abortScan_Click()
  Dim errMsg As String
  Dim nn As Integer
  Dim StartTime As Single
  Dim uniMsg As String

  unity_main.errorstring = "Scan Status screen Abort button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  cmd_abortScan.Visible = False
  cmd_pauseScan.Visible = False

#If ABBFT Then
  If (unity_main.m_mb3000.m_scanType = NO_SCAN_TYPE) Then
    unity_main.m_mb3000.m_newRefReq = False
    unity_main.m_mb3000.m_newSmplReq = False
    Unload frm_status
  Else
    unity_main.m_mb3000.m_scanAborted = True
  End If
#Else
  ' Check if scan is running
  If (unity_main.m_scanState <> SS_PAUSE) Then
    ' Clear any previous errors
    Clear_MS11_Error_Codes
  
#If SSRCS Then
    SSRCSClientError = unity_main.SSRCSClient.ScanStop
    
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
      SSRCSClientError = unity_main.SSRCSClient.ScanDataClr
      
      If (SSRCSClientError <> 0) Then
#Else
      If (unity_main.MS11srv.ScanDataClr() = False) Then
#End If
        ' Report error codes
        uniMsg = MLSupport.GSS("OperStatus", "status27", "Error clearing scan data")
        Call report_error_codes("Error clearing scan data", uniMsg)
        cmd_exitScan.Visible = True
      End If
    Else
      ' Report error codes
      uniMsg = MLSupport.GSS("OperStatus", "status28", "Error stopping scan")
      Call report_error_codes("Error stopping scan", uniMsg)
      cmd_exitScan.Visible = True
    End If
  End If
#End If
End Sub

Private Sub cmd_cancelScan_Click()

  unity_main.errorstring = "Scan Status screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  unity_main.m_extRefPosition = 0
  cmd_cancelScan.Visible = False
  cmd_startScan.Visible = False
End Sub

Private Sub cmd_exitScan_Click()
  Dim ctr As Long
  
  ' Check if aborted product scan
  If (unity_main.m_scanDataType = SDT_PRODSMPL) Then
    frm_collect.roll_back_name_ctr
    unity_main.txtsamplename.Text = ""
    unity_main.txtsampcomment.Text = ""
  Else
    ' Check if aborted internal reference scan for every product sample
    If (unity_main.m_backFreq = REF_FREQ_ALL_SMPLS) And (unity_main.m_scanDataType = SDT_PRODINTREF) Then
      m_scanAborted = True
    End If
  End If
  
#If ABBFT Then
  unity_main.m_mb3000.m_scanType = NO_SCAN_TYPE
#End If
  
  Call unity_main.kill_loop(LOG_DBG_LEVEL1, True, "Scan Status screen Exit button selected")
  Unload frm_status
End Sub

#If SSTAR Then
Private Sub cmd_pauseScan_Click()
  Dim errMsg As String
  Dim uniMsg As String

  unity_main.errorstring = "Scan Status screen Pause button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  cmd_pauseScan.Visible = False
  cmd_abortScan.Visible = False

  ' Clear any previous errors
  Clear_MS11_Error_Codes
  
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.ScanStop
  
  If (SSRCSClientError = 0) Then
#Else
  If (unity_main.MS11srv.ScanStop() = True) Then
#End If
    unity_main.m_scanState = SS_PAUSE
    uniMsg = MLSupport.GSS("frm_status", "statMsg2", "User paused scan")
    Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "User paused scan", uniMsg)
  Else
    ' Report error codes
    uniMsg = MLSupport.GSS("OperStatus", "status29", "Error pausing scan")
    Call report_error_codes("Error pausing scan", uniMsg)
    cmd_abortScan.Visible = True
    cmd_exitScan.Visible = True
  End If
End Sub
#End If

#If SSTAR Then
Private Sub cmd_resumeScan_Click()
  Dim errMsg As String
  Dim uniMsg As String

  unity_main.errorstring = "Scan Status screen Resume button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  cmd_resumeScan.Visible = False

  ' Clear any previous errors
  Clear_MS11_Error_Codes

#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.ScanResume
  
  If (SSRCSClientError = 0) Then
#Else
  If (unity_main.MS11srv.ScanResume() = True) Then
#End If
    uniMsg = MLSupport.GSS("frm_status", "statMsg3", "User resumed scan")
    Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "User resumed scan", uniMsg)
  Else
    ' Report error codes
    uniMsg = MLSupport.GSS("OperStatus", "status30", "Error resuming scan")
    Call report_error_codes("Error resuming scan", uniMsg)
    cmd_exitScan.Visible = True
    cmd_retryScan.Visible = True
  End If
End Sub
#End If

Private Sub cmd_retryScan_Click()
  Dim uniMsg As String

  unity_main.errorstring = "Scan Status screen Retry button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  cmd_retryScan.Visible = False
  cmd_exitScan.Visible = False
  lbl_statusCmd.Caption = ""
  scanProgress.percent = 0
  
  uniMsg = MLSupport.GSS("frm_status", "statMsg4", "User retried scan")
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "User retried scan", uniMsg)
  unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status26", "Scan Retried by User")
  
#If ABBFT Then
  If (unity_main.m_mb3000.m_scanType = REF_SCAN_TYPE) Then
    unity_main.m_mb3000.m_newRefReq = True
  Else
    unity_main.m_mb3000.m_newSmplReq = True
  End If
#Else
  unity_main.m_scanState = SS_STOP
  unity_main.m_scanTmrState = STS_SETUP
#End If
End Sub

Private Sub cmd_startScan_Click()

  unity_main.errorstring = "Scan Status screen Start button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  unity_main.m_extRefPosition = 1
  cmd_cancelScan.Visible = False
  cmd_startScan.Visible = False
End Sub

Private Sub Form_Load()
  
  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
  unity_main.m_scanState = SS_STOP
  
  scanProgress.percent = 0
    
  ' Hide all buttons for now
  cmd_abortScan.Visible = False
  cmd_cancelScan.Visible = False
  cmd_exitScan.Visible = False
  cmd_pauseScan.Visible = False
  cmd_resumeScan.Visible = False
  cmd_retryScan.Visible = False
  cmd_startScan.Visible = False
  
  ' NOTE: SET TO INVISIBLE FOR RELEASE, VISIBLE FOR DEBUG ONLY
  txt_GenNotice.Visible = False
  txt_TrayStatus.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
#If SSTAR Then
  unity_main.clear_GN_eventQ
#End If
End Sub








