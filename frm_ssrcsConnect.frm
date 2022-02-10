VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_ssrcsConnect 
   Caption         =   "SpectraStar RCS Client Connection"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr_ssrcsConnect 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1080
      Top             =   4560
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   480
      Top             =   4560
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5175
      FormDesignWidth =   8070
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_shutdown 
      Height          =   645
      Left            =   3000
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1138
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
      Caption         =   "frm_ssrcsConnect.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_ssrcsConnect.frx":0030
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_ssrcsConnect.frx":0050
   End
   Begin HexUniControls.ctlUniLabel lbl_instStatus 
      Height          =   375
      Left            =   240
      Top             =   3480
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_ssrcsConnect.frx":006C
      BackColor       =   -2147483633
      ForeColor       =   32768
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_ssrcsConnect.frx":008C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_ssrcsConnect.frx":00AC
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_clear 
      Height          =   645
      Left            =   5520
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1138
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
      Caption         =   "frm_ssrcsConnect.frx":00C8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_ssrcsConnect.frx":00F2
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_ssrcsConnect.frx":0112
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   645
      Left            =   3000
      TabIndex        =   7
      Top             =   4200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1138
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
      Caption         =   "frm_ssrcsConnect.frx":012E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_ssrcsConnect.frx":015A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_ssrcsConnect.frx":017A
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_connectLocal 
      Height          =   645
      Left            =   5520
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1138
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
      Caption         =   "frm_ssrcsConnect.frx":0196
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_ssrcsConnect.frx":01D4
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_ssrcsConnect.frx":01F4
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_connectRemote 
      Height          =   645
      Left            =   5520
      TabIndex        =   1
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1138
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
      Caption         =   "frm_ssrcsConnect.frx":0210
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_ssrcsConnect.frx":0250
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_ssrcsConnect.frx":0270
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_ipAddr 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_ssrcsConnect.frx":028C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   0
      MultiLine       =   0   'False
      Alignment       =   2
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frm_ssrcsConnect.frx":02AC
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_ssrcsConnect.frx":02CC
   End
   Begin HexUniControls.ctlUniLabel lbl_ipAddr 
      Height          =   255
      Left            =   120
      Top             =   240
      Width           =   4860
      _ExtentX        =   8573
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
      Caption         =   "frm_ssrcsConnect.frx":02E8
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_ssrcsConnect.frx":0346
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_ssrcsConnect.frx":0366
   End
   Begin HexUniControls.ctlUniLabel lbl_connectTimeout 
      Height          =   735
      Left            =   720
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_ssrcsConnect.frx":0382
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_ssrcsConnect.frx":03CE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_ssrcsConnect.frx":03EE
   End
   Begin HexUniControls.ctlNumIncXP numInc_connectTimeout 
      Height          =   600
      Left            =   3240
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1058
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "10"
      Min             =   1
      Max             =   30
      AllowSpace      =   0   'False
      BorderColor     =   -1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ButtonBackColor =   -2147483633
      ButtonForeColor =   0
      ButtonStyle     =   4
      ButtonWidth     =   30
      Tip             =   ""
      MousePointer    =   0
      MouseIcon       =   "frm_ssrcsConnect.frx":040A
      TrapTabKey      =   0   'False
   End
   Begin HexUniControls.ctlUniLabel lbl_rspTimeout 
      Height          =   735
      Left            =   720
      Top             =   2280
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_ssrcsConnect.frx":0426
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_ssrcsConnect.frx":0474
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_ssrcsConnect.frx":0494
   End
   Begin HexUniControls.ctlNumIncXP numInc_rspTimeout 
      Height          =   600
      Left            =   3240
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1058
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "5"
      Min             =   1
      Max             =   10
      AllowSpace      =   0   'False
      BorderColor     =   -1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ButtonBackColor =   -2147483633
      ButtonForeColor =   0
      ButtonStyle     =   4
      ButtonWidth     =   30
      Tip             =   ""
      MousePointer    =   0
      MouseIcon       =   "frm_ssrcsConnect.frx":04B0
      TrapTabKey      =   0   'False
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   0
      Top             =   4560
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
Attribute VB_Name = "frm_ssrcsConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_startFlg As Boolean

Private m_ssrcsConnectTime As Integer
Private m_ssrcsIPAddr As String
Private m_ssrcsRspTime As Long
Private m_errMsgFlg As Boolean
Private m_connectTimeoutCtr As Integer

Public Sub initButtons(startFlg As Boolean)

  m_startFlg = startFlg
  
  If (m_startFlg = True) Then
    cmd_cancel.Visible = False
    cmd_shutdown.Visible = True
  Else
    cmd_cancel.Visible = True
    cmd_shutdown.Visible = False
  End If
End Sub

Private Sub connect_ssrcs(ipAddr As String, rspTime As Long, connectTime As Integer)

  lbl_instStatus.ForeColor = RGB(0, 128, 0)  ' dark green
  lbl_instStatus.Caption = MLSupport.GSS("frm_ssrcsConnect", "statMsg1", "Attempting to connect to SpectraStar RCS server")

  m_ssrcsIPAddr = ipAddr
  m_ssrcsRspTime = rspTime
  m_ssrcsConnectTime = connectTime
  
  unity_main.SSRCSClient.OpenSSRCS m_ssrcsIPAddr, MS11SRV_IF_VER, m_ssrcsRspTime
  m_connectTimeoutCtr = 0
  tmr_ssrcsConnect.enabled = True
End Sub

Private Sub cmd_cancel_Click()

  tmr_ssrcsConnect.enabled = False
  Unload frm_ssrcsConnect
  Call unity_main.restart_loop(LOG_DBG_LEVEL3, "SpectraStar RCS Client Connection screen Cancel button selected")
End Sub

Private Sub cmd_clear_Click()

  unity_main.errorstring = "SpectraStar RCS Client Connection screen Clear button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  txt_ipAddr.Text = ""
  txt_ipAddr.SetFocus
End Sub

Private Sub cmd_connectLocal_Click()

  unity_main.errorstring = "SpectraStar RCS Client Connection screen Connect Locally button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  connect_ssrcs LOCAL_SSRCS_IP_ADDR, LOCAL_SSRCS_RESPONSE_TIMEOUT, LOCAL_SSRCS_CONNECT_TIMEOUT
End Sub

Private Sub cmd_connectRemote_Click()

  unity_main.errorstring = "SpectraStar RCS Client Connection screen Connect Remotely button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  If (txt_ipAddr.Text = "") Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_ssrcsConnect", "errMsg2", "You must enter a SpectraStar IP Address/Hostname"), vbCritical
    txt_ipAddr.SetFocus
  Else
    connect_ssrcs txt_ipAddr.Text, (numInc_rspTimeout.Text * 1000), numInc_connectTimeout.Text
  End If
End Sub

Private Sub cmd_shutdown_Click()
  
  unity_main.errorstring = "SpectraStar RCS Client Connection screen Shutdown button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  tmr_ssrcsConnect.enabled = False
  frmSplash2.m_shutdownFlg = True
  Unload frm_ssrcsConnect
End Sub

Private Sub Form_Load()
    
  ' Apply language file to form
  MLSupport.ApplyToForm Me
    
  If (unity_main.m_ssrcsIPAddr <> LOCAL_SSRCS_IP_ADDR) Then
    txt_ipAddr.Text = unity_main.m_ssrcsIPAddr
    numInc_connectTimeout.Text = unity_main.m_ssrcsConnectTime
    numInc_rspTimeout.Text = unity_main.m_ssrcsRspTime / 1000
  End If
End Sub

Private Sub numInc_connectTimeout_DblClick()

  unity_main.formfrom = 20
  unity_main.varfrom = 2
  frm_numpad.lbl_num.Caption = lbl_connectTimeout.Caption
  frm_numpad.txt_num.Text = numInc_connectTimeout.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_rspTimeout_DblClick()

  unity_main.formfrom = 20
  unity_main.varfrom = 3
  frm_numpad.lbl_num.Caption = lbl_rspTimeout.Caption
  frm_numpad.txt_num.Text = numInc_rspTimeout.Text
  frm_numpad.Show 1
End Sub

Private Sub tmr_ssrcsConnect_Timer()
  Dim errBits As MASTER_ERRORS
  Dim err As MASTER_ERRORS
  Dim errCodeMsg As String
  Dim errMsg As String
  Dim uniMsg As String

  If (unity_main.m_ssrcsConnected = True) Then
    lbl_instStatus.Caption = MLSupport.GSS("frm_ssrcsConnect", "statMsg2", "Connected to SpectraStar RCS server")
    DoEvents
    
    ' Save connection info
    unity_main.m_ssrcsConnectTime = m_ssrcsConnectTime
    unity_main.m_ssrcsIPAddr = m_ssrcsIPAddr
    unity_main.m_ssrcsRspTime = m_ssrcsRspTime
    
    ' Check if system startup
    If (m_startFlg = True) Then
      tmr_ssrcsConnect.enabled = False
      Unload frm_ssrcsConnect
      Exit Sub
    End If
    
    If (MS11Initialized = True) Then
      tmr_ssrcsConnect.enabled = False
      Call frm_collect.savescansettings(False)
      Call frm_Inst.savemyinsts(False, False)
      Call frm_dynRptCfg.save_cfg(False, False)
      Call unity_main.system_startup(True)
      Unload frm_ssrcsConnect
    Else
'      MS11MasterErrors = unity_main.MS11srv.MasterErrors
      unity_main.SSRCSClient.GetMasterErrors MS11MasterErrors
      errBits = ME_CODE_DEAD + ME_MSCOM_DEAD + ME_MS11_INIT + ME_SCNR_FLD + ME_TRAY_FLD
      err = MS11MasterErrors And errBits
    
      ' Check if any major error bits set
      If (err <> ME_OKAY) Then
'        unity_main.MS11srv.MasterErrors = MS11MasterErrors
        unity_main.SSRCSClient.SetMasterErrors MS11MasterErrors
    
        ' Check if BIT(0) set, Operational Errors within Code prevent Running
        If ((err And ME_CODE_DEAD) <> 0) Then
          lbl_instStatus.ForeColor = vbRed
          errCodeMsg = "MErr = 0x" & Hex(MS11MasterErrors)
          uniMsg = MLSupport.GGS_Params("MS11srv.errMsg1", "SpectraStar has major operational error (%1)", errCodeMsg)
        Else
          lbl_instStatus.ForeColor = RGB(0, 128, 0)  ' dark green
        
          ' Check if BIT(4) set, Scanner is Out-of-Service (OOS) or similarly incapacitated
          If ((err And ME_SCNR_FLD) <> 0) Then
            uniMsg = MLSupport.GSS("MS11srv", "errMsg3", "Waiting for SpectraStar Scanner to Complete Reset")
          Else
            ' Check if BIT(5) set, Tray is Out-of-Service (OOS) or similarly incapacitated
            If ((err And ME_TRAY_FLD) <> 0) Then
              uniMsg = MLSupport.GSS("MS11srv", "errMsg4", "Waiting for SpectraStar Tray to Complete Reset")
            Else
              ' Check if BIT(3) set, MS1100 is Initializing and needs to comnplete before further operations take place
              If ((err And ME_MS11_INIT) <> 0) Then
                uniMsg = MLSupport.GSS("MS11srv", "errMsg5", "Waiting for SpectraStar Initialization to Complete")
              End If
            End If
          End If
      
          If (m_errMsgFlg = False) Then
            m_errMsgFlg = True
          Else
            m_errMsgFlg = False
            uniMsg = ""
          End If
        End If
        
        lbl_instStatus.Caption = uniMsg
      Else
        MS11Initialized = True
      End If
    End If
  Else
    m_connectTimeoutCtr = m_connectTimeoutCtr + 1
    
    If (m_connectTimeoutCtr > (m_ssrcsConnectTime * (1000 / tmr_ssrcsConnect.interVal))) Then
      uniMsg = MLSupport.GSS("frm_ssrcsConnect", "errMsg1", "Timed out trying to connect to SpectraStar RCS server")
      lbl_instStatus.ForeColor = vbRed
      tmr_ssrcsConnect.enabled = False
    Else
      lbl_instStatus.ForeColor = RGB(0, 128, 0)  ' dark green
    
      If (m_errMsgFlg = False) Then
        m_errMsgFlg = True
        uniMsg = MLSupport.GSS("frm_ssrcsConnect", "statMsg1", "Attempting to connect to SpectraStar RCS server")
      Else
        m_errMsgFlg = False
        uniMsg = ""
      End If
    End If
    
    lbl_instStatus.Caption = uniMsg
  End If
End Sub

Private Sub txt_ipAddr_DblClick(Button As Integer)

  unity_main.formfrom = 20
  unity_main.varfrom = 1
  frm_kybd.lbl_kybd.Caption = lbl_ipAddr.Caption
  frm_kybd.txt_kybd.Text = txt_ipAddr.Text
  frm_kybd.Show 1
End Sub





