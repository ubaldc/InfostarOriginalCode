VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_reset 
   Caption         =   "SpectraStar in Reset Mode"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
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
   ScaleHeight     =   3525
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   1680
      Top             =   2880
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3525
      FormDesignWidth =   9780
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_exit 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   3890
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
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
      Caption         =   "frm_reset.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_reset.frx":0028
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_reset.frx":0048
   End
   Begin VB.Timer tmr_reset 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   360
      Top             =   2880
   End
   Begin HexUniControls.ctlUniLabel lblResetStatus 
      Height          =   1095
      Left            =   240
      Top             =   1200
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1931
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_reset.frx":0064
      BackColor       =   -2147483633
      ForeColor       =   32768
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_reset.frx":0084
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_reset.frx":00A4
   End
   Begin HexUniControls.ctlUniLabel lbl_errMsg 
      Height          =   615
      Left            =   240
      Top             =   240
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_reset.frx":00C0
      BackColor       =   -2147483633
      ForeColor       =   255
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_reset.frx":00E0
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_reset.frx":0100
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   2280
      Top             =   2880
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
      Left            =   1080
      Top             =   2760
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_reset.frx":011C
   End
End
Attribute VB_Name = "frm_reset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_errMsgFlg As Boolean
Private m_tmrCtr As Integer

Private Sub cmd_exit_Click()
  
  unity_main.errorstring = "SpectraStar in Reset Mode screen Exit button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  tmr_reset.enabled = False
  Unload frm_reset
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
  tmr_reset.enabled = True
  m_tmrCtr = 0
End Sub

Private Sub tmr_reset_Timer()
  Dim errBits As MASTER_ERRORS
  Dim err As MASTER_ERRORS
  Dim errCodeMsg As String
  Dim errMsg As String
  Dim uniMsg As String
  
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.GetMasterErrors(MS11MasterErrors)
#Else
  MS11MasterErrors = unity_main.MS11srv.MasterErrors
#End If
  errBits = ME_CODE_DEAD + ME_MSCOM_DEAD + ME_MS11_INIT + ME_SCNR_FLD + ME_TRAY_FLD
  err = MS11MasterErrors And errBits
    
  ' Check if any major error bits set
  If (err <> ME_OKAY) Then
#If SSRCS Then
    SSRCSClientError = unity_main.SSRCSClient.SetMasterErrors(MS11MasterErrors)
#Else
    unity_main.MS11srv.MasterErrors = MS11MasterErrors
#End If
    
    ' Check if BIT(0) set, Operational Errors within Code prevent Running
    If ((err And ME_CODE_DEAD) <> 0) Then
      errCodeMsg = "MErr = 0x" & Hex(MS11MasterErrors)
      errMsg = "SpectraStar has major operational error (" & errCodeMsg & "). InfoStar will shutdown automatically"
      unity_main.errorstring = errMsg
      unity_main.write_error
      uniMsg = MLSupport.GGS_Params("MS11srv.errMsg1", "SpectraStar has major operational error (%1)", errCodeMsg)
      uniMsg = MLSupport.GGS_Params("errMsg2", "%1. InfoStar will shutdown automatically", uniMsg)
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
      End
    End If
    
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
      lblResetStatus.Caption = uniMsg
    Else
      m_errMsgFlg = False
      lblResetStatus.Caption = ""
    End If
    
    m_tmrCtr = m_tmrCtr + 1
    
    ' If error not clear in 30 seconds, show exit button
    If (m_tmrCtr > 60) Then
      cmd_exit.Visible = True
    End If
  Else
    tmr_reset.enabled = False
    Unload frm_reset
  End If
End Sub






