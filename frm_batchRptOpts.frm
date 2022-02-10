VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_batchRptOpts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Batch Report File Options"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7230
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
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniButtonImageXP cmd_overWrite 
      Height          =   705
      Left            =   2520
      TabIndex        =   2
      Top             =   2040
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
      Caption         =   "frm_batchRptOpts.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_batchRptOpts.frx":0032
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_batchRptOpts.frx":0052
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   1695
      Left            =   240
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   2990
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_batchRptOpts.frx":006E
      BackColor       =   -2147483633
      ForeColor       =   192
      Alignment       =   2
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_batchRptOpts.frx":019C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_batchRptOpts.frx":01BC
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   2640
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3195
      FormDesignWidth =   7230
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   705
      Left            =   4920
      TabIndex        =   0
      Top             =   2040
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
      Caption         =   "frm_batchRptOpts.frx":01D8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_batchRptOpts.frx":0204
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_batchRptOpts.frx":0224
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_append 
      Height          =   705
      Left            =   120
      TabIndex        =   1
      Top             =   2040
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
      Caption         =   "frm_batchRptOpts.frx":0240
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_batchRptOpts.frx":026C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_batchRptOpts.frx":028C
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   1080
      Top             =   2640
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
      Left            =   600
      Top             =   2640
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_batchRptOpts.frx":02A8
   End
End
Attribute VB_Name = "frm_batchRptOpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmd_append_Click()
  
  unity_main.errorstring = "Batch Report File Options screen Append button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  frm_batchRunCfg.m_batchRptAppendFlg = True
  frm_batchRunCfg.m_batchRptCancelFlg = False
  Unload frm_batchRptOpts
End Sub

Private Sub cmd_cancel_Click()

  unity_main.errorstring = "Batch Report File Options screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  frm_batchRunCfg.m_batchRptAppendFlg = False
  frm_batchRunCfg.m_batchRptCancelFlg = True
  Unload frm_batchRptOpts
End Sub

Private Sub cmd_overWrite_Click()

  unity_main.errorstring = "Batch Report File Options screen Overwrite button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  frm_batchRunCfg.m_batchRptAppendFlg = False
  frm_batchRunCfg.m_batchRptCancelFlg = False
  Unload frm_batchRptOpts
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub






