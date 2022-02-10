VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Object = "{488D84E9-A9A9-4C74-949B-F7752803ADEE}#99.2#0"; "EMDVirtualKeyBoard.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frm_guipw 
   Caption         =   "Supervisor Login Password"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8535
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
   Icon            =   "frm_guipw.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin FPUSpreadADO.fpSpread ss_pwmingui 
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   8055
      _Version        =   458752
      _ExtentX        =   14208
      _ExtentY        =   1085
      _StockProps     =   64
      ColHeaderDisplay=   0
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   1
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      ScrollBars      =   0
      SpreadDesigner  =   "frm_guipw.frx":0442
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_clear 
      Height          =   650
      Left            =   1920
      TabIndex        =   4
      Top             =   1560
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
      Caption         =   "frm_guipw.frx":06BB
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_guipw.frx":06E5
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_guipw.frx":0705
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7680
      Top             =   3600
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5550
      FormDesignWidth =   8535
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   6240
      TabIndex        =   3
      Top             =   1560
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
      Caption         =   "frm_guipw.frx":0721
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_guipw.frx":074D
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_guipw.frx":076D
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_guipw 
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5535
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_guipw.frx":0789
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   18
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
      Tip             =   "frm_guipw.frx":07A9
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_guipw.frx":07C9
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Default         =   -1  'True
      Height          =   650
      Left            =   6240
      TabIndex        =   2
      Top             =   480
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
      Caption         =   "frm_guipw.frx":07E5
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_guipw.frx":081D
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_guipw.frx":083D
   End
   Begin VBoard_EMD.KeySet_STD KeySet_STD1 
      Height          =   2430
      Left            =   240
      Top             =   2520
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   4286
      LargeFontName   =   "Arial Unicode MS"
      LargeFontSize   =   6.75
      SmallFontName   =   "Arial Unicode MS"
      SendKeysEnabled =   -1  'True
      RegAtDT         =   -1  'True
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   7680
      Top             =   4200
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
      Left            =   7680
      Top             =   3000
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_guipw.frx":0859
   End
End
Attribute VB_Name = "frm_guipw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public guipw As String

Private Sub cmd_save_Click()
  Dim lenpw As Integer
  Dim ret As Boolean

  unity_main.errorstring = "Supervisor Login Password screen Save Changes button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  lenpw = Len(frm_guipw.txt_guipw.Text)
  
  If (lenpw < 2) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_guipw", "errMsg1", "Your password must have at least 2 characters"), vbCritical
    Exit Sub
  End If

  frm_guipw.ss_pwmingui.Row = 1
  frm_guipw.ss_pwmingui.Col = 1
  frm_guipw.ss_pwmingui.Text = Trim(frm_guipw.txt_guipw.Text)

  ret = frm_guipw.ss_pwmingui.SaveToFile((CFG_DIR & GUI_PW_FILE), False)
  unity_main.errorstring = ("User saved new settings for configuration file: " & (CFG_DIR & GUI_PW_FILE))
  unity_main.write_error (LOG_DBG_LEVEL1)
  
  frm_guipw.guipw = Trim(frm_guipw.ss_pwmingui.Text)
  Unload frm_guipw
End Sub

Private Sub cmd_cancel_Click()
  
  unity_main.errorstring = "Supervisor Login Password screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Unload frm_guipw
End Sub

Sub loadguipw()
  Dim ret As Boolean
  Dim tempstring As String
  Dim uniMsg As String

  On Error Resume Next
  ret = frm_guipw.ss_pwmingui.LoadFromFile(CFG_DIR & GUI_PW_FILE)

  If (ret = False) Then
    tempstring = (CFG_DIR & GUI_PW_FILE & " file open/read error. " & Error$)
    unity_main.errorstring = tempstring
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg1", "%1 file open/read error. %2", CFG_DIR & GUI_PW_FILE, Error$)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    Exit Sub
  End If

  frm_guipw.ss_pwmingui.Row = 1
  frm_guipw.ss_pwmingui.Col = 1
  tempstring = frm_guipw.ss_pwmingui.Text
  frm_guipw.txt_guipw.Text = tempstring
  frm_guipw.guipw = Trim(tempstring)
End Sub

Private Sub cmd_clear_Click()
  
  unity_main.errorstring = "Supervisor Login Password screen Clear button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  txt_guipw.Text = ""
  txt_guipw.SetFocus
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub








