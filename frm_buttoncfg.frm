VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Object = "{488D84E9-A9A9-4C74-949B-F7752803ADEE}#99.2#0"; "EMDVirtualKeyBoard.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frm_buttoncfg 
   Caption         =   "User Inputs Configuration"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12555
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
   ScaleHeight     =   8400
   ScaleWidth      =   12555
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   9480
      TabIndex        =   3
      Top             =   7080
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
      Caption         =   "frm_buttoncfg.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_buttoncfg.frx":002C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_buttoncfg.frx":004C
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   6360
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8400
      FormDesignWidth =   12555
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   650
      Left            =   9480
      TabIndex        =   2
      Top             =   6360
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
      Caption         =   "frm_buttoncfg.frx":0068
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_buttoncfg.frx":00A0
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_buttoncfg.frx":00C0
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_view 
      Height          =   650
      Left            =   9480
      TabIndex        =   1
      Top             =   5640
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
      Caption         =   "frm_buttoncfg.frx":00DC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_buttoncfg.frx":0104
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_buttoncfg.frx":0124
   End
   Begin FPUSpreadADO.fpSpread ss_buttonconfig 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12135
      _Version        =   458752
      _ExtentX        =   21405
      _ExtentY        =   9340
      _StockProps     =   64
      EditModePermanent=   -1  'True
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      MaxRows         =   103
      ProcessTab      =   -1  'True
      SelectBlockOptions=   0
      SpreadDesigner  =   "frm_buttoncfg.frx":0140
      UserResize      =   1
   End
   Begin VBoard_EMD.KeySet_STD KeySet_STD1 
      Height          =   2055
      Left            =   1200
      Top             =   5640
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   3625
      LargeFontName   =   "Arial Unicode MS"
      LargeFontSize   =   6.75
      SmallFontName   =   "Arial Unicode MS"
      ButtonWidth     =   25
      ButtonHeight    =   25
      SendKeysEnabled =   -1  'True
      RegAtDT         =   -1  'True
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   0
      Top             =   6960
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
      Left            =   0
      Top             =   5880
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_buttoncfg.frx":1045
   End
End
Attribute VB_Name = "frm_buttoncfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub loadit(mustBeCfg As Boolean)
  Dim ret As Boolean
  Dim buttonFile As String
  Dim uniMsg As String

  buttonFile = (CFG_DIR & USER_INPUTS_CFG_FILE)
  
  If (CFile.st_FileExist(buttonFile) = False) Then
    ' Check if file should be configured
    If (mustBeCfg = True) Then
      uniMsg = MLSupport.GGS_Params("fileErrMsg6", "%1 file not found", buttonFile)
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_buttoncfg.errMsg1", "%1. Loaded product configured for user inputs; please configure the User Inputs settings and save them to file.", uniMsg), vbExclamation
    End If
    
    Exit Sub
  End If

  On Error GoTo badload
  ret = frm_buttoncfg.ss_buttonconfig.LoadFromFile(buttonFile)
  
  frm_buttoncfg.ss_buttonconfig.ColWidth(0) = 25
  frm_buttoncfg.ss_buttonconfig.MaxRows = 103
  
  frm_buttoncfg.ss_buttonconfig.Row = 0
  frm_buttoncfg.ss_buttonconfig.Col = 0
  frm_buttoncfg.ss_buttonconfig.Row2 = frm_buttoncfg.ss_buttonconfig.MaxRows
  frm_buttoncfg.ss_buttonconfig.Col2 = frm_buttoncfg.ss_buttonconfig.MaxCols
  frm_buttoncfg.ss_buttonconfig.BlockMode = True
  frm_buttoncfg.ss_buttonconfig.Font.Name = "Arial Unicode MS"
  frm_buttoncfg.ss_buttonconfig.Font.Size = 10
  frm_buttoncfg.ss_buttonconfig.BlockMode = False
  Exit Sub
  
badload:
  uniMsg = MLSupport.GGS_Params("fileErrMsg4", "%1 file read error. %2", buttonFile, Error$)
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub  'loadit

Sub loadbuttonconfig(showLastValues As Boolean)
  Dim ii As Integer
  Dim nn As Integer
  Dim tempstring As String
  Dim uniMsg As String
  
  On Error GoTo badload
  clear_form

  For ii = 1 To MAX_MAN_INPUTS
    frm_buttoncfg.ss_buttonconfig.Col = ii
    frm_buttoncfg.ss_buttonconfig.Row = 1
  
    ' Check if input enabled
    If (frm_buttoncfg.ss_buttonconfig.Value = 1) Then
      frm_scanname.lbl(ii).Visible = True
      frm_buttoncfg.ss_buttonconfig.Row = 3
      tempstring = frm_buttoncfg.ss_buttonconfig.Text
      frm_scanname.lbl(ii).Caption = Trim(tempstring)
      frm_buttoncfg.ss_buttonconfig.Row = 2
    
      ' Check if using combo-box
      If (frm_buttoncfg.ss_buttonconfig.Value = 1) Then
        frm_scanname.combo(ii).Visible = True
      Else
        frm_scanname.txtbx(ii).Visible = True
      End If
    Else
      frm_scanname.lbl(ii).Visible = False
    End If
  Next ii
  
  Call filllists(showLastValues)
  Exit Sub

badload:
  uniMsg = MLSupport.GGS_Params("fileErrMsg4", "%1 file read error. %2", CFG_DIR & USER_INPUTS_CFG_FILE, Error$)
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub

Sub clear_form()
  Dim ii As Integer
  
  frm_scanname.txt_fname.Text = ""
  frm_scanname.txt_sampinfo.Text = ""
  
  For ii = 1 To MAX_MAN_INPUTS
    frm_scanname.lbl(ii).Caption = ""
    frm_scanname.combo(ii).Text = ""
    frm_scanname.combo(ii).Visible = False
    frm_scanname.txtbx(ii).Text = ""
    frm_scanname.txtbx(ii).Visible = False
  Next ii
End Sub

Sub filllists(showLastValues As Boolean)
  Dim ii, zz As Integer
  Dim tempstring As String
  Dim errMsg As String
  Dim uniMsg As String
  
  On Error GoTo BAD_FILL
  For ii = 1 To MAX_MAN_INPUTS
    frm_buttoncfg.ss_buttonconfig.Col = ii
    frm_buttoncfg.ss_buttonconfig.Row = 2
  
    ' Check if using combo-box
    If (frm_buttoncfg.ss_buttonconfig.Value = 1) Then
      frm_scanname.combo(ii).Clear

      For zz = 4 To 103
        frm_buttoncfg.ss_buttonconfig.Row = zz
        tempstring = Trim(frm_buttoncfg.ss_buttonconfig.Text)
      
        If (tempstring <> "") Then
          frm_scanname.combo(ii).AddItem (tempstring)
        End If
      Next zz
      
      ' Check if to show last entered user value
      If (showLastValues = True) Then
        frm_scanname.combo(ii).Text = UserInputs(ii)
      End If
    Else     ' text box
      ' Check if to show last entered user value
      If (showLastValues = True) Then
        frm_scanname.txtbx(ii).Text = UserInputs(ii)
      End If
    End If
  Next ii
  
  Exit Sub

BAD_FILL:
  errMsg = ("Problem building manual data entry screen. " & Error$)
  unity_main.errorstring = errMsg
  unity_main.write_error
  uniMsg = MLSupport.GGS_Params("frm_buttoncfg.errMsg2", "Problem building manual data entry screen. %1", Error$)
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub

Sub loadbuttonform()
  Dim zz As Integer
  
  For zz = 1 To MAX_MAN_INPUTS
    frm_buttoncfg.ss_buttonconfig.Col = zz
    frm_buttoncfg.ss_buttonconfig.ColWidth(zz) = 16
    frm_buttoncfg.ss_buttonconfig.Row = 0
    frm_buttoncfg.ss_buttonconfig.Text = MLSupport.GSS("frm_buttoncfg", "header" & zz, "Input " & zz)
    frm_buttoncfg.ss_buttonconfig.Row = 1
    frm_buttoncfg.ss_buttonconfig.TypeCheckText = MLSupport.GSS("frm_buttoncfg", "chkbox1", "Check to use")
    frm_buttoncfg.ss_buttonconfig.Row = 2
    frm_buttoncfg.ss_buttonconfig.TypeCheckText = MLSupport.GSS("frm_buttoncfg", "chkbox2", "Check for list")
  Next zz
  
  frm_buttoncfg.ss_buttonconfig.Col = 0

  frm_buttoncfg.ss_buttonconfig.Row = 1
  frm_buttoncfg.ss_buttonconfig.Text = MLSupport.GSS("frm_buttoncfg", "label1", "Enable Input")
  frm_buttoncfg.ss_buttonconfig.Row = 2
  frm_buttoncfg.ss_buttonconfig.Text = MLSupport.GSS("frm_buttoncfg", "label2", "Text Entry/List Selection")
  frm_buttoncfg.ss_buttonconfig.Row = 3
  frm_buttoncfg.ss_buttonconfig.Text = MLSupport.GSS("frm_buttoncfg", "label3", "Title")
  
  For zz = 4 To 103
    frm_buttoncfg.ss_buttonconfig.Row = zz
    frm_buttoncfg.ss_buttonconfig.Text = CStr(zz - 3)
  Next zz
End Sub  'loadbuttonform

Private Sub cmd_cancel_Click()
  
  unity_main.errorstring = "User Inputs Configuration screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  frm_buttoncfg.Visible = False
  
  ' Setup use input button/list selection
  If (unity_main.m_useMIV = True) Then
    frm_buttoncfg.loadit True
    frm_buttoncfg.loadbuttonform
    frm_buttoncfg.loadbuttonconfig True
  Else
    frm_collect.chk_userInputs.Value = 0
    frm_buttoncfg.clear_form
  End If
End Sub

Private Sub cmd_save_Click()
  Dim ret As Boolean
  Dim fileName As String
  Dim errMsg As String
  Dim uniMsg As String
  Dim ii As Integer
  
  unity_main.errorstring = "User Inputs Configuration screen Save Changes button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  On Error GoTo badload
  fileName = (CFG_DIR & USER_INPUTS_CFG_FILE)
  ret = frm_buttoncfg.ss_buttonconfig.SaveToFile(fileName, False)
  
  unity_main.errorstring = ("User saved new settings for configuration file: " & (CFG_DIR & USER_INPUTS_CFG_FILE))
  unity_main.write_error (LOG_DBG_LEVEL1)
  frm_buttoncfg.Visible = False
  
  ' Clear last saved user inputs values
  For ii = 1 To MAX_MAN_INPUTS
    UserInputs(ii) = ""
  Next ii
  
  ' Setup use input button/list selection
  If (unity_main.m_useMIV = True) Then
    frm_buttoncfg.loadit True
    frm_buttoncfg.loadbuttonform
    frm_buttoncfg.loadbuttonconfig True
  Else
    frm_collect.chk_userInputs.Value = 0
    frm_buttoncfg.clear_form
  End If

  Exit Sub

badload:
  errMsg = (fileName & " file write error. " & Error$)
  unity_main.errorstring = errMsg
  unity_main.write_error
  uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", fileName, Error$)
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub

Private Sub cmd_view_Click()
  
  unity_main.errorstring = "User Inputs Configuration screen View button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  frm_buttoncfg.loadbuttonconfig False
  frm_scanname.cmd_cancel.Visible = False
  frm_scanname.cmd_cancel.Cancel = False
  frm_scanname.cmd_exit.Visible = True
  frm_scanname.cmd_exit.Cancel = True
  frm_scanname.cmd_run.Visible = False
  
  frm_scanname.Show 1
  
  frm_scanname.cmd_exit.Visible = False
  frm_scanname.cmd_exit.Cancel = False
  frm_scanname.cmd_cancel.Visible = True
  frm_scanname.cmd_cancel.Cancel = True
  frm_scanname.cmd_run.Visible = True
End Sub

Private Sub Form_Load()
  
  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub








