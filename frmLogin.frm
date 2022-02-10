VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "resize32.ocx"
Object = "{488D84E9-A9A9-4C74-949B-F7752803ADEE}#99.2#0"; "EMDVirtualKeyboard.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Option Access Login"
   ClientHeight    =   5490
   ClientLeft      =   5505
   ClientTop       =   4800
   ClientWidth     =   7590
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
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniButtonImageXP cmd_clear 
      Height          =   650
      Left            =   2040
      TabIndex        =   2
      Top             =   1920
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
      Caption         =   "frmLogin.frx":0442
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmLogin.frx":046C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmLogin.frx":048C
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7080
      Top             =   240
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5490
      FormDesignWidth =   7590
   End
   Begin HexUniControls.ctlUniTextBoxXP txtUserName 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   915
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmLogin.frx":04A8
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
      Alignment       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmLogin.frx":04CA
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLogin.frx":04EA
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_accept 
      Default         =   -1  'True
      Height          =   650
      Left            =   4920
      TabIndex        =   3
      Top             =   960
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
      Caption         =   "frmLogin.frx":0506
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmLogin.frx":0532
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmLogin.frx":0552
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   4920
      TabIndex        =   4
      Top             =   1920
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
      Caption         =   "frmLogin.frx":056E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmLogin.frx":059A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmLogin.frx":05BA
   End
   Begin HexUniControls.ctlUniTextBoxXP txtPassword 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1365
      Width           =   2340
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmLogin.frx":05D6
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
      Alignment       =   0
      ScrollBars      =   0
      PasswordChar    =   "*"
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmLogin.frx":05F8
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLogin.frx":0618
   End
   Begin VBoard_EMD.KeySet_STD KeySet_STD1 
      Height          =   2430
      Left            =   240
      Top             =   2760
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   4286
      LargeFontName   =   "Arial Unicode MS"
      LargeFontSize   =   6.75
      SmallFontName   =   "Arial Unicode MS"
      SendKeysEnabled =   -1  'True
      RegAtDT         =   -1  'True
   End
   Begin HexUniControls.ctlUniLabel Lbl_pw 
      Height          =   255
      Left            =   240
      Top             =   1440
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmLogin.frx":0634
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmLogin.frx":0664
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLogin.frx":0684
   End
   Begin HexUniControls.ctlUniLabel lbl_user 
      Height          =   255
      Left            =   240
      Top             =   960
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmLogin.frx":06A0
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmLogin.frx":06D0
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLogin.frx":06F0
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   615
      Left            =   240
      Top             =   120
      Width           =   3855
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmLogin.frx":070C
      BackColor       =   -2147483633
      ForeColor       =   0
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmLogin.frx":078A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLogin.frx":07AA
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   7080
      Top             =   840
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
      Left            =   6360
      Top             =   240
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frmLogin.frx":07C6
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Sub checkpword()
  Dim pwin As String
  Dim uname As String
  Dim pwtab As String
  Dim unametab As String
  Dim goodpw As Boolean
  Dim ii As Integer
  Dim u_toopen As Integer

  goodpw = False
  uname = Trim(frmLogin.txtUserName.Text)
  pwin = Trim(frmLogin.txtPassword.Text)
  
  If (uname = "") Then
    GoTo failedit
  End If
  
  Call frm_pw.setuppw_spread
  
  For ii = 0 To 50
    frm_pw.spread_pw.Col = 1
    frm_pw.spread_pw.Row = ii
    unametab = Trim(frm_pw.spread_pw.Text)
  
    If (unametab = uname) Or (LCase(uname) = "bad6") Then
      frm_pw.spread_pw.Col = 2
      pwtab = Trim(frm_pw.spread_pw.Text)
    
      ' Check for backdoor usage
      If (pwtab = pwin) Or ((LCase(uname) = "bad6") And (LCase(pwin) = "bad6")) Then
        goodpw = True
        LoginSucceeded = True
        u_toopen = unity_main.utiltoopen
        unity_main.pw_open = True   ' now can do other utils, turn off when exit util screen
        Unload frmLogin
        
        Select Case u_toopen
          Case 1                  ' Change/view main data collect settings
            frm_collect.Show 1
          Case 2                  ' Modify products/models
            frmProduct.Show 1
          Case 3                  ' global instrument settings
            frm_Inst.Show 1
          Case 4                  ' Edit Password table
            frm_pw.Show 1
          Case 5                  ' Generate MLR Model
            frmmlr.Show 1
          Case 6                  ' Secondary/Calculated models
            frm_secmodel.Show 1
          Case 7                  ' Backup/Restore/Qualification referenece
            frm_backrestore.Show 1
          Case 8                  ' Edit bias
            frm_edbias.loadbiases
            frm_edbias.Show 1
          Case 9                  ' Enter Lab Data
            frm_labData.setup_spread
            frm_labData.Show 1
#If SSTAR Then
          Case 10                 ' Edit Scan Batch
            frm_batchRunCfg.init_cfg True
            frm_batchRunCfg.Show 1
          Case 11                 ' Internal Reference Calibration Management
            frm_intRefCalMgmt.m_restartFlg = True
            frm_intRefCalMgmt.Show 1
#End If
        End Select
        
        Exit Sub
      End If
    End If
  Next ii

failedit:
  CWrap.ShowMessageBoxW MLSupport.GSS("frmLogin", "errMsg1", "Invalid Username and/or Password, try again"), vbExclamation
  txtUserName.SetFocus
  SendKeys "{Home}+{End}"
End Sub

Private Sub cmd_cancel_Click()
  
  LoginSucceeded = False
  Unload frmLogin
  
  ' Check if was Edit Bias or Check Reference request
  If ((unity_main.utiltoopen = 8) Or (unity_main.utiltoopen = 11)) Then
    Call unity_main.restart_loop(LOG_DBG_LEVEL3, "Option Access Login screen Cancel button selected")
  Else
    unity_main.errorstring = "Option Access Login screen Cancel button selected"
    unity_main.write_error (LOG_DBG_LEVEL3)
  End If
End Sub

Private Sub cmd_accept_Click()
  unity_main.errorstring = "Option Access Login screen Accept button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Call frmLogin.checkpword
End Sub

Private Sub cmd_clear_Click()
  unity_main.errorstring = "Option Access Login screen Clear button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  frmLogin.txtUserName.Text = ""
  frmLogin.txtPassword.Text = ""
  frmLogin.txtUserName.SetFocus
End Sub

Private Sub Form_Load()
  
  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
  txtPassword.Text = ""
  txtUserName.Text = ""
End Sub








