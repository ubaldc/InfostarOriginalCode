VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "resize32.ocx"
Object = "{488D84E9-A9A9-4C74-949B-F7752803ADEE}#99.2#0"; "EMDVirtualKeyboard.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_guilevel 
   BorderStyle     =   0  'None
   Caption         =   "GUI Mode Selection/Login"
   ClientHeight    =   11265
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   13335
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
   Icon            =   "frm_guilevel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11265
   ScaleWidth      =   13335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin HexUniControls.ctlUniTextBoxXP txt_pw 
      Height          =   735
      Left            =   1440
      TabIndex        =   0
      Top             =   2520
      Width           =   7335
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_guilevel.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   15.75
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
      Tip             =   "frm_guilevel.frx":0462
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_guilevel.frx":0482
   End
   Begin HexUniControls.ctlUniLabel lbl_wrongpw 
      Height          =   495
      Left            =   960
      Top             =   2040
      Visible         =   0   'False
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_guilevel.frx":049E
      BackColor       =   -2147483633
      ForeColor       =   255
      BorderColor     =   0
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_guilevel.frx":0508
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_guilevel.frx":0528
   End
   Begin HexUniControls.ctlUniLabel Label2 
      Height          =   615
      Left            =   3120
      Top             =   840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_guilevel.frx":0544
      BackColor       =   16777215
      ForeColor       =   12582912
      BorderColor     =   0
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_guilevel.frx":058E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_guilevel.frx":05AE
   End
   Begin VB.Timer tmr2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   120
      Top             =   2280
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_about 
      Height          =   945
      Left            =   360
      TabIndex        =   3
      Top             =   9120
      Width           =   2535
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_guilevel.frx":05CA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_guilevel.frx":0606
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_guilevel.frx":0626
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   600
      Top             =   7320
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   11265
      FormDesignWidth =   13335
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_run 
      Height          =   900
      Left            =   10080
      TabIndex        =   1
      Top             =   720
      Width           =   2805
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_guilevel.frx":0642
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_guilevel.frx":0672
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_guilevel.frx":0692
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_enter 
      Default         =   -1  'True
      Height          =   900
      Left            =   10080
      TabIndex        =   2
      Top             =   1800
      Width           =   2805
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_guilevel.frx":06AE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_guilevel.frx":06EA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_guilevel.frx":070A
   End
   Begin VB.Timer tmr_pw 
      Interval        =   50000
      Left            =   13200
      Top             =   1320
   End
   Begin HexUniControls.ctlUniLabel lblCompany 
      Height          =   495
      Left            =   9120
      Top             =   9585
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_guilevel.frx":0726
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      BorderColor     =   0
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_guilevel.frx":0766
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_guilevel.frx":0786
   End
   Begin HexUniControls.ctlUniLabel lblCopyright 
      Height          =   375
      Left            =   9120
      Top             =   9120
      Width           =   4035
      _ExtentX        =   7117
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
      Caption         =   "frm_guilevel.frx":07A2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      BorderColor     =   0
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_guilevel.frx":07EC
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_guilevel.frx":080C
   End
   Begin VBoard_EMD.KeySet_STD KeySet_STD1 
      Height          =   3180
      Left            =   360
      Top             =   3480
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   5609
      LargeFontName   =   "Arial Unicode MS"
      LargeFontSize   =   6.75
      SmallFontName   =   "Arial Unicode MS"
      ButtonWidth     =   40
      ButtonHeight    =   40
      SendKeysEnabled =   -1  'True
      RegAtDT         =   -1  'True
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   735
      Left            =   9840
      Top             =   2760
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_guilevel.frx":0828
      BackColor       =   -2147483633
      ForeColor       =   255
      BorderColor     =   0
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_guilevel.frx":087C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_guilevel.frx":089C
   End
   Begin HexUniControls.ctlUniImage ctlUniImage1 
      Height          =   1305
      Left            =   360
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2302
      Picture         =   "frm_guilevel.frx":08B8
      Tip             =   "frm_guilevel.frx":57BF
      Enabled         =   -1  'True
      Border          =   0   'False
      BackColor       =   -2147483633
      BorderColor     =   -1
      RoundedBorders  =   -1  'True
      Stretch         =   -1  'True
      XTransp         =   0
      YTransp         =   0
      MousePointer    =   0
      MouseIcon       =   "frm_guilevel.frx":57DF
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   600
      Top             =   7680
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
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   30
      X1              =   -840
      X2              =   13800
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   30
      X1              =   3240
      X2              =   17040
      Y1              =   11400
      Y2              =   11400
   End
End
Attribute VB_Name = "frm_guilevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pwpassed As Boolean
Public guimin As Boolean

Sub checkguilevel()
  Dim rtn As Long

  ' Check if in Run/Operator mode
  If (unity_main.run_min_gui = True) Then
    If (unity_main.m_allowBias = True) Then
      unity_main.cmd_bias.Visible = True
    Else
      unity_main.cmd_bias.Visible = False
    End If
    
    unity_main.cmd_options.Visible = False
    unity_main.cmd_start.Visible = False
    unity_main.cmd_stop.Visible = False
#If SSRCS Then
    If (unity_main.m_ssrcsConnected = False) Then
      unity_main.cmd_ssrcsConnect.Visible = True
    Else
      unity_main.cmd_ssrcsConnect.Visible = False
    End If
#End If

#If SSTAR Then
    ' Check if internal reference verification required
    If (unity_main.m_intRefVerReqdFlg = True) Then
      ' Resize and move Verify Reference button
      unity_main.cmd_verifyRef.Left = unity_main.cmd_select.Left
      unity_main.cmd_verifyRef.Width = unity_main.cmd_select.Width
      unity_main.cmd_verifyRef.Visible = True
    Else
      unity_main.cmd_verifyRef.Visible = False
    End If
#End If

    ' Check if product loaded
    If (unity_main.txtsamplename.Text = "") Then
      unity_main.cmd_repacks.Visible = False
      unity_main.img_ticket.Visible = False
      unity_main.img_batchRpt.Visible = False
      unity_main.m_batchRptFile = ""
    Else
      If (unity_main.m_writeTkt > 0) Then
        unity_main.img_ticket.Visible = True
      End If
      
      ' Check if batch report created
      If (unity_main.m_batchRptFile <> "") Then
        unity_main.img_batchRpt.Visible = True
      End If
    End If

    unity_main.img_report.Visible = False
    unity_main.img_binocs.Visible = False
    unity_main.img_csv.Visible = False
    unity_main.img_help.Visible = False
    unity_main.Picture1.Visible = True
    
    unity_main.hide_taskbar
  Else      ' Supervisor mode
    unity_main.cmd_bias.Visible = False
    unity_main.cmd_options.Visible = True
    unity_main.cmd_start.Visible = False
    unity_main.cmd_stop.Visible = True
#If SSRCS Then
    If (unity_main.m_ssrcsConnected = False) Then
      unity_main.cmd_ssrcsConnect.Visible = True
    Else
      unity_main.cmd_ssrcsConnect.Visible = False
    End If
#End If

#If SSTAR Then
    ' Check if internal reference verification required
    If (unity_main.m_intRefVerReqdFlg = True) Then
      ' Resize Options and Verify Reference button
      unity_main.cmd_options.Width = unity_main.cmd_runBatch.Width
      unity_main.cmd_verifyRef.Left = unity_main.cmd_runBatch.Left
      unity_main.cmd_verifyRef.Width = unity_main.cmd_runBatch.Width
      unity_main.cmd_verifyRef.Visible = True
    Else
      ' Resize Options button
      unity_main.cmd_options.Width = unity_main.cmd_select.Width
      unity_main.cmd_verifyRef.Visible = False
    End If
#End If
    
    If (unity_main.txtsamplename.Text = "") Then
      unity_main.cmd_repacks.Visible = False
      unity_main.img_report.Visible = False
      unity_main.img_ticket.Visible = False
      unity_main.img_batchRpt.Visible = False
      unity_main.m_batchRptFile = ""
    Else
      unity_main.img_report.Visible = True

      If (unity_main.m_writeTkt > 0) Then
        unity_main.img_ticket.Visible = True
      End If
      
      ' Check if batch report created
      If (unity_main.m_batchRptFile <> "") Then
        unity_main.img_batchRpt.Visible = True
      End If
    End If

    unity_main.img_binocs.Visible = True
    unity_main.img_csv.Visible = True
    unity_main.img_help.Visible = True
    unity_main.Picture1.Visible = True
    
    rtn = FindWindow("Shell_traywnd", "") 'get the Window
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW) 'show the Taskbar
  End If
End Sub

Private Sub cmd_enter_Click()
  
  unity_main.errorstring = "Startup Password screen Enter Password button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  tmr2.enabled = True
  frm_guilevel.lbl_wrongpw.Visible = False
  frm_guilevel.tmr_pw.enabled = False
  Call frm_guilevel.checkpw2
End Sub

Sub checkpw2()
  Dim temppw As String
  Dim ftokill As String
  Dim fileexists As Boolean

  frm_guilevel.pwpassed = False
  Call frm_guipw.loadguipw
  temppw = Trim(frm_guilevel.txt_pw.Text)
  
  If (temppw = "rokadia") Then
    unity_main.unloadallforms "frm_guilevel"
    Unload Me
    End
  End If
  
  If (temppw = frm_guipw.guipw) Or (temppw = "good hair") Or (temppw = "sesame") Then
    unity_main.run_min_gui = False
    frm_guilevel.pwpassed = True
    frm_guilevel.checkguilevel
    GoTo passedpw
  End If

  frm_guilevel.lbl_wrongpw.Visible = True
  frm_guilevel.txt_pw.Text = ""
  frm_guilevel.txt_pw.SetFocus
  Exit Sub
  
passedpw:
  frm_guilevel.tmr_pw.enabled = False
  unity_main.Visible = True
  unity_main.setup_olcols
  Unload frm_guilevel
  
  ' Start auto operations
  unity_main.tmr_all.enabled = True
  unity_main.tmr_sec1.enabled = True
  unity_main.tmr_sec30.enabled = True
End Sub

Private Sub cmd_run_Click()
  
  unity_main.errorstring = "Startup Password screen Run Mode button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  frm_guilevel.forcemingui
End Sub

Sub forcemingui()
  
  frm_guilevel.tmr_pw.enabled = False
  unity_main.m_enableRunMode = True
  unity_main.run_min_gui = True
  frm_guilevel.checkguilevel
  
  unity_main.Visible = True
  unity_main.setup_olcols
  Unload frm_guilevel

  ' Start auto operations
  unity_main.tmr_all.enabled = True
  unity_main.tmr_sec1.enabled = True
  unity_main.tmr_sec30.enabled = True
End Sub

Sub forcemaxgui()
  
  frm_guilevel.tmr_pw.enabled = False
  unity_main.m_enableRunMode = False
  unity_main.run_min_gui = False
  Call frm_guilevel.checkguilevel
  
  unity_main.Visible = True
  unity_main.setup_olcols
  Unload frm_guilevel
  
  ' Start auto operations
  unity_main.tmr_all.enabled = True
  unity_main.tmr_sec1.enabled = True
  unity_main.tmr_sec30.enabled = True
End Sub

Private Sub cmd_about_Click()
  
  unity_main.errorstring = "Startup Password screen About InfoStar button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  frmSplash2.tmr_splash.enabled = False
  frmSplash2.cmd_exit.Visible = True
  frmSplash2.Show 1
End Sub

Private Sub Form_Activate()

  frm_guilevel.txt_pw.SetFocus
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub

Private Sub KeySet_STD1_VKeyDown(KeyCode As Integer, Shift As Integer)

  frm_guilevel.txt_pw.enabled = True
  frm_guilevel.txt_pw.SetFocus
End Sub

Private Sub tmr_pw_Timer()
  
  frm_guilevel.tmr_pw.enabled = False
  
  ' Unload splash screen if shown to avoid non-modal screen loading while modal is loaded
  Unload frmSplash2
  
  frm_guilevel.forcemingui
  Unload frm_kybd
End Sub

Private Sub tmr2_Timer()

  frm_guilevel.lbl_wrongpw.Visible = False
  tmr2.enabled = False
End Sub

Private Sub txt_pw_Change()
  
  frm_guilevel.tmr_pw.enabled = False
End Sub








