VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{488D84E9-A9A9-4C74-949B-F7752803ADEE}#99.2#0"; "EMDVirtualKeyBoard.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_pw 
   Caption         =   "Options Access Passwords"
   ClientHeight    =   7245
   ClientLeft      =   180
   ClientTop       =   525
   ClientWidth     =   9870
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
   Icon            =   "frm_pw.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7560
      Top             =   5160
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7245
      FormDesignWidth =   9870
   End
   Begin FPUSpreadADO.fpSpread spread_pw 
      Height          =   4260
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5025
      _Version        =   458752
      _ExtentX        =   8864
      _ExtentY        =   7514
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      ColHeaderDisplay=   0
      EditModePermanent=   -1  'True
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      MaxRows         =   50
      ProcessTab      =   -1  'True
      SelectBlockOptions=   0
      SpreadDesigner  =   "frm_pw.frx":0442
      UserResize      =   1
      Appearance      =   2
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   7680
      TabIndex        =   4
      Top             =   3720
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
      Caption         =   "frm_pw.frx":06B1
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_pw.frx":06DD
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_pw.frx":06FD
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_deluser 
      Height          =   360
      Left            =   6615
      TabIndex        =   2
      Top             =   3240
      Width           =   1935
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_pw.frx":0719
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
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
      Tip             =   "frm_pw.frx":0739
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_pw.frx":0759
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_delete 
      Height          =   650
      Left            =   6600
      TabIndex        =   1
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
      Caption         =   "frm_pw.frx":0775
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_pw.frx":07AB
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_pw.frx":07CB
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   650
      Left            =   5400
      TabIndex        =   3
      Top             =   3720
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
      Caption         =   "frm_pw.frx":07E7
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_pw.frx":081F
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_pw.frx":083F
   End
   Begin VBoard_EMD.KeySet_STD KeySet_STD1 
      Height          =   2055
      Left            =   120
      Top             =   4560
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
   Begin HexUniControls.ctlUniLabel Label2 
      Height          =   1575
      Left            =   5400
      Top             =   240
      Width           =   4320
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_pw.frx":085B
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_pw.frx":099D
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_pw.frx":09BD
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   495
      Left            =   6615
      Top             =   2640
      Width           =   1815
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
      Caption         =   "frm_pw.frx":09D9
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_pw.frx":0A2B
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_pw.frx":0A4B
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   9120
      Top             =   5160
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
      Left            =   8400
      Top             =   5160
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_pw.frx":0A67
   End
End
Attribute VB_Name = "frm_pw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_save_Click()
  Dim ret As Boolean
  
  unity_main.errorstring = "Options Access Passwords screen Save Changes button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  ret = frm_pw.spread_pw.SaveToFile((CFG_DIR & LOGIN_PW_FILE), False)
  unity_main.errorstring = ("User saved new settings for configuration file: " & (CFG_DIR & LOGIN_PW_FILE))
  unity_main.write_error (LOG_DBG_LEVEL1)
  
  frm_pw.txt_deluser.Text = ""
  Unload frm_pw
End Sub

Private Sub cmd_delete_Click()
  Dim ii As Integer
  
  unity_main.errorstring = "Options Access Passwords screen Delete User button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  frm_pw.spread_pw.Col = 1
  
  For ii = 1 To 50
    frm_pw.spread_pw.Row = ii
    
    If Trim(frm_pw.spread_pw.Text) = Trim(frm_pw.txt_deluser.Text) Then
      frm_pw.spread_pw.Action = ActionDeleteRow
    End If
  Next ii
End Sub

Private Sub cmd_cancel_Click()
  
  unity_main.errorstring = "Options Access Passwords screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  frm_pw.txt_deluser.Text = ""
  Unload frm_pw
End Sub

Private Sub Form_Load()
  
  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
  Call frm_pw.setuppw_spread
End Sub

Sub setuppw_spread()
  Dim ret As Boolean
  Dim endrow As Integer
  Dim ii As Integer
  
  ret = frm_pw.spread_pw.LoadFromFile(CFG_DIR & LOGIN_PW_FILE)
  
  frm_pw.spread_pw.Row = 0
  frm_pw.spread_pw.Col = 1
  frm_pw.spread_pw.ColWidth(1) = 20
  frm_pw.spread_pw.Text = MLSupport.GSS("frm_pw", "header1", "Username")
  frm_pw.spread_pw.Font.Bold = True
  frm_pw.spread_pw.Col = 2
  frm_pw.spread_pw.ColWidth(2) = 20
  frm_pw.spread_pw.Text = MLSupport.GSS("frm_pw", "header2", "Password")
  frm_pw.spread_pw.Font.Bold = True

  For ii = 1 To 50
    frm_pw.spread_pw.Row = ii
    frm_pw.spread_pw.Col = 1
    
    If (Trim(frm_pw.spread_pw.Text) = "") Then
      frm_pw.spread_pw.Col = 2
      
      If (Trim(frm_pw.spread_pw.Text) = "") Then
        endrow = ii
      End If
    End If
  Next ii
  
  frm_pw.spread_pw.MaxRows = endrow
  
  frm_pw.spread_pw.Row = 1
  frm_pw.spread_pw.Col = 0
  frm_pw.spread_pw.Row2 = frm_pw.spread_pw.MaxRows
  frm_pw.spread_pw.Col2 = frm_pw.spread_pw.MaxCols
  frm_pw.spread_pw.BlockMode = True
  frm_pw.spread_pw.Font.Bold = False
  frm_pw.spread_pw.FontSize = 12
  frm_pw.spread_pw.Font.Name = "Arial Unicode MS"
  frm_pw.spread_pw.BlockMode = False
End Sub








