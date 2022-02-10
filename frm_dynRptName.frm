VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Object = "{488D84E9-A9A9-4C74-949B-F7752803ADEE}#99.2#0"; "EMDVirtualKeyBoard.ocx"
Begin VB.Form frm_dynRptName 
   Caption         =   "Dynamic Report Name Entry"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8430
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
   Icon            =   "frm_dynRptName.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniButtonImageXP cmd_clear 
      Height          =   645
      Left            =   1920
      TabIndex        =   1
      Top             =   1920
      Width           =   1995
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
      Caption         =   "frm_dynRptName.frx":0442
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_dynRptName.frx":046C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_dynRptName.frx":048C
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_accept 
      Default         =   -1  'True
      Height          =   645
      Left            =   6240
      TabIndex        =   2
      Top             =   1080
      Width           =   1995
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
      Caption         =   "frm_dynRptName.frx":04A8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_dynRptName.frx":04D4
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_dynRptName.frx":04F4
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   645
      Left            =   6240
      TabIndex        =   3
      Top             =   1920
      Width           =   1995
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
      Caption         =   "frm_dynRptName.frx":0510
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_dynRptName.frx":053C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_dynRptName.frx":055C
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6360
      Top             =   2400
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5910
      FormDesignWidth =   8430
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_nameEntry 
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   4935
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_dynRptName.frx":0578
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   0
      MultiLine       =   -1  'True
      Alignment       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frm_dynRptName.frx":0598
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_dynRptName.frx":05B8
   End
   Begin VBoard_EMD.KeySet_STD KeySet_STD1 
      Height          =   2805
      Left            =   120
      Top             =   2880
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   4948
      LargeFontName   =   "Arial Unicode MS"
      LargeFontSize   =   6.75
      SmallFontName   =   "Arial Unicode MS"
      ButtonWidth     =   35
      ButtonHeight    =   35
      SendKeysEnabled =   -1  'True
      RegAtDT         =   -1  'True
   End
   Begin HexUniControls.ctlUniLabel lbl_nameEntry 
      Height          =   735
      Left            =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_dynRptName.frx":05D4
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   2
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_dynRptName.frx":0630
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_dynRptName.frx":0650
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   7920
      Top             =   2400
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
      Left            =   7200
      Top             =   2280
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_dynRptName.frx":066C
   End
End
Attribute VB_Name = "frm_dynRptName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_cancel_Click()

  unity_main.errorstring = "Dynamic Report Name Entry screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Unload frm_dynRptName
End Sub

Private Sub cmd_accept_Click()
  Dim jpgName As String
  Dim fileExt As String

  unity_main.errorstring = "Dynamic Report Name Entry screen Accept button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  If (Trim(txt_nameEntry.Text) = "") Then
    txt_nameEntry.Text = ""
    txt_nameEntry.SetFocus
  Else
    unity_main.m_rptNameEntry = Trim(txt_nameEntry.Text)
    Unload frm_dynRptName
  End If
End Sub

Private Sub cmd_clear_Click()
  
  unity_main.errorstring = "Dynamic Report Name Entry screen Clear button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  txt_nameEntry.Text = ""
  txt_nameEntry.SetFocus
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub








