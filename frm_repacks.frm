VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_repacks 
   Caption         =   "Individual Repack Readings"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
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
   Icon            =   "frm_repacks.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin FPUSpreadADO.fpSpread ss_repacks 
      Height          =   5175
      Left            =   480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   9495
      _Version        =   458752
      _ExtentX        =   16748
      _ExtentY        =   9128
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   -2147483633
      MaxCols         =   102
      MaxRows         =   50
      OperationMode   =   1
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   0
      SpreadDesigner  =   "frm_repacks.frx":030A
      UserResize      =   1
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   1320
      Top             =   5880
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6360
      FormDesignWidth =   10485
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_exit 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   4242
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
      Caption         =   "frm_repacks.frx":057E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_repacks.frx":05A6
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_repacks.frx":05C6
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   1920
      Top             =   5760
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
      Left            =   480
      Top             =   5760
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_repacks.frx":05E2
   End
End
Attribute VB_Name = "frm_repacks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_exit_Click()
  Call unity_main.restart_loop(LOG_DBG_LEVEL3, "Individual Repack Readings screen Exit button selected")
  frm_repacks.Visible = False
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub








