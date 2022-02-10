VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_csvCfg 
   Caption         =   "CSV Report Configuration"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
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
   ScaleHeight     =   7290
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   480
      Top             =   6360
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7290
      FormDesignWidth =   9630
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   645
      Left            =   6360
      TabIndex        =   0
      Top             =   6240
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
      Caption         =   "frm_csvCfg.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_csvCfg.frx":002C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_csvCfg.frx":004C
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   645
      Left            =   3720
      TabIndex        =   26
      Top             =   6240
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
      Caption         =   "frm_csvCfg.frx":0068
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_csvCfg.frx":00A0
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_csvCfg.frx":00C0
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Left            =   360
      Top             =   240
      Width           =   8895
      _ExtentX        =   15690
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
      Caption         =   "frm_csvCfg.frx":00DC
      BackColor       =   -2147483633
      ForeColor       =   12582912
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_csvCfg.frx":014E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_csvCfg.frx":016E
   End
   Begin HexUniControls.ctlUniFrameXP frame_userInfo 
      Height          =   2535
      Left            =   360
      Top             =   3480
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_csvCfg.frx":018A
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_csvCfg.frx":01D6
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_csvCfg.frx":01F6
      Begin HexUniControls.ctlUniCheckXP chk_input 
         Height          =   450
         Index           =   8
         Left            =   2200
         TabIndex        =   16
         Top             =   1860
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":0212
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":0240
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":0260
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input 
         Height          =   450
         Index           =   7
         Left            =   2200
         TabIndex        =   15
         Top             =   1360
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":027C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":02AA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":02CA
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input 
         Height          =   450
         Index           =   6
         Left            =   2200
         TabIndex        =   14
         Top             =   860
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":02E6
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":0314
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":0334
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input 
         Height          =   450
         Index           =   5
         Left            =   2200
         TabIndex        =   13
         Top             =   360
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":0350
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":037E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":039E
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input 
         Height          =   450
         Index           =   4
         Left            =   300
         TabIndex        =   12
         Top             =   1860
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":03BA
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":03E8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":0408
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input 
         Height          =   450
         Index           =   3
         Left            =   300
         TabIndex        =   11
         Top             =   1360
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":0424
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":0452
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":0472
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input 
         Height          =   450
         Index           =   2
         Left            =   300
         TabIndex        =   10
         Top             =   860
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":048E
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":04BC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":04DC
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input 
         Height          =   450
         Index           =   1
         Left            =   300
         TabIndex        =   9
         Top             =   360
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":04F8
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":0526
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":0546
         ShowFocus       =   -1  'True
      End
   End
   Begin HexUniControls.ctlUniFrameXP frame_propInfo 
      Height          =   2535
      Left            =   5160
      Top             =   3480
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_csvCfg.frx":0562
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_csvCfg.frx":05AA
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_csvCfg.frx":05CA
      Begin HexUniControls.ctlUniCheckXP chk_propND 
         Height          =   450
         Left            =   2200
         TabIndex        =   22
         Top             =   860
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":05E6
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":060A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":062A
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_propSlope 
         Height          =   450
         Left            =   2205
         TabIndex        =   24
         Top             =   1860
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":0646
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":0670
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":0690
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_propIntercept 
         Height          =   450
         Left            =   2205
         TabIndex        =   23
         Top             =   1360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":06AC
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":06DE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":06FE
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_propOutlier 
         Height          =   450
         Left            =   2200
         TabIndex        =   21
         Top             =   360
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":071A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":0748
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":0768
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_propSResid 
         Height          =   450
         Left            =   300
         TabIndex        =   20
         Top             =   1860
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":0784
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":07B8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":07D8
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_propMDist 
         Height          =   450
         Left            =   300
         TabIndex        =   19
         Top             =   1360
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":07F4
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":0828
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":0848
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_propValue 
         Height          =   450
         Left            =   300
         TabIndex        =   18
         Top             =   860
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":0864
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":088E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":08AE
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_propName 
         Height          =   450
         Left            =   300
         TabIndex        =   17
         Top             =   360
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":08CA
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":08F2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":0912
         ShowFocus       =   -1  'True
      End
   End
   Begin HexUniControls.ctlUniFrameXP frame_miscInfo 
      Height          =   2535
      Left            =   360
      Top             =   840
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_csvCfg.frx":092E
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_csvCfg.frx":0980
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_csvCfg.frx":09A0
      Begin HexUniControls.ctlUniCheckXP chk_dataQuotes 
         Height          =   450
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":09BC
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":0A08
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":0A28
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_prodName 
         Height          =   450
         Left            =   5040
         TabIndex        =   6
         Top             =   860
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":0A44
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":0A72
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":0A92
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_sampComment 
         Height          =   450
         Left            =   5040
         TabIndex        =   8
         Top             =   1860
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":0AAE
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":0AEA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":0B0A
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_sampID 
         Height          =   450
         Left            =   5040
         TabIndex        =   7
         Top             =   1360
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":0B26
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":0B58
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":0B78
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_time 
         Height          =   450
         Left            =   240
         TabIndex        =   4
         Top             =   1860
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":0B94
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":0BBC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":0BDC
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_date 
         Height          =   450
         Left            =   240
         TabIndex        =   3
         Top             =   1360
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":0BF8
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":0C20
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":0C40
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_dateTime 
         Height          =   450
         Left            =   240
         TabIndex        =   2
         Top             =   860
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":0C5C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":0C92
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":0CB2
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_serialNum 
         Height          =   450
         Left            =   5040
         TabIndex        =   5
         Top             =   360
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_csvCfg.frx":0CCE
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_csvCfg.frx":0D16
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_csvCfg.frx":0D36
         ShowFocus       =   -1  'True
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_setDflts 
      Height          =   645
      Left            =   1080
      TabIndex        =   25
      Top             =   6240
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
      Caption         =   "frm_csvCfg.frx":0D52
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_csvCfg.frx":0D90
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_csvCfg.frx":0DB0
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   120
      Top             =   6840
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
      Top             =   6360
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_csvCfg.frx":0DCC
   End
End
Attribute VB_Name = "frm_csvCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Configuration file parameters
Private m_csvDataQuotes As Integer
Private m_csvDateTime As Integer
Private m_csvDate As Integer
Private m_csvTime As Integer
Private m_csvSerialNum As Integer
Private m_csvProdName As Integer
Private m_csvSampleID As Integer
Private m_csvSampleComment As Integer
Private m_csvPropName As Integer
Private m_csvPropValue As Integer
Private m_csvPropMDist As Integer
Private m_csvPropSResid As Integer
Private m_csvPropOutlier As Integer
Private m_csvPropND As Integer
Private m_csvPropIntercept As Integer
Private m_csvPropSlope As Integer
Private m_csvUserInputs(1 To MAX_MAN_INPUTS) As Boolean

' Misc variables
Private m_badCSVIniVal As Boolean

Public Sub load_cfg(setupCfg As Boolean)
  Dim ii As Integer
  Dim fileName As String
  Dim exist As Boolean
  Dim errMsg As String
  Dim uniMsg As String

  ' Setup screen default values
  m_csvDataQuotes = 1
  m_csvDateTime = 0
  m_csvDate = 1
  m_csvTime = 1
  m_csvSerialNum = 0
  m_csvProdName = 1
  m_csvSampleID = 1
  m_csvSampleComment = 1
  
  For ii = 1 To MAX_MAN_INPUTS
    m_csvUserInputs(ii) = False
  Next ii
  
  m_csvPropName = 1
  m_csvPropValue = 1
  m_csvPropMDist = 0
  m_csvPropSResid = 0
  m_csvPropOutlier = 0
  m_csvPropND = 0
  m_csvPropIntercept = 0
  m_csvPropSlope = 0

  m_badCSVIniVal = False
  fileName = (CFG_DIR & CSV_REPORT_CFG_FILE)
  
  exist = CFile.st_FileExist(fileName)
  
  If (exist = True) Then
    ' Setup default values
    unity_main.m_fileVersion = INFOSTAR_VER
    
    If (load_csv_file_vals(fileName) = True) Then
      ' Check for invalid file version
      If (unity_main.m_fileVersion <> INFOSTAR_VER) Then
        unity_main.errorstring = (fileName & " had incompatible value. Version was " & unity_main.m_fileVersion & "; updated to " & INFOSTAR_VER)
        unity_main.write_error
        unity_main.m_fileVersion = INFOSTAR_VER
        m_badCSVIniVal = True
      End If
    End If
  End If
  
  If (m_csvDataQuotes <> 0) Then
    frm_csvCfg.chk_dataQuotes.Value = 1
  Else
    frm_csvCfg.chk_dataQuotes.Value = 0
  End If
      
  If (m_csvDateTime <> 0) Then
    frm_csvCfg.chk_dateTime.Value = 1
  Else
    frm_csvCfg.chk_dateTime.Value = 0
  End If
      
  If (m_csvDate <> 0) Then
    frm_csvCfg.chk_date.Value = 1
  Else
    frm_csvCfg.chk_date.Value = 0
  End If
      
  If (m_csvTime <> 0) Then
    frm_csvCfg.chk_time.Value = 1
  Else
    frm_csvCfg.chk_time.Value = 0
  End If
      
  If (m_csvSerialNum <> 0) Then
    frm_csvCfg.chk_serialNum.Value = 1
  Else
    frm_csvCfg.chk_serialNum.Value = 0
  End If
      
  If (m_csvProdName <> 0) Then
    frm_csvCfg.chk_prodName.Value = 1
  Else
    frm_csvCfg.chk_prodName.Value = 0
  End If
      
  If (m_csvSampleID <> 0) Then
    frm_csvCfg.chk_sampID.Value = 1
  Else
    frm_csvCfg.chk_sampID.Value = 0
  End If
      
  If (m_csvSampleComment <> 0) Then
    frm_csvCfg.chk_sampComment.Value = 1
  Else
    frm_csvCfg.chk_sampComment.Value = 0
  End If
      
  For ii = 1 To MAX_MAN_INPUTS
    If (m_csvUserInputs(ii) = True) Then
      frm_csvCfg.chk_input(ii).Value = 1
    Else
      frm_csvCfg.chk_input(ii).Value = 0
    End If
  Next ii
    
  If (m_csvPropName <> 0) Then
    frm_csvCfg.chk_propName.Value = 1
  Else
    frm_csvCfg.chk_propName.Value = 0
  End If
        
  If (m_csvPropValue <> 0) Then
    frm_csvCfg.chk_propValue.Value = 1
  Else
    frm_csvCfg.chk_propValue.Value = 0
  End If
    
  If (m_csvPropMDist <> 0) Then
    frm_csvCfg.chk_propMDist.Value = 1
  Else
    frm_csvCfg.chk_propMDist.Value = 0
  End If
    
  If (m_csvPropSResid <> 0) Then
    frm_csvCfg.chk_propSResid.Value = 1
  Else
    frm_csvCfg.chk_propSResid.Value = 0
  End If

  If (m_csvPropOutlier <> 0) Then
    frm_csvCfg.chk_propOutlier.Value = 1
  Else
    frm_csvCfg.chk_propOutlier.Value = 0
  End If
      
  If (m_csvPropND <> 0) Then
    frm_csvCfg.chk_propND.Value = 1
  Else
    frm_csvCfg.chk_propND.Value = 0
  End If
      
  If (m_csvPropIntercept <> 0) Then
    frm_csvCfg.chk_propIntercept.Value = 1
  Else
    frm_csvCfg.chk_propIntercept.Value = 0
  End If
      
  If (m_csvPropSlope <> 0) Then
    frm_csvCfg.chk_propSlope.Value = 1
  Else
    frm_csvCfg.chk_propSlope.Value = 0
  End If
  
  ' Check if ini file had bad value
  If (m_badCSVIniVal = True) Then
    unity_main.errorstring = (fileName & " had incompatible value(s). Updated file with proper/default value(s).")
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg7", "%1 had incompatible value(s). Updated file with proper/default value(s)", fileName)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg3", "%1. Refer to system log for more details", uniMsg), vbOKOnly
    Call save_cfg(False)
  Else
    ' Check if to create configuration file for first time
    If (exist = False) Then
      Call save_cfg(False)
    End If
  End If
  
  If (setupCfg = True) Then
    ' Copy loaded CSV config into system operational variables
    unity_main.m_csvDataQuotes = m_csvDataQuotes
    unity_main.m_csvDateTime = m_csvDateTime
    unity_main.m_csvDate = m_csvDate
    unity_main.m_csvTime = m_csvTime
    unity_main.m_csvSerialNum = m_csvSerialNum
    unity_main.m_csvProdName = m_csvProdName
    unity_main.m_csvSampleID = m_csvSampleID
    unity_main.m_csvSampleComment = m_csvSampleComment
    unity_main.m_csvPropName = m_csvPropName
    unity_main.m_csvPropValue = m_csvPropValue
    unity_main.m_csvPropMDist = m_csvPropMDist
    unity_main.m_csvPropSResid = m_csvPropSResid
    unity_main.m_csvPropOutlier = m_csvPropOutlier
    unity_main.m_csvPropND = m_csvPropND
    unity_main.m_csvPropIntercept = m_csvPropIntercept
    unity_main.m_csvPropSlope = m_csvPropSlope
  
    For ii = 1 To MAX_MAN_INPUTS
      CSVUserInputs(ii) = m_csvUserInputs(ii)
    Next ii
  End If
End Sub

Public Sub write_csv_report()
  Dim printStrg As String
  Dim inputStrg As String
  Dim fileName As String
  Dim ii As Integer
  Dim numprops As Integer
  Dim tmpStrg As String
  Dim errMsg As String
  Dim uniFile As New clsUniFile
  Dim uniMsg As String

  ' Check if to include Date-Time (military format)
  If (unity_main.m_csvDateTime = 1) Then
    If (printStrg <> "") Then
      printStrg = printStrg & ","
    End If

    ' Check if to enclose data within quotes
    If (unity_main.m_csvDataQuotes = 1) Then
      printStrg = printStrg & (Chr(34) & unity_main.lbl_miltime.Caption & Chr(34))
    Else
      printStrg = printStrg & unity_main.lbl_miltime.Caption
    End If
  End If

  ' Check if to include Date
  If (unity_main.m_csvDate = 1) Then
    If (printStrg <> "") Then
      printStrg = printStrg & ","
    End If

    ' Check if to enclose data within quotes
    If (unity_main.m_csvDataQuotes = 1) Then
      printStrg = printStrg & (Chr(34) & unity_main.lbl_date.Caption & Chr(34))
    Else
      printStrg = printStrg & unity_main.lbl_date.Caption
    End If
  End If

  ' Check if to include Time
  If (unity_main.m_csvTime = 1) Then
    If (printStrg <> "") Then
      printStrg = printStrg & ","
    End If

    ' Check if to enclose data within quotes
    If (unity_main.m_csvDataQuotes = 1) Then
      printStrg = printStrg & (Chr(34) & unity_main.lbl_time.Caption & Chr(34))
    Else
      printStrg = printStrg & unity_main.lbl_time.Caption
    End If
  End If

  ' Check if to include System Serial Number
  If (unity_main.m_csvSerialNum = 1) Then
    If (printStrg <> "") Then
      printStrg = printStrg & ","
    End If

    ' Check if to enclose data within quotes
    If (unity_main.m_csvDataQuotes = 1) Then
      printStrg = printStrg & (Chr(34) & unity_main.m_sysSerialNum & Chr(34))
    Else
      printStrg = printStrg & unity_main.m_sysSerialNum
    End If
  End If

  ' Check if to include Product
  If (unity_main.m_csvProdName = 1) Then
    If (printStrg <> "") Then
      printStrg = printStrg & ","
    End If

    ' Check if to enclose data within quotes
    If (unity_main.m_csvDataQuotes = 1) Then
      printStrg = printStrg & (Chr(34) & Trim(unity_main.lblProd1.Caption) & Chr(34))
    Else
      printStrg = printStrg & Trim(unity_main.lblProd1.Caption)
    End If
  End If
  
  ' Check if to include Sample ID
  If (unity_main.m_csvSampleID = 1) Then
    If (printStrg <> "") Then
      printStrg = printStrg & ","
    End If

    ' Check if to enclose data within quotes
    If (unity_main.m_csvDataQuotes = 1) Then
      printStrg = printStrg & (Chr(34) & Trim(unity_main.txtsamplename.Text) & Chr(34))
    Else
      printStrg = printStrg & Trim(unity_main.txtsamplename.Text)
    End If
  End If
  
  ' Check if to include Comment
  If (unity_main.m_csvSampleComment = 1) Then
    If (printStrg <> "") Then
      printStrg = printStrg & ","
    End If

    ' Check if to enclose data within quotes
    If (unity_main.m_csvDataQuotes = 1) Then
      printStrg = printStrg & (Chr(34) & Trim(unity_main.txtsampcomment.Text) & Chr(34))
    Else
      printStrg = printStrg & Trim(unity_main.txtsampcomment.Text)
    End If
  End If

  For ii = 1 To MAX_MAN_INPUTS
    ' Check if to include Input
    If (CSVUserInputs(ii) = True) Then
      frm_buttoncfg.ss_buttonconfig.Col = ii
      frm_buttoncfg.ss_buttonconfig.Row = 1
      
      ' Check if input enabled
      If (unity_main.m_useMIV = True) And (frm_buttoncfg.ss_buttonconfig.Value = 1) Then
        ' Setup for text entry/list box selection field
        frm_buttoncfg.ss_buttonconfig.Col = ii
        frm_buttoncfg.ss_buttonconfig.Row = 2
  
        ' Check if using text entry
        If (frm_buttoncfg.ss_buttonconfig.Value = 0) Then
          inputStrg = Trim(frm_scanname.txtbx(ii).Text)
        Else    ' Using list
          inputStrg = Trim(frm_scanname.combo(ii).Text)
        End If
      Else
        inputStrg = MLSupport.GSS("Headers", "na", "NA")
      End If

      If (printStrg <> "") Then
        printStrg = printStrg & ","
      End If

      ' Check if to enclose data within quotes
      If (unity_main.m_csvDataQuotes = 1) Then
        printStrg = printStrg & (Chr(34) & inputStrg & Chr(34))
      Else
        printStrg = printStrg & inputStrg
      End If
    End If
  Next ii

  numprops = Trim(frmedmod.numprops.Text)
  
  For ii = 1 To numprops
    ' Check if to include Property Name
    If (unity_main.m_csvPropName = 1) Then
      unity_main.fpspread_pred.Row = ii
      unity_main.fpspread_pred.Col = 1

      If (printStrg <> "") Then
        printStrg = printStrg & ","
      End If

      ' Check if to enclose data within quotes
      If (unity_main.m_csvDataQuotes = 1) Then
        printStrg = printStrg & (Chr(34) & Trim(unity_main.fpspread_pred.Text) & Chr(34))
      Else
        printStrg = printStrg & Trim(unity_main.fpspread_pred.Text)
      End If
    End If
    
    ' Check if to include Property Value
    If (unity_main.m_csvPropValue = 1) Then
      unity_main.fpspread_pred.Row = ii
      unity_main.fpspread_pred.Col = 2

      If (printStrg <> "") Then
        printStrg = printStrg & ","
      End If

      ' Check if to enclose data within quotes
      If (unity_main.m_csvDataQuotes = 1) Then
        printStrg = printStrg & (Chr(34) & Trim(unity_main.fpspread_pred.Text) & Chr(34))
      Else
        printStrg = printStrg & Trim(unity_main.fpspread_pred.Text)
      End If
    End If
    
    ' Check if to include Property M-Distance
    If (unity_main.m_csvPropMDist = 1) Then
      If (printStrg <> "") Then
        printStrg = printStrg & ","
      End If

      ' Check if no prediction made
      If (unity_main.lstmd.ListCount = 0) Then
        tmpStrg = unity_main.m_noOLVal
      Else
        tmpStrg = unity_main.lstmd.List(ii - 1)
      End If
      
      ' Check if to enclose data within quotes
      If (unity_main.m_csvDataQuotes = 1) Then
        printStrg = printStrg & (Chr(34) & tmpStrg & Chr(34))
      Else
        printStrg = printStrg & tmpStrg
      End If
    End If
    
    ' Check if to include Property S-Residual
    If (unity_main.m_csvPropSResid = 1) Then
      If (printStrg <> "") Then
        printStrg = printStrg & ","
      End If

      ' Check if no prediction made
      If (unity_main.lstresrat.ListCount = 0) Then
        tmpStrg = unity_main.m_noOLVal
      Else
        tmpStrg = unity_main.lstresrat.List(ii - 1)
      End If
      
      ' Check if to enclose data within quotes
      If (unity_main.m_csvDataQuotes = 1) Then
        printStrg = printStrg & (Chr(34) & tmpStrg & Chr(34))
      Else
        printStrg = printStrg & tmpStrg
      End If
    End If
  
    ' Check if to include Property Outlier
    If (unity_main.m_csvPropOutlier = 1) Then
      If (printStrg <> "") Then
        printStrg = printStrg & ","
      End If

      ' Check if no prediction made
      If (unity_main.lst_qual.ListCount = 0) Then
        tmpStrg = unity_main.m_noOLVal
      Else
        tmpStrg = unity_main.lst_qual.List(ii - 1)
      End If
      
      ' Check if to enclose data within quotes
      If (unity_main.m_csvDataQuotes = 1) Then
        printStrg = printStrg & (Chr(34) & tmpStrg & Chr(34))
      Else
        printStrg = printStrg & tmpStrg
      End If
    End If
  
    ' Check if to include Property Neighborhood Distance
    If (unity_main.m_csvPropND = 1) Then
      If (printStrg <> "") Then
        printStrg = printStrg & ","
      End If

      ' Check if no prediction made
      If (unity_main.lst_nd.ListCount = 0) Then
        tmpStrg = unity_main.m_noOLVal
      Else
        tmpStrg = unity_main.lst_nd.List(ii - 1)
      End If
      
      ' Check if to enclose data within quotes
      If (unity_main.m_csvDataQuotes = 1) Then
        printStrg = printStrg & (Chr(34) & tmpStrg & Chr(34))
      Else
        printStrg = printStrg & tmpStrg
      End If
    End If
    
    ' Check if to include Property Intercept
    If (unity_main.m_csvPropIntercept = 1) Then
      If (printStrg <> "") Then
        printStrg = printStrg & ","
      End If

      ' Check if no prediction made
      If (unity_main.lstint.ListCount = 0) Then
        tmpStrg = unity_main.m_noOLVal
      Else
        tmpStrg = unity_main.lstint.List(ii - 1)
      End If
      
      ' Check if to enclose data within quotes
      If (unity_main.m_csvDataQuotes = 1) Then
        printStrg = printStrg & (Chr(34) & tmpStrg & Chr(34))
      Else
        printStrg = printStrg & tmpStrg
      End If
    End If
  
    ' Check if to include Property Slope
    If (unity_main.m_csvPropSlope = 1) Then
      If (printStrg <> "") Then
        printStrg = printStrg & ","
      End If

      ' Check if no prediction made
      If (unity_main.lstslope.ListCount = 0) Then
        tmpStrg = unity_main.m_noOLVal
      Else
        tmpStrg = unity_main.lstslope.List(ii - 1)
      End If
      
      ' Check if to enclose data within quotes
      If (unity_main.m_csvDataQuotes = 1) Then
        printStrg = printStrg & (Chr(34) & tmpStrg & Chr(34))
      Else
        printStrg = printStrg & tmpStrg
      End If
    End If
  Next ii

  uniMsg = MLSupport.GGS_Params("unity_main.statMsg27", "Writing sample results to csv file: %1", unity_main.m_saveCSVFile)
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Writing sample results to csv file: " & unity_main.m_saveCSVFile), uniMsg)
  
  On Error GoTo FILE_ERROR
  CreatePath CFile.st_FilePath(unity_main.m_saveCSVFile)
  
  If (uniFile.st_FileExist(unity_main.m_saveCSVFile) = True) Then
    If (uniFile.OpenFileAppend(unity_main.m_saveCSVFile) = False) Then GoTo FILE_ERROR
  Else
    If (uniFile.OpenFileWrite(unity_main.m_saveCSVFile) = False) Then GoTo FILE_ERROR
    uniFile.WriteBOM fe_UTF16LE
  End If
    
  uniFile.WriteUnicodeLine printStrg
  uniFile.Flush
  uniFile.CloseFile
  Exit Sub

FILE_ERROR:
  uniFile.CloseFile
  errMsg = (unity_main.m_saveCSVFile & " file write error. " & Error$)
  unity_main.errorstring = errMsg
  unity_main.write_error
  uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", unity_main.m_saveCSVFile, Error$)
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub

Private Function load_csv_file_vals(srcFile As String) As Boolean
  Dim inString As String
  Dim xx As String
  Dim tmpStrg As String
  Dim cfgVar As String
  Dim varVal As String
  Dim strlen As Integer
  Dim lineCnt As Integer
  Dim errMsg As String
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim rc As Boolean
  Dim uniMsg As String

  If (uniFile.OpenFileRead(srcFile) = True) Then
    On Error GoTo FILE_ERROR
    fEncoding = uniFile.ReadBOM
  
    ' Process each line in .ini file
    While Not (uniFile.EOF())
      ' Read line from file
      On Error GoTo FILE_ERROR
      lineCnt = lineCnt + 1
      
      If (fEncoding = fe_ANSI) Then
        rc = uniFile.ReadAnsiLine(inString)
      Else
        rc = uniFile.ReadUnicodeLine(inString)
      End If
      
      If (rc = True) Then
        ' Get variable name and its value
        xx = InStr(1, inString, "=")
        strlen = Len(inString)
        tmpStrg = Trim(Mid(inString, 1, xx - 1))
        cfgVar = LCase(tmpStrg)
        varVal = Trim(Mid(inString, xx + 1))
          
        ' Process value by variable name
        On Error GoTo BAD_INI_VALUE
        Select Case (cfgVar)
          Case "version"
            unity_main.m_fileVersion = Trim(varVal)
          Case "include_dataquotes"
            m_csvDataQuotes = CInt(varVal)
          Case "include_datetime"
            m_csvDateTime = CInt(varVal)
          Case "include_date"
            m_csvDate = CInt(varVal)
          Case "include_time"
            m_csvTime = CInt(varVal)
          Case "include_serialnum"
            m_csvSerialNum = CInt(varVal)
          Case "include_product"
            m_csvProdName = CInt(varVal)
          Case "include_sampleid"
            m_csvSampleID = CInt(varVal)
          Case "include_comment"
            m_csvSampleComment = CInt(varVal)
          Case "include_input1"
            m_csvUserInputs(1) = CBool(varVal)
          Case "include_input2"
            m_csvUserInputs(2) = CBool(varVal)
          Case "include_input3"
            m_csvUserInputs(3) = CBool(varVal)
          Case "include_input4"
            m_csvUserInputs(4) = CBool(varVal)
          Case "include_input5"
            m_csvUserInputs(5) = CBool(varVal)
          Case "include_input6"
            m_csvUserInputs(6) = CBool(varVal)
          Case "include_input7"
            m_csvUserInputs(7) = CBool(varVal)
          Case "include_input8"
            m_csvUserInputs(8) = CBool(varVal)
          Case "include_propname"
            m_csvPropName = CInt(varVal)
          Case "include_propvalue"
            m_csvPropValue = CInt(varVal)
          Case "include_propmdist"
            m_csvPropMDist = CInt(varVal)
          Case "include_propsresid"
            m_csvPropSResid = CInt(varVal)
          Case "include_propoutlier"
            m_csvPropOutlier = CInt(varVal)
          Case "include_propnd"
            m_csvPropND = CInt(varVal)
          Case "include_propintercept"
            m_csvPropIntercept = CInt(varVal)
          Case "include_propslope"
            m_csvPropSlope = CInt(varVal)
        End Select
      End If
    Wend
  
    load_csv_file_vals = True
  Else
FILE_ERROR:
    errMsg = (srcFile & " file read error. " & Error$)
    unity_main.errorstring = errMsg
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg4", "%1 file read error. %2", srcFile, Error$)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    load_csv_file_vals = False
  End If
  
  uniFile.CloseFile
  Exit Function
  
BAD_INI_VALUE:
  unity_main.errorstring = (srcFile & " line " & CStr(lineCnt) & " had incompatible value. " & Trim(tmpStrg) & " - " & varVal & "; will use default value")
  unity_main.write_error
  m_badCSVIniVal = True
  Resume Next
End Function

Private Sub save_cfg(setupCfg As Boolean)
  Dim ii As Integer
  Dim fileName As String
  Dim uniFile As New clsUniFile
  Dim errMsg As String
  Dim uniMsg As String

  m_csvDataQuotes = frm_csvCfg.chk_dataQuotes.Value
  m_csvDateTime = frm_csvCfg.chk_dateTime.Value
  m_csvDate = frm_csvCfg.chk_date.Value
  m_csvTime = frm_csvCfg.chk_time.Value
  m_csvSerialNum = frm_csvCfg.chk_serialNum.Value
  m_csvProdName = frm_csvCfg.chk_prodName.Value
  m_csvSampleID = frm_csvCfg.chk_sampID.Value
  m_csvSampleComment = frm_csvCfg.chk_sampComment.Value
      
  For ii = 1 To MAX_MAN_INPUTS
    m_csvUserInputs(ii) = frm_csvCfg.chk_input(ii).Value
  Next ii
    
  m_csvPropName = frm_csvCfg.chk_propName.Value
  m_csvPropValue = frm_csvCfg.chk_propValue.Value
  m_csvPropMDist = frm_csvCfg.chk_propMDist.Value
  m_csvPropSResid = frm_csvCfg.chk_propSResid.Value
  m_csvPropOutlier = frm_csvCfg.chk_propOutlier.Value
  m_csvPropND = frm_csvCfg.chk_propND.Value
  m_csvPropIntercept = frm_csvCfg.chk_propIntercept.Value
  m_csvPropSlope = frm_csvCfg.chk_propSlope.Value

  fileName = (CFG_DIR & CSV_REPORT_CFG_FILE)

  If (uniFile.OpenFileWrite(fileName) = True) Then
    On Error GoTo FILE_ERROR
    uniFile.WriteBOM fe_UTF16LE
    uniFile.WriteUnicodeLine ("Version=" & INFOSTAR_VER)
    uniFile.WriteUnicodeLine ("Include_DataQuotes=" & m_csvDataQuotes)
    uniFile.WriteUnicodeLine ("Include_DateTime=" & m_csvDateTime)
    uniFile.WriteUnicodeLine ("Include_Date=" & m_csvDate)
    uniFile.WriteUnicodeLine ("Include_Time=" & m_csvTime)
    uniFile.WriteUnicodeLine ("Include_SerialNum=" & m_csvSerialNum)
    uniFile.WriteUnicodeLine ("Include_Product=" & m_csvProdName)
    uniFile.WriteUnicodeLine ("Include_SampleID=" & m_csvSampleID)
    uniFile.WriteUnicodeLine ("Include_Comment=" & m_csvSampleComment)
    
    For ii = 1 To MAX_MAN_INPUTS
      uniFile.WriteUnicodeLine ("Include_Input" & ii & "=" & m_csvUserInputs(ii))
    Next ii
    
    uniFile.WriteUnicodeLine ("Include_PropName=" & m_csvPropName)
    uniFile.WriteUnicodeLine ("Include_PropValue=" & m_csvPropValue)
    uniFile.WriteUnicodeLine ("Include_PropMDist=" & m_csvPropMDist)
    uniFile.WriteUnicodeLine ("Include_PropSResid=" & m_csvPropSResid)
    uniFile.WriteUnicodeLine ("Include_PropOutlier=" & m_csvPropOutlier)
    uniFile.WriteUnicodeLine ("Include_PropND=" & m_csvPropND)
    uniFile.WriteUnicodeLine ("Include_PropIntercept=" & m_csvPropIntercept)
    uniFile.WriteUnicodeLine ("Include_PropSlope=" & m_csvPropSlope)
    uniFile.Flush
    
    If (setupCfg = True) Then
      ' Copy saved CSV config into system operational variables
      unity_main.m_csvDataQuotes = m_csvDataQuotes
      unity_main.m_csvDateTime = m_csvDateTime
      unity_main.m_csvDate = m_csvDate
      unity_main.m_csvTime = m_csvTime
      unity_main.m_csvSerialNum = m_csvSerialNum
      unity_main.m_csvProdName = m_csvProdName
      unity_main.m_csvSampleID = m_csvSampleID
      unity_main.m_csvSampleComment = m_csvSampleComment
      unity_main.m_csvPropName = m_csvPropName
      unity_main.m_csvPropValue = m_csvPropValue
      unity_main.m_csvPropMDist = m_csvPropMDist
      unity_main.m_csvPropSResid = m_csvPropSResid
      unity_main.m_csvPropOutlier = m_csvPropOutlier
      unity_main.m_csvPropND = m_csvPropND
      unity_main.m_csvPropIntercept = m_csvPropIntercept
      unity_main.m_csvPropSlope = m_csvPropSlope

      For ii = 1 To MAX_MAN_INPUTS
        CSVUserInputs(ii) = m_csvUserInputs(ii)
      Next ii
    End If
    
    unity_main.errorstring = ("User saved new settings for configuration file: " & fileName)
    unity_main.write_error (LOG_DBG_LEVEL1)
  Else
FILE_ERROR:
    errMsg = (fileName & " file write error. " & Error$)
    unity_main.errorstring = errMsg
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", fileName, Error$)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  End If
  
  uniFile.CloseFile
End Sub

Private Sub cmd_cancel_Click()

  unity_main.errorstring = "CSV Report Configuration screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Unload frm_csvCfg
End Sub

Private Sub cmd_save_Click()

  unity_main.errorstring = "CSV Report Configuration screen Save Changes button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Call save_cfg(True)
  Unload frm_csvCfg
End Sub

Private Sub cmd_setDflts_Click()
  Dim ii As Integer

  unity_main.errorstring = "CSV Report Configuration screen Set to Defaults button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  ' Set selections to default
  chk_dataQuotes.Value = 1
  chk_dateTime.Value = 0
  chk_date.Value = 1
  chk_time.Value = 1
  chk_serialNum.Value = 0
  chk_sampID.Value = 1
  chk_sampComment.Value = 1
  chk_prodName.Value = 1
  
  For ii = 1 To MAX_MAN_INPUTS
    chk_input(ii).Value = 0
  Next ii
  
  chk_propName.Value = 1
  chk_propValue.Value = 1
  chk_propMDist.Value = 0
  chk_propSResid.Value = 0
  chk_propOutlier.Value = 0
  chk_propIntercept.Value = 0
  chk_propSlope.Value = 0
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub








