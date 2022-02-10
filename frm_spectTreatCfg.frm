VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_spectTreatCfg 
   Caption         =   "Spectrum Treatment Configuration"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5670
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniFrameXP frame_saveSpectrum 
      Height          =   1455
      Left            =   480
      Top             =   5880
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2566
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_spectTreatCfg.frx":0000
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_spectTreatCfg.frx":003A
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_spectTreatCfg.frx":005A
      Begin HexUniControls.ctlUniRadioXP opt_saveTreated 
         Height          =   450
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
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
         Caption         =   "frm_spectTreatCfg.frx":0076
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_spectTreatCfg.frx":00C0
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_spectTreatCfg.frx":00E0
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniRadioXP opt_saveBoth 
         Height          =   450
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   4455
         _ExtentX        =   7858
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
         Caption         =   "frm_spectTreatCfg.frx":00FC
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_spectTreatCfg.frx":014A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_spectTreatCfg.frx":016A
         ShowFocus       =   -1  'True
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   645
      Left            =   3360
      TabIndex        =   0
      Top             =   7560
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
      Caption         =   "frm_spectTreatCfg.frx":0186
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_spectTreatCfg.frx":01B2
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_spectTreatCfg.frx":01D2
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   480
      Top             =   8280
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8550
      FormDesignWidth =   5670
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   645
      Left            =   360
      TabIndex        =   1
      Top             =   7560
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
      Caption         =   "frm_spectTreatCfg.frx":01EE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_spectTreatCfg.frx":0226
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_spectTreatCfg.frx":0246
   End
   Begin HexUniControls.ctlUniFrameXP frame_smooth 
      Height          =   4905
      Left            =   480
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   8652
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_spectTreatCfg.frx":0262
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_spectTreatCfg.frx":02A2
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_spectTreatCfg.frx":02C2
      Begin HexUniControls.ctlUniCheckXP chk_useProgSmooth 
         Height          =   450
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   4455
         _ExtentX        =   7858
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
         Caption         =   "frm_spectTreatCfg.frx":02DE
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_spectTreatCfg.frx":0330
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_spectTreatCfg.frx":0350
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniComboBoxXP combo_smoothType 
         Height          =   450
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   794
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Tip             =   "frm_spectTreatCfg.frx":036C
         Sorted          =   0   'False
         HScroll         =   0   'False
         Style           =   3
         ButtonBackColor =   -2147483633
         ButtonForeColor =   0
         ButtonWidth     =   30
         Locked          =   0   'False
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
         TrapTab         =   0   'False
         ButtonStyle     =   4
         SelectorStyle   =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_spectTreatCfg.frx":038C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
      End
      Begin HexUniControls.ctlNumIncXP numInc_smoothNumPts 
         Height          =   600
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1058
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
         Text            =   "1"
         Min             =   1
         Max             =   30
         AllowSpace      =   0   'False
         BorderColor     =   -1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   -2147483633
         ButtonForeColor =   0
         ButtonStyle     =   4
         ButtonWidth     =   30
         Tip             =   ""
         MousePointer    =   0
         MouseIcon       =   "frm_spectTreatCfg.frx":03A8
         TrapTabKey      =   0   'False
      End
      Begin HexUniControls.ctlUniLabel lbl_smoothType 
         Height          =   450
         Left            =   2280
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
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
         Caption         =   "frm_spectTreatCfg.frx":03C4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_spectTreatCfg.frx":03FA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_spectTreatCfg.frx":041A
      End
      Begin HexUniControls.ctlUniLabel lbl_smoothNumPts 
         Height          =   600
         Left            =   1440
         Top             =   1680
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1058
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "@Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_spectTreatCfg.frx":0436
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_spectTreatCfg.frx":0486
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_spectTreatCfg.frx":04A6
      End
      Begin HexUniControls.ctlNumIncXP numInc_startSmoothNumPts 
         Height          =   600
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1058
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
         Text            =   "1"
         Min             =   1
         Max             =   30
         AllowSpace      =   0   'False
         BorderColor     =   -1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   -2147483633
         ButtonForeColor =   0
         ButtonStyle     =   4
         ButtonWidth     =   30
         Tip             =   ""
         MousePointer    =   0
         MouseIcon       =   "frm_spectTreatCfg.frx":04C2
         TrapTabKey      =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txt_startSmoothWvln 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3840
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frm_spectTreatCfg.frx":04DE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   0
         MultiLine       =   0   'False
         Alignment       =   1
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frm_spectTreatCfg.frx":04FE
         NoHideSel       =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_spectTreatCfg.frx":051E
      End
      Begin HexUniControls.ctlUniLabel lbl_startSmoothWvln 
         Height          =   375
         Left            =   1200
         Top             =   3840
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_spectTreatCfg.frx":053A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_spectTreatCfg.frx":0580
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_spectTreatCfg.frx":05A0
      End
      Begin HexUniControls.ctlUniTextBoxXP txt_endSmoothWvln 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   4320
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frm_spectTreatCfg.frx":05BC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   0
         MultiLine       =   0   'False
         Alignment       =   1
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frm_spectTreatCfg.frx":05DC
         NoHideSel       =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_spectTreatCfg.frx":05FC
      End
      Begin HexUniControls.ctlUniLabel lbl_endSmoothWvln 
         Height          =   375
         Left            =   1200
         Top             =   4320
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_spectTreatCfg.frx":0618
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_spectTreatCfg.frx":065A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_spectTreatCfg.frx":067A
      End
      Begin HexUniControls.ctlUniLabel lbl_startSmoothNumPts 
         Height          =   600
         Left            =   1440
         Top             =   1680
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1058
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "@Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_spectTreatCfg.frx":0696
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_spectTreatCfg.frx":06F8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_spectTreatCfg.frx":0718
      End
      Begin HexUniControls.ctlNumIncXP numInc_endSmoothNumPts 
         Height          =   600
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1058
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
         Text            =   "1"
         Min             =   1
         Max             =   30
         AllowSpace      =   0   'False
         BorderColor     =   -1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   -2147483633
         ButtonForeColor =   0
         ButtonStyle     =   4
         ButtonWidth     =   30
         Tip             =   ""
         MousePointer    =   0
         MouseIcon       =   "frm_spectTreatCfg.frx":0734
         TrapTabKey      =   0   'False
      End
      Begin HexUniControls.ctlUniLabel lbl_endSmoothNumPts 
         Height          =   600
         Left            =   1440
         Top             =   2400
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1058
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "@Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_spectTreatCfg.frx":0750
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_spectTreatCfg.frx":07AE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_spectTreatCfg.frx":07CE
      End
      Begin HexUniControls.ctlNumIncXP numInc_progSmoothRate 
         Height          =   600
         Left            =   120
         TabIndex        =   9
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1058
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
         Text            =   "1"
         Min             =   0
         Max             =   30
         AllowSpace      =   0   'False
         BorderColor     =   -1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   -2147483633
         ButtonForeColor =   0
         ButtonStyle     =   4
         ButtonWidth     =   30
         Tip             =   ""
         MousePointer    =   0
         MouseIcon       =   "frm_spectTreatCfg.frx":07EA
         TrapTabKey      =   0   'False
      End
      Begin HexUniControls.ctlUniLabel lbl_progSmoothRate 
         Height          =   600
         Left            =   1440
         Top             =   3120
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1058
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "@Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_spectTreatCfg.frx":0806
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_spectTreatCfg.frx":084C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_spectTreatCfg.frx":086C
      End
   End
   Begin HexUniControls.ctlUniCheckXP chk_enableSmooth 
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_spectTreatCfg.frx":0888
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_spectTreatCfg.frx":08DA
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_spectTreatCfg.frx":08FA
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniFormCaption uniTitle 
      Left            =   960
      Top             =   8280
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_spectTreatCfg.frx":0916
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   1560
      Top             =   8280
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
End
Attribute VB_Name = "frm_spectTreatCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_badIniVal As Boolean
Private m_enableSmooth As Boolean
Private m_endSmoothWvln As Double
Private m_endSmoothNumPts As Integer
Private m_progSmoothRate As Integer
Private m_saveSpectra As String
Private m_smoothNumPts As Integer
Private m_smoothTypeEnum As SMOOTH_TYPES
Private m_smoothType As String
Private m_startSmoothWvln As Double
Private m_startSmoothNumPts As Integer
Private m_useProgSmooth As Boolean

Private Const DEFLT_NUM_PTS = 7
Private Const DEFLT_PROG_RATE = 1
Private Const DEFLT_SMOOTH_TYPE = "BoxCar"
Private Const DEFLT_SAVE_SPECTRA = "Treated"

Public Sub save_cfg()
  Dim errMsg As String
  Dim uniFile As New clsUniFile
  Dim uniMsg As String
  
  ' Save spectrum treatment variables
  m_enableSmooth = CBool(chk_enableSmooth.Value)
  
  ' Save progressive smooth ending wavelength
  m_endSmoothWvln = CDbl(txt_endSmoothWvln.Text)
  
  ' Save progressive smooth ending number of end-sums pts
  m_endSmoothNumPts = numInc_endSmoothNumPts.Text
  
  ' Save progressive smooth rate
  m_progSmoothRate = numInc_progSmoothRate.Text
  
  ' Save save spectrum settings
  If (opt_saveTreated.Value = True) Then
    m_saveSpectra = "Treated"
  Else
    m_saveSpectra = "Both"
  End If
    
  ' Save smooth number of end-sums pts
  m_smoothNumPts = numInc_smoothNumPts.Text
  
  ' Save smooth type
  m_smoothTypeEnum = combo_smoothType.ListIndex + 1
    
  Select Case (m_smoothTypeEnum)
    Case ST_BOX_CAR               ' Box Car
      m_smoothType = "BoxCar"
    Case ST_TRIANGULAR            ' Triangular
      m_smoothType = "Triangular"
    Case ST_SAVITSKY_GOLAY        ' Savitsky Golay
      m_smoothType = "SavitskyGolay"
  End Select
  
  ' Save progressive smooth starting wavelength
  m_startSmoothWvln = CDbl(txt_startSmoothWvln.Text)
  
  ' Save progressive smooth starting number of end-sums pts
  m_startSmoothNumPts = numInc_startSmoothNumPts.Text
  
  ' Save use progressive algortihm
  m_useProgSmooth = CBool(chk_useProgSmooth.Value)
  
  ' Check if file can be created
  If (uniFile.OpenFileWrite(CFG_DIR & SPECT_TREAT_CFG_FILE) = True) Then
    On Error GoTo FILE_ERROR
    uniFile.WriteBOM fe_UTF16LE
    uniFile.WriteUnicodeLine "[signature settings]"
    uniFile.WriteUnicodeLine ("DevID=" & MS11CfgData.devID)
    uniFile.WriteUnicodeLine ("SmplTable=" & unity_main.m_smplTable)
    uniFile.WriteUnicodeLine ("ScanMode=" & unity_main.m_sysScanMode)
    uniFile.WriteUnicodeLine ("Version=" & INFOSTAR_VER)
    uniFile.WriteUnicodeLine "[treatment settings]"
    uniFile.WriteUnicodeLine ("EnableSmooth=" & m_enableSmooth)
    uniFile.WriteUnicodeLine ("EndSmoothWvln=" & m_endSmoothWvln)
    uniFile.WriteUnicodeLine ("EndSmoothNumPts=" & m_endSmoothNumPts)
    uniFile.WriteUnicodeLine ("ProgSmoothRate=" & m_progSmoothRate)
    uniFile.WriteUnicodeLine ("SaveSpectra=" & m_saveSpectra)
    uniFile.WriteUnicodeLine ("SmoothNumPts=" & m_smoothNumPts)
    uniFile.WriteUnicodeLine ("SmoothType=" & m_smoothType)
    uniFile.WriteUnicodeLine ("StartSmoothWvln=" & m_startSmoothWvln)
    uniFile.WriteUnicodeLine ("StartSmoothNumPts=" & m_startSmoothNumPts)
    uniFile.WriteUnicodeLine ("UseProgSmooth=" & m_useProgSmooth)
    uniFile.Flush
    
    ' Copy spectrum treatment config into system operational variables
    unity_main.m_enableTreatment = m_enableSmooth   ' OR all treatment options
    unity_main.m_enableSmooth = m_enableSmooth
    unity_main.m_endSmoothWvln = m_endSmoothWvln
    unity_main.m_endSmoothNumPts = m_endSmoothNumPts
    unity_main.m_progSmoothRate = m_progSmoothRate
    unity_main.m_saveSpectra = m_saveSpectra
    unity_main.m_smoothNumPts = m_smoothNumPts
    unity_main.m_smoothTypeEnum = m_smoothTypeEnum
    unity_main.m_smoothType = m_smoothType
    unity_main.m_startSmoothWvln = m_startSmoothWvln
    unity_main.m_startSmoothNumPts = m_startSmoothNumPts
    unity_main.m_useProgSmooth = m_useProgSmooth
  Else
FILE_ERROR:
    errMsg = ((CFG_DIR & SPECT_TREAT_CFG_FILE) & " file write error. " & Error$)
    unity_main.errorstring = errMsg
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", (CFG_DIR & SPECT_TREAT_CFG_FILE), Error$)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  End If
  
  uniFile.CloseFile
End Sub

Public Sub load_cfg()
  Dim fileName As String
  Dim exist As Boolean
  Dim iniString As String
  Dim lineCnt As Integer
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim rc As Boolean
  Dim errMsg As String
  Dim uniMsg As String
  
  m_badIniVal = False

  ' Initialize default setting values
  unity_main.m_fileDevID = 0
  unity_main.m_fileVersion = INFOSTAR_VER
  m_enableSmooth = False
  m_endSmoothWvln = ProdDfltData.endWvln
  m_endSmoothNumPts = DEFLT_NUM_PTS
  m_progSmoothRate = DEFLT_PROG_RATE
  m_saveSpectra = DEFLT_SAVE_SPECTRA
  m_smoothNumPts = DEFLT_NUM_PTS
  m_smoothTypeEnum = ST_BOX_CAR
  m_smoothType = DEFLT_SMOOTH_TYPE
  m_startSmoothWvln = ProdDfltData.startWvln
  m_startSmoothNumPts = DEFLT_NUM_PTS
  m_useProgSmooth = False

  ' Check if configuration file exists
  fileName = (CFG_DIR & SPECT_TREAT_CFG_FILE)
  exist = CFile.st_FileExist(fileName)
    
  If (exist = True) Then
    ' Load spectrum treatment configuration file
    If (uniFile.OpenFileRead(fileName) = False) Then GoTo FILE_ERROR
  
    On Error GoTo FILE_ERROR
    fEncoding = uniFile.ReadBOM
    
    ' Read first line of file
    lineCnt = lineCnt + 1
  
    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(iniString)
    Else
      rc = uniFile.ReadUnicodeLine(iniString)
    End If
      
    If (rc = False) Then GoTo FILE_ERROR
          
    ' Process each line in .ini file
    While Not (uniFile.EOF())
      Select Case (iniString)
        Case "[signature settings]"
          Call unity_main.load_file_signature_vals(fileName, uniFile, fEncoding, lineCnt, m_badIniVal)
          iniString = unity_main.m_iniString
      
          If (iniString <> "") Then
            If (unity_main.process_file_signature_vals(fileName, False, m_badIniVal) = False) Then
              GoTo LOAD_ERROR
            End If
          End If
        
        Case "[treatment settings]"
          If (load_cfg_settings(fileName, uniFile, fEncoding, lineCnt) = False) Then
            GoTo FILE_ERROR
          End If
        
        Case Else
          GoTo FILE_ERROR
      End Select
    Wend

    ' Close .ini file
    uniFile.CloseFile
  End If

  ' Setup enable spectrum smoothing selection
  If (m_enableSmooth = False) Then
    chk_enableSmooth.Value = 0
    frame_smooth.enabled = False
  Else
    chk_enableSmooth.Value = 1
    frame_smooth.enabled = True
  End If
  
  ' Check ending progressive smooth wavelength
  If (m_endSmoothWvln > ProdDfltData.endWvln) Or (m_endSmoothWvln <= m_startSmoothWvln) Then
    unity_main.errorstring = (fileName & " had incompatible value. EndSmoothWvln was " & m_endSmoothWvln & "; updated to " & ProdDfltData.endWvln)
    unity_main.write_error
    m_endSmoothWvln = ProdDfltData.endWvln
    m_badIniVal = True
  End If
  
  ' Setup ending progressive smooth wavelength selection
  txt_endSmoothWvln.Text = m_endSmoothWvln
  
  ' Check ending progressive smooth number of end-sum pts
  If (m_endSmoothNumPts < numInc_endSmoothNumPts.Min) Or (m_endSmoothNumPts > numInc_endSmoothNumPts.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. EndSmoothNumPts was " & m_endSmoothNumPts & "; updated to " & DEFLT_NUM_PTS)
    unity_main.write_error
    m_endSmoothNumPts = DEFLT_NUM_PTS
    m_badIniVal = True
  End If
  
  ' Setup ending progressive smooth number of end-sum pts selection
  numInc_endSmoothNumPts.Text = m_endSmoothNumPts
  
  ' Check progressive smooth rate
  If (m_progSmoothRate < numInc_progSmoothRate.Min) Or (m_progSmoothRate > numInc_progSmoothRate.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. ProgSmoothRate was " & m_progSmoothRate & "; updated to " & DEFLT_PROG_RATE)
    unity_main.write_error
    m_progSmoothRate = DEFLT_PROG_RATE
    m_badIniVal = True
  End If
  
  ' Setup progressive smooth rate selection
  numInc_progSmoothRate.Text = m_progSmoothRate
  
  ' Setup save spectrum selection
PROCESS_SAVESPECTRA:
  Select Case (m_saveSpectra)
    Case "Treated"
      opt_saveTreated.Value = True
    Case "Both"
      opt_saveBoth.Value = True
    Case Else
      unity_main.errorstring = (fileName & " had incompatible value. SaveSpectra was " & m_saveSpectra & "; updated to " & DEFLT_SAVE_SPECTRA)
      unity_main.write_error
      m_saveSpectra = DEFLT_SAVE_SPECTRA
      m_badIniVal = True
      GoTo PROCESS_SAVESPECTRA
  End Select
  
  ' Check smooth number of end-sum pts
  If (m_smoothNumPts < numInc_smoothNumPts.Min) Or (m_smoothNumPts > numInc_smoothNumPts.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. SmoothNumPts was " & m_smoothNumPts & "; updated to " & DEFLT_NUM_PTS)
    unity_main.write_error
    m_smoothNumPts = DEFLT_NUM_PTS
    m_badIniVal = True
  End If
  
  ' Setup smooth number of end-sum pts selection
  numInc_smoothNumPts.Text = m_smoothNumPts
  
  ' Setup smooth type selection
PROCESS_SMOOTHTYPE:
  Select Case (m_smoothType)
    Case "BoxCar"
      m_smoothTypeEnum = ST_BOX_CAR
      combo_smoothType.ListIndex = 0
    Case "Triangular"
      m_smoothTypeEnum = ST_TRIANGULAR
      combo_smoothType.ListIndex = 1
    Case "SavitskyGolay"
      m_smoothTypeEnum = ST_SAVITSKY_GOLAY
      combo_smoothType.ListIndex = 2
    Case Else
      unity_main.errorstring = (fileName & " had incompatible value. SmoothType was " & m_smoothType & "; updated to BoxCar")
      unity_main.write_error
      m_smoothType = "BoxCar"
      m_badIniVal = True
      GoTo PROCESS_SMOOTHTYPE
  End Select
  
  ' Check starting progressive smooth wavelength
  If (m_startSmoothWvln < ProdDfltData.startWvln) Or (m_startSmoothWvln >= m_endSmoothWvln) Then
    unity_main.errorstring = (fileName & " had incompatible value. StartSmoothWvln was " & m_startSmoothWvln & "; updated to " & ProdDfltData.startWvln)
    unity_main.write_error
    m_startSmoothWvln = ProdDfltData.startWvln
    m_badIniVal = True
  End If
  
  ' Setup starting progressive smooth wavelength selection
  txt_startSmoothWvln.Text = m_startSmoothWvln
  
  ' Check starting progressive smooth number of end-sum pts
  If (m_startSmoothNumPts < numInc_startSmoothNumPts.Min) Or (m_startSmoothNumPts > numInc_startSmoothNumPts.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. StartSmoothNumPts was " & m_startSmoothNumPts & "; updated to " & DEFLT_NUM_PTS)
    unity_main.write_error
    m_startSmoothNumPts = DEFLT_NUM_PTS
    m_badIniVal = True
  End If
  
  ' Setup starting progressive smooth number of end-sum pts selection
  numInc_startSmoothNumPts.Text = m_startSmoothNumPts
  
  ' Setup use progressive smooth algorithm selection
  If (m_useProgSmooth = False) Then
    chk_useProgSmooth.Value = 0
  Else
    chk_useProgSmooth.Value = 1
  End If
  
  ' Display smooth algorithm parameters
  disp_smooth_parameters
  
  ' Copy loaded spectrum treatment config into system operational variables
  unity_main.m_enableTreatment = m_enableSmooth   ' OR all treatment options
  unity_main.m_enableSmooth = m_enableSmooth
  unity_main.m_endSmoothWvln = m_endSmoothWvln
  unity_main.m_endSmoothNumPts = m_endSmoothNumPts
  unity_main.m_progSmoothRate = m_progSmoothRate
  unity_main.m_saveSpectra = m_saveSpectra
  unity_main.m_smoothNumPts = m_smoothNumPts
  unity_main.m_smoothTypeEnum = m_smoothTypeEnum
  unity_main.m_smoothType = m_smoothType
  unity_main.m_startSmoothWvln = m_startSmoothWvln
  unity_main.m_startSmoothNumPts = m_startSmoothNumPts
  unity_main.m_useProgSmooth = m_useProgSmooth
  
  ' Check if ini file had bad value
  If (m_badIniVal = True) Then
    unity_main.errorstring = (fileName & " had incompatible value(s). Updated file with proper/default value(s).")
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg7", "%1 had incompatible value(s). Updated file with proper/default value(s)", fileName)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg3", "%1. Refer to system log for more details", uniMsg), vbOKOnly
    Call save_cfg
  Else
    ' Check if to create configuration file for first time
    If (exist = False) Then
      save_cfg
    End If
  End If
  
  Exit Sub
  
FILE_ERROR:
  If (lineCnt = 0) Then
    errMsg = (fileName & " file open error." & Error$)
    uniMsg = MLSupport.GGS_Params("fileErrMsg2", "%1 file open error. %2", fileName, Error$)
  Else
    errMsg = (fileName & " file has error on line " & CStr(lineCnt) & ". " & Error$)
    uniMsg = MLSupport.GGS_Params("fileErrMsg3", "%1 file has error on line %2. %3", fileName, CStr(lineCnt), Error$)
  End If
  
  unity_main.errorstring = errMsg
  unity_main.write_error
  
LOAD_ERROR:
  uniMsg = MLSupport.GGS_Params("frm_spectTreatCfg.errMsg1", "%1. Using default smooth algorithm values", uniMsg)
  CWrap.ShowMessageBoxW uniMsg, vbCritical
  uniFile.CloseFile
End Sub

Private Function load_cfg_settings(ByVal fileName As String, ByVal uniFile As clsUniFile, ByVal fEncoding As eFileEncoding, ByRef lineCnt As Integer) As Boolean
  Dim iniString As String
  Dim pos As Integer
  Dim tmpStrg As String
  Dim cfgVar As String
  Dim varVal As String
  Dim strlen As Integer
  Dim rc As Boolean

  ' Process each remaining line in .ini file
  While Not (uniFile.EOF())
    On Error GoTo FILE_ERROR
    ' Read line from file
    lineCnt = lineCnt + 1
    
    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(iniString)
    Else
      rc = uniFile.ReadUnicodeLine(iniString)
    End If
      
    If (rc = False) Then GoTo FILE_ERROR
          
    ' Get variable name and its value
    pos = InStr(1, iniString, "=")
    strlen = Len(iniString)
    tmpStrg = Trim(Mid(iniString, 1, pos - 1))
    cfgVar = LCase(tmpStrg)
    varVal = Trim(Mid(iniString, pos + 1))
    
    ' Process value by variable name
    On Error GoTo BAD_INI_VALUE
    Select Case (cfgVar)
      Case "enablesmooth"
        m_enableSmooth = CBool(varVal)
      Case "endsmoothwvln"
        m_endSmoothWvln = CDbl(varVal)
      Case "endsmoothnumpts"
        m_endSmoothNumPts = CInt(varVal)
      Case "progsmoothrate"
        m_progSmoothRate = CInt(varVal)
      Case "savespectra"
        m_saveSpectra = varVal
      Case "smoothnumpts"
        m_smoothNumPts = CInt(varVal)
      Case "smoothtype"
        m_smoothType = varVal
      Case "startsmoothwvln"
        m_startSmoothWvln = CDbl(varVal)
      Case "startsmoothnumpts"
        m_startSmoothNumPts = CInt(varVal)
      Case "useprogsmooth"
        m_useProgSmooth = CBool(varVal)
    End Select
  Wend
  
  load_cfg_settings = True
  Exit Function
  
BAD_INI_VALUE:
    unity_main.errorstring = (fileName & " had incompatible value. " & cfgVar & " = " & varVal & "; will use default value")
    unity_main.write_error
    m_badIniVal = True
    Resume Next
  
FILE_ERROR:
  load_cfg_settings = False
End Function

Private Function check_cfg_settings() As Boolean
  Dim startWvln As Double
  Dim endWvln As Double
  Dim rc As Boolean
  Dim userReq As Integer

  check_cfg_settings = False

  ' See if progressive algorithm has been selected
  If (chk_useProgSmooth.Value = 1) Then
    txt_startSmoothWvln.Text = Trim(txt_startSmoothWvln.Text)
    txt_endSmoothWvln.Text = Trim(txt_endSmoothWvln.Text)

    On Error GoTo BAD_VALUE
    startWvln = CDbl(txt_startSmoothWvln.Text)
    endWvln = CDbl(txt_endSmoothWvln.Text)

    If (startWvln < ProdDfltData.startWvln) Or (startWvln > ProdDfltData.endWvln) Then
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_spectTreatCfg.errMsg2", "Please enter a Starting Wavelength value between %1 and %2", CStr(ProdDfltData.startWvln), CStr(ProdDfltData.endWvln)), vbExclamation
      frm_spectTreatCfg.txt_startSmoothWvln.SetFocus
      Exit Function
    Else
      If (endWvln < ProdDfltData.startWvln) Or (endWvln > ProdDfltData.endWvln) Then
        CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_spectTreatCfg.errMsg3", "Please enter a Ending Wavelength value between %1 and %2", CStr(ProdDfltData.startWvln), CStr(ProdDfltData.endWvln)), vbExclamation
        frm_spectTreatCfg.txt_endSmoothWvln.SetFocus
        Exit Function
      Else
        If (endWvln <= startWvln) Then
          CWrap.ShowMessageBoxW MLSupport.GSS("frm_spectTreatCfg", "errMsg1", "Please enter a Ending Wavelength value greater than the Starting Wavelength value"), vbExclamation
          frm_spectTreatCfg.txt_endSmoothWvln.SetFocus
          Exit Function
        Else
          ' Check wavelengths with instrument's table
#If SSRCS Then
          SSRCSClientError = unity_main.SSRCSClient.ChkWvlnRange(startWvln, endWvln, rc)
#Else
          rc = unity_main.MS11srv.ChkWvlnRange(startWvln, endWvln)
#End If
            
          ' Check if wavelengths changed by instrument, if so ask user if okay
          If (rc = False) Then
            userReq = CWrap.ShowMessageBoxW(MLSupport.GGS_Params("frm_spectTreatCfg.statMsg1", "Starting/Ending Wavelengths were changed to %1 - %2. Do you want use these values?", CStr(startWvln), CStr(endWvln)), vbYesNo)
            
            If (userReq = vbNo) Then
              Exit Function
            Else
              txt_startSmoothWvln.Text = startWvln
              txt_endSmoothWvln.Text = endWvln
            End If
          End If
        End If
      End If
    End If
  End If
  
  check_cfg_settings = True
  Exit Function
  
BAD_VALUE:
  CWrap.ShowMessageBoxW MLSupport.GSS("frm_spectTreatCfg", "errMsg2", "Please enter a valid number for Starting/Ending Wavelength"), vbOKOnly
  frm_spectTreatCfg.txt_startSmoothWvln.SetFocus
End Function

Private Sub disp_smooth_parameters()

  ' See if progressive algorithm has been selected
  If (chk_useProgSmooth.Value = 1) Then
    ' Display progressive parameters
    numInc_startSmoothNumPts.Visible = True
    lbl_startSmoothNumPts.Visible = True
    numInc_endSmoothNumPts.Visible = True
    lbl_endSmoothNumPts.Visible = True
    numInc_progSmoothRate.Visible = True
    lbl_progSmoothRate.Visible = True
    txt_startSmoothWvln.Visible = True
    lbl_startSmoothWvln.Visible = True
    txt_endSmoothWvln.Visible = True
    lbl_endSmoothWvln.Visible = True
    
    ' Hide nonprogressive parameters
    numInc_smoothNumPts.Visible = False
    lbl_smoothNumPts.Visible = False
  Else
    ' Display nonprogressive parameters
    numInc_smoothNumPts.Visible = True
    lbl_smoothNumPts.Visible = True
    
    ' Hide progressive parameters
    numInc_startSmoothNumPts.Visible = False
    lbl_startSmoothNumPts.Visible = False
    numInc_endSmoothNumPts.Visible = False
    lbl_endSmoothNumPts.Visible = False
    numInc_progSmoothRate.Visible = False
    lbl_progSmoothRate.Visible = False
    txt_startSmoothWvln.Visible = False
    lbl_startSmoothWvln.Visible = False
    txt_endSmoothWvln.Visible = False
    lbl_endSmoothWvln.Visible = False
  End If
End Sub

Private Sub chk_enableSmooth_Click()

  If (chk_enableSmooth.Value = 1) Then
    frame_smooth.enabled = True
  Else
    frame_smooth.enabled = False
  End If
End Sub

Private Sub chk_useProgSmooth_Click()

  disp_smooth_parameters
End Sub

Private Sub cmd_cancel_Click()
  
  unity_main.errorstring = "Spectrum Treatment Configuration screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Unload frm_spectTreatCfg
End Sub

Private Sub cmd_save_Click()
  
  unity_main.errorstring = "Spectrum Treatment Configuration screen Save Changes button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  ' Check treatment parameters are configured properly before saving to file
  If (check_cfg_settings = True) Then
    Call save_cfg
    Unload frm_spectTreatCfg
  
    unity_main.errorstring = ("User saved new settings for configuration file: " & CFG_DIR & SPECT_TREAT_CFG_FILE)
    unity_main.write_error (LOG_DBG_LEVEL1)
  End If
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
  ' Build spectrum smoothing types list
  combo_smoothType.AddItem MLSupport.GSS("frm_spectTreatCfg", "smoothType1", "Box Car")
  combo_smoothType.AddItem MLSupport.GSS("frm_spectTreatCfg", "smoothType2", "Triangular")
  combo_smoothType.AddItem MLSupport.GSS("frm_spectTreatCfg", "smoothType3", "Savitsky Golay")
End Sub

Private Sub numInc_endSmoothNumPts_dblclick()
  
  unity_main.formfrom = 19
  unity_main.varfrom = 3
  frm_numpad.lbl_num.Caption = lbl_endSmoothNumPts.Caption
  frm_numpad.txt_num.Text = numInc_endSmoothNumPts.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_progSmoothRate_dblclick()

  unity_main.formfrom = 19
  unity_main.varfrom = 4
  frm_numpad.lbl_num.Caption = lbl_progSmoothRate.Caption
  frm_numpad.txt_num.Text = numInc_progSmoothRate.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_smoothNumPts_dblclick()
  
  unity_main.formfrom = 19
  unity_main.varfrom = 1
  frm_numpad.lbl_num.Caption = lbl_smoothNumPts.Caption
  frm_numpad.txt_num.Text = numInc_smoothNumPts.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_startSmoothNumPts_dblclick()
  
  unity_main.formfrom = 19
  unity_main.varfrom = 2
  frm_numpad.lbl_num.Caption = lbl_startSmoothNumPts.Caption
  frm_numpad.txt_num.Text = numInc_startSmoothNumPts.Text
  frm_numpad.Show 1
End Sub

Private Sub txt_endSmoothWvln_DblCLick(Button As Integer)
  
  unity_main.formfrom = 19
  unity_main.varfrom = 6
  frm_numpad.lbl_num.Caption = lbl_endSmoothWvln.Caption
  frm_numpad.txt_num.Text = txt_endSmoothWvln.Text
  frm_numpad.Show 1
End Sub

Private Sub txt_startSmoothWvln_DblCLick(Button As Integer)
  
  unity_main.formfrom = 19
  unity_main.varfrom = 5
  frm_numpad.lbl_num.Caption = lbl_startSmoothWvln.Caption
  frm_numpad.txt_num.Text = txt_startSmoothWvln.Text
  frm_numpad.Show 1
End Sub






