VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{E6AC3E35-BC5B-44AC-B1A0-251A8A08AD90}#17.0#0"; "XYPlot.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Object = "{4411935E-3987-4955-879C-C5272EC64407}#1.0#0"; "SpxFileIF.ocx"
Begin VB.Form frm_intRefCalMgmt 
   Caption         =   "Internal Reference Calibration Management"
   ClientHeight    =   8040
   ClientLeft      =   2055
   ClientTop       =   3255
   ClientWidth     =   11190
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
   Icon            =   "frm_intRefCalMgmt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr_scan 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3360
      Top             =   7080
   End
   Begin HexUniControls.ctlUniListBoxXP lst_spxHistory 
      Height          =   1935
      Left            =   3720
      TabIndex        =   6
      Top             =   4800
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3413
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TrapTab         =   0   'False
      Tip             =   "frm_intRefCalMgmt.frx":0442
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_intRefCalMgmt.frx":0462
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin HexUniControls.ctlUniLabel lbl_newScanTime 
      Height          =   255
      Left            =   2160
      Top             =   240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_intRefCalMgmt.frx":047E
      BackColor       =   16777215
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_intRefCalMgmt.frx":049E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_intRefCalMgmt.frx":04BE
   End
   Begin HexUniControls.ctlUniLabel lbl_currScanTime 
      Height          =   255
      Left            =   120
      Top             =   240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_intRefCalMgmt.frx":04DA
      BackColor       =   16777215
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_intRefCalMgmt.frx":04FA
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_intRefCalMgmt.frx":051A
   End
   Begin HexUniControls.ctlUniFrameXP frame_ref 
      Height          =   1935
      Left            =   120
      Top             =   4800
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3413
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_intRefCalMgmt.frx":0536
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_intRefCalMgmt.frx":0572
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_intRefCalMgmt.frx":0592
      Begin HexUniControls.ctlUniLabel lbl_numRefScans 
         Height          =   345
         Left            =   2160
         Top             =   600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_intRefCalMgmt.frx":05AE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   1
         AutoSize        =   0   'False
         Tip             =   "frm_intRefCalMgmt.frx":05CE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_intRefCalMgmt.frx":05EE
      End
      Begin HexUniControls.ctlUniLabel lbl_numScans1 
         Height          =   345
         Left            =   120
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_intRefCalMgmt.frx":060A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_intRefCalMgmt.frx":0648
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_intRefCalMgmt.frx":0668
      End
      Begin HexUniControls.ctlUniLabel lbl_refScanProg 
         Height          =   345
         Left            =   120
         Top             =   1080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_intRefCalMgmt.frx":0684
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_intRefCalMgmt.frx":06BE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_intRefCalMgmt.frx":06DE
      End
      Begin HexUniControls.ctlProgressXP prg_refScan 
         Height          =   375
         Left            =   120
         Top             =   1440
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         BackColor       =   16777215
         ForeColor       =   49152
         Border          =   -1  'True
         BorderSpace     =   -1  'True
         Spaces          =   -1  'True
         Tip             =   "frm_intRefCalMgmt.frx":06FA
         Style           =   -1
         BackStyle       =   -1
         RoundedBorders  =   -1  'True
         MousePointer    =   0
         MouseIcon       =   "frm_intRefCalMgmt.frx":071A
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_startScan 
      Height          =   705
      Left            =   720
      TabIndex        =   2
      Top             =   6960
      Width           =   2205
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
      Caption         =   "frm_intRefCalMgmt.frx":0736
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_intRefCalMgmt.frx":076A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_intRefCalMgmt.frx":078A
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   8040
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8040
      FormDesignWidth =   11190
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_exit 
      Cancel          =   -1  'True
      Height          =   705
      Left            =   8160
      TabIndex        =   1
      Top             =   6960
      Width           =   2205
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
      Caption         =   "frm_intRefCalMgmt.frx":07A6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_intRefCalMgmt.frx":07CE
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_intRefCalMgmt.frx":07EE
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_abortScan 
      Height          =   705
      Left            =   4440
      TabIndex        =   0
      Top             =   6960
      Width           =   2205
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
      Caption         =   "frm_intRefCalMgmt.frx":080A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_intRefCalMgmt.frx":083E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_intRefCalMgmt.frx":085E
   End
   Begin HexUniControls.ctlUniFrameXP frame_smpl 
      Height          =   1935
      Left            =   7560
      Top             =   4800
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3413
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_intRefCalMgmt.frx":087A
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_intRefCalMgmt.frx":08B0
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_intRefCalMgmt.frx":08D0
      Begin HexUniControls.ctlUniLabel lbl_numScans2 
         Height          =   345
         Left            =   120
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_intRefCalMgmt.frx":08EC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_intRefCalMgmt.frx":092A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_intRefCalMgmt.frx":094A
      End
      Begin HexUniControls.ctlUniLabel lbl_smplScanProg 
         Height          =   345
         Left            =   120
         Top             =   1080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_intRefCalMgmt.frx":0966
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_intRefCalMgmt.frx":09A0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_intRefCalMgmt.frx":09C0
      End
      Begin HexUniControls.ctlProgressXP prg_smplScan 
         Height          =   375
         Left            =   120
         Top             =   1440
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         BackColor       =   16777215
         ForeColor       =   49152
         Border          =   -1  'True
         BorderSpace     =   -1  'True
         Spaces          =   -1  'True
         Tip             =   "frm_intRefCalMgmt.frx":09DC
         Style           =   -1
         BackStyle       =   -1
         RoundedBorders  =   -1  'True
         MousePointer    =   0
         MouseIcon       =   "frm_intRefCalMgmt.frx":09FC
      End
      Begin HexUniControls.ctlUniLabel lbl_numSmplScans 
         Height          =   345
         Left            =   2280
         Top             =   600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_intRefCalMgmt.frx":0A18
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   1
         AutoSize        =   0   'False
         Tip             =   "frm_intRefCalMgmt.frx":0A38
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_intRefCalMgmt.frx":0A58
      End
   End
   Begin HexUniControls.ctlUniFrameXP frame_ss 
      Height          =   1935
      Left            =   3720
      Top             =   4800
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3413
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_intRefCalMgmt.frx":0A74
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_intRefCalMgmt.frx":0AAA
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_intRefCalMgmt.frx":0ACA
      Begin HexUniControls.ctlUniLabel lbl_serNumber 
         Height          =   345
         Left            =   120
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_intRefCalMgmt.frx":0AE6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_intRefCalMgmt.frx":0B20
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_intRefCalMgmt.frx":0B40
      End
      Begin HexUniControls.ctlUniLabel lbl_minWavelen 
         Height          =   345
         Left            =   120
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_intRefCalMgmt.frx":0B5C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_intRefCalMgmt.frx":0BA0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_intRefCalMgmt.frx":0BC0
      End
      Begin HexUniControls.ctlUniLabel lbl_maxWavelen 
         Height          =   345
         Left            =   120
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_intRefCalMgmt.frx":0BDC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_intRefCalMgmt.frx":0C20
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_intRefCalMgmt.frx":0C40
      End
      Begin HexUniControls.ctlUniLabel lbl_spxCorrFilename 
         Height          =   345
         Left            =   120
         Top             =   1440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_intRefCalMgmt.frx":0C5C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_intRefCalMgmt.frx":0CAA
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_intRefCalMgmt.frx":0CCA
      End
      Begin HexUniControls.ctlUniLabel lbl_serNum 
         Height          =   345
         Left            =   2280
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_intRefCalMgmt.frx":0CE6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   1
         AutoSize        =   0   'False
         Tip             =   "frm_intRefCalMgmt.frx":0D06
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_intRefCalMgmt.frx":0D26
      End
      Begin HexUniControls.ctlUniLabel lbl_minWvln 
         Height          =   345
         Left            =   2280
         Top             =   720
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_intRefCalMgmt.frx":0D42
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   1
         AutoSize        =   0   'False
         Tip             =   "frm_intRefCalMgmt.frx":0D62
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_intRefCalMgmt.frx":0D82
      End
      Begin HexUniControls.ctlUniLabel lbl_maxWvln 
         Height          =   345
         Left            =   2280
         Top             =   1080
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_intRefCalMgmt.frx":0D9E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   1
         AutoSize        =   0   'False
         Tip             =   "frm_intRefCalMgmt.frx":0DBE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_intRefCalMgmt.frx":0DDE
      End
      Begin HexUniControls.ctlUniLabel lbl_spxCorrFile 
         Height          =   345
         Left            =   2280
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_intRefCalMgmt.frx":0DFA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   1
         AutoSize        =   0   'False
         Tip             =   "frm_intRefCalMgmt.frx":0E1A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_intRefCalMgmt.frx":0E3A
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_updateCalib 
      Height          =   705
      Left            =   4440
      TabIndex        =   4
      Top             =   6960
      Width           =   2205
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
      Caption         =   "frm_intRefCalMgmt.frx":0E56
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_intRefCalMgmt.frx":0E9A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_intRefCalMgmt.frx":0EBA
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_restoreCalib 
      Height          =   705
      Left            =   4440
      TabIndex        =   5
      Top             =   6960
      Width           =   2205
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
      Caption         =   "frm_intRefCalMgmt.frx":0ED6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_intRefCalMgmt.frx":0F1C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_intRefCalMgmt.frx":0F3C
   End
   Begin XYPlotGraph.XYPlot xyPlot 
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7858
   End
   Begin SPXFILEIFLib.SpxFileIF SpxFileIF1 
      Left            =   9960
      Top             =   360
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   1296
      _StockProps     =   0
   End
   Begin HexUniControls.ctlUTF8Menu ctlUTF8Menu 
      Left            =   120
      Top             =   6840
      _ExtentX        =   794
      _ExtentY        =   794
      IconDim         =   16
      FadeBackground  =   0   'False
      MenuBackColor   =   -2147483644
      MenuForeColorNormal=   -2147483641
      MenuForeColorDisabled=   -2147483631
      SelectorBackColor=   -2147483635
      SelectorForeColor=   -2147483634
      ExtSepBackColor =   -2147483644
      ExtSepForeColor =   -2147483641
      FadeExtSep      =   0   'False
      SelectorStyle   =   -1
      FullSelector    =   -1  'True
      LeftBand        =   0   'False
      LeftBandBackColor=   -2147483644
      FadeLeftBand    =   0   'False
      Separators3D    =   -1  'True
      LeftBandSeparators=   0   'False
      StringMode      =   2
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   1080
      Top             =   8040
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
      Top             =   8040
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_intRefCalMgmt.frx":0F58
   End
   Begin VB.Menu CheckMenu 
      Caption         =   "Check Reference"
   End
   Begin VB.Menu RecalibrateMenu 
      Caption         =   "Recalibrate Reference"
   End
   Begin VB.Menu HistoryMenu 
      Caption         =   "Calibration History"
   End
End
Attribute VB_Name = "frm_intRefCalMgmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Number of sample scans
Const NUM_CHECK_SCANS = 24              ' check reference
Const NUM_RECAL_SCANS = 100           ' recalibrate reference

' Reference expiration time in minutes
Const REF_EXPIRE_TIME = (60 * 15)

' Default check reference limits
Const CHK_MAX_REFLECT_VAL_LMT = 1.1     ' maximum reflectance value
Const CHK_MIN_REFLECT_VAL_LMT = 0.9     ' minimum reflectance value
Const CHK_MAX_DEV_VAL_LMT = 0.05        ' maximum reflectance deviation value
Const CHK_TARGET_VAL = 1#               ' target reflectance value

' Allow Changes to ScanMode in the SCM_SCNSYSB Bits
Const MSC_OVRDSCNM = &H10

' 2500 Math Treatment Operation File Names
Const MATH_TREAT_2500_BAT_FILE_NAME = "math2500.bat"  ' 2500 math treatment batch file
Const SPX_2500_CAL_FILE_NAME = "cal2500.dat"          ' 2500 sampled calibration spectra file
Const SPX_2500_RECAL_FILE_NAME = "recal.dat"          ' 2500 math treated recalibrated spectra file

' Spectra Correction File Interface Error Codes
Enum SPX_CORR_FILE_ERRORS
  SPX_FILE_GOOD = 0
  SPX_FILE_NULL_ARG_ERR           ' nulled argument passed
  SPX_CORR_FILE_NOT_LOAD_ERR      ' correction file not loaded
  SPX_CORR_FILE_ERR               ' correction file error (details stored in m_spxLastErrorCode)
  SPX_CORR_FILE_INV_WVLN_ERR      ' correction file has wavelength parameters mismatch
  SPX_CORR_FILE_INV_RECS_ERR      ' correction file has improper number of correction spectra records for system type
  SPX_HIST_FILE_NOT_LOAD_ERR      ' history correction file not loaded
  SPX_HIST_FILE_ERR               ' history correction file error (details stored in m_spxLastErrorCode)
  SPX_HIST_FILE_INV_WVLN_ERR      ' history correction file has wavelength parameters mismatch
  SPX_HIST_FILE_INV_REC_INDX_ERR  ' invalid record index for history correction file
  SPX_BAT_FILE_CREATE_ERR         ' error creating 2500 math treatment batch file
  SPX_BAT_FILE_EXE_ERR            ' error excuting 2500 math treatment batch file (details stored in m_spxLastErrorCode)
End Enum

' Scanning States
Enum CAL_SCAN_STATES
  CSS_START = 1               ' start scan
  CSS_CHK_COMP                ' check if scan completed
  CSS_GET_SPECTRUM            ' get scan spectrum
  CSS_TREAT_SPECTRUM          ' apply 2500 math treatment to spectrum
  CSS_PLOT_SPECTRUM           ' plot scan spectrum
  CSS_DONE                    ' scanning completed
  CSS_ERROR                   ' scan error
  CSS_ABORT                   ' scan aborted
  CSS_SHUTDOWN                ' major scan operation error, shutdown
End Enum

' Scan Completion Codes
Enum SCAN_COMPLETE_CODES
  SC_WAITING = 0             ' waiting for scan completion
  SC_GOOD                    ' scan completed
  SC_ERROR                   ' scan completed with error
  SC_ABORT                   ' scan aborted
  SC_SHUTDOWN                ' major scan operation error, shutdown
End Enum

' Spectra Correction File Interface "time_t2" data type
Private Type TimeT2
  msecs As Long
  mins As Long
End Type

Public m_restartFlg As Boolean

' System
Private m_badIniVal As Boolean
Private m_calibFunc As CALIB_FUNC
Private m_chkRefState As Integer
Private m_errMsg As String
Private m_ms11SrvCtlb As Long
Private m_numPts As Integer
Private m_scanAbortFlg As Boolean
Private m_scanEndWvln As Double
Private m_scanFlg As Boolean
Private m_scanMode As Long
Private m_scanStartWvln As Double
Private m_scanState As CAL_SCAN_STATES
Private m_spxCorFile As String
Private m_uniErrMsg As String

' Internal reference scan
Private m_refCalTakenFlg As Boolean
Private m_refCalUpdatedFlg As Boolean
Private m_refChkTaken As Boolean
Private m_refChkTimeStamp As Single
Private m_refFlag As Boolean
Private m_refNumScans As Integer
Private m_refPPT As Integer
Private m_refUpdTaken As Boolean
Private m_refUpdTimeStamp As Single
Private m_refUpdTimeT2 As TimeT2
 
  ' Sample scan
Private m_smplChkTimeT2 As TimeT2
Private m_smplNum As Integer
Private m_smplNumReferScans As Integer
Private m_smplNum2ndOrderScans As Integer
Private m_smplReady2Save As Boolean
Private m_smplRotateIndexSteps As Integer
Private m_smplRotateDir As TRAY_ROTATE_DIRS
Private m_smplRotateMoveMode As TRAY_ROTATE_MOVEMENTS
Private m_smplRotateSpeed As Integer
Private m_smplRotateStepSteps As Integer
Private m_smplTrayNum As Integer
Private m_smplReferTimeT2 As TimeT2
Private m_smpl2ndOrderTimeT2 As TimeT2

' Spectra correction file
Private m_corrNumRecs As Integer
Private m_corrReferTimeT2 As TimeT2
Private m_corr2ndOrderTimeT2 As TimeT2

' History spectra correction file
Private m_histNumRecs As Integer
Private m_histNumReferScans As Integer
Private m_histNum2ndOrderScans As Integer
Private m_histReferTimeT2 As TimeT2
Private m_hist2ndOrderTimeT2 As TimeT2

' Spectra data
Private m_corrReferVals() As Double
Private m_corr2ndOrderVals() As Double
Private m_histReferVals() As Double
Private m_hist2ndOrderVals() As Double
Private m_refIntenVals() As Double
Private m_smplChkReferVals() As Double
Private m_smplChk2ndOrderVals() As Double
Private m_smplIntenVals() As Double
Private m_smplUpdReferVals() As Double
Private m_smplUpd2ndOrderVals() As Double
Private m_wvlnVals() As Double

' Check reference limits
Private m_chkEndWvlnLmt As Double
Private m_chkMaxDevVal As Double
Private m_chkMaxDevValLmt As Double
Private m_chkMaxReflectVal As Double
Private m_chkMaxReflectValLmt As Double
Private m_chkMinReflectVal As Double
Private m_chkMinReflectValLmt As Double
Private m_chkStartWvlnLmt As Double

' Check reference limits plotting
Private m_chkStartIndx As Integer
Private m_chkEndIndx As Integer
Private m_chkMaxReflectLmts() As Double
Private m_chkMinReflectLmts() As Double
Private m_chkTargets() As Double

' Backup of current internal reference scan parameters
Private m_intRefNScans As Integer
Private m_intRefPPT As Integer

Public Sub update_scan_progress(ByVal percent As Long)

  If (m_refFlag = True) Then
    prg_refScan.percent = percent
  Else
    prg_smplScan.percent = percent
  End If
End Sub

Public Sub scan_aborted()
           
  m_scanAbortFlg = True

  ' Check if aborted reference scan
  If (m_refFlag = True) Then
    unity_main.tmr_ref.enabled = False
    m_uniErrMsg = MLSupport.GSS("frm_intRefCalMgmt", "statMsg1", "Reference Scan Aborted")
  Else
    unity_main.tmr_sample.enabled = False
    m_uniErrMsg = MLSupport.GSS("frm_intRefCalMgmt", "statMsg2", "Sample Scan Aborted")
  End If
  
  m_scanFlg = False
  m_scanState = CSS_ABORT
End Sub

Private Function setup_spx_if() As Boolean


  ' Check scan mode for proper setting
  If (unity_main.m_sys2500 = True) Then
    If (MS11CfgData.sysScanMode <> &H1C06) Then
      m_errMsg = "Scan Mode configured incorrectly for instrument " & Hex(MS11CfgData.sysScanMode) & "; expected '1C06'"
      m_uniErrMsg = MLSupport.GGS_Params("frm_intRefCalMgmt.errMsg1", "Scan Mode configured incorrectly for instrument %1; expected '1C06'", Hex(MS11CfgData.sysScanMode))
      GoTo SetupError
    End If
  Else
    If (MS11CfgData.sysScanMode <> &HC06) Then
      m_errMsg = "Scan Mode configured incorrectly for instrument " & Hex(MS11CfgData.sysScanMode) & "; expected '0C06'"
      m_uniErrMsg = MLSupport.GGS_Params("frm_intRefCalMgmt.errMsg2", "Scan Mode configured incorrectly for instrument %1; expected '0C06'", Hex(MS11CfgData.sysScanMode))
      GoTo SetupError
    End If
  End If
  
  ' Initialze scan and sample plot
  init_scan_info
  init_plot
  setup_plot
  
  ' Check if any error loading in spectra correction file
  If (load_correction_file() = False) Then
    GoTo SetupError
  End If

  ' Check if any error loading in spectra history file
  If (load_history_file() = False) Then
    GoTo SetupError
  End If

  ' Load in check reference limits file
  load_chk_ref_lmts_file

  ' Save and modify MS11srv control bits to change scan mode bits
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.GetMS11srvCtlb(m_ms11SrvCtlb)
  
  If (SSRCSClientError = 0) Then
    SSRCSClientError = unity_main.SSRCSClient.SetMS11srvCtlb(m_ms11SrvCtlb Or MSC_OVRDSCNM)
    
    If (SSRCSClientError = 0) Then
      GoTo SETUP_CMPL
    End If
  End If
  
  m_errMsg = "Cannot access SpectraStar control bits"
  m_uniErrMsg = MLSupport.GSS("MS11srv", "errMsg10", "Cannot access SpectraStar control bits")
  GoTo SetupError
#Else
  m_ms11SrvCtlb = unity_main.MS11srv.MS11srvCtlb
  unity_main.MS11srv.MS11srvCtlb = m_ms11SrvCtlb Or MSC_OVRDSCNM
#End If

SETUP_CMPL:
  setup_spx_if = True
  Exit Function
  
SetupError:
  unity_main.errorstring = m_errMsg
  unity_main.write_error
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", m_uniErrMsg), vbCritical
End Function

Private Function load_correction_file() As Boolean
  Dim dirname As String
  Dim rc As Boolean
  
  ' Check if no spectral correction file configured
  If (m_spxCorFile = "") Then
    m_errMsg = "No spectra correction file configured"
    m_uniErrMsg = MLSupport.GSS("frm_intRefCalMgmt", "errMsg1", "No spectra correction file configured")
    Exit Function
  End If

  dirname = MS11_DIR

  ' Load correction file and get spectra data
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.LoadSpxCorrectionFile(dirname, m_spxCorFile, m_scanStartWvln, m_scanEndWvln, MS11CfgData.wvlnIncr, unity_main.m_sys2500, m_corrNumRecs)
  
  If (SSRCSClientError = 0) Then
    SSRCSClientError = unity_main.SSRCSClient.GetSpxCorrectionSpectra(unity_main.m_sys2500, m_numPts, m_corrReferVals(0), VarPtr(m_corrReferTimeT2), m_corr2ndOrderVals(0), VarPtr(m_corr2ndOrderTimeT2))
    
    If (SSRCSClientError = 0) Then
      rc = True
    End If
  End If

#Else
'MsgBox ("correct file " & m_spxCorFile)

  If (SpxFileIF1.LoadCorrectionFile(dirname, m_spxCorFile, m_scanStartWvln, m_scanEndWvln, MS11CfgData.wvlnIncr, unity_main.m_sys2500, m_corrNumRecs) = True) Then
    rc = SpxFileIF1.GetCorrectionSpectra(unity_main.m_sys2500, m_corrReferVals(0), VarPtr(m_corrReferTimeT2), m_corr2ndOrderVals(0), VarPtr(m_corr2ndOrderTimeT2))
  End If
#End If

  If (rc = False) Then
    build_spx_err_msg
  End If

  load_correction_file = rc
End Function

Function update_correction_file(updReferVals() As Double, timeT2ReferMins As Long, timeT2ReferMsecs As Long, numScansRefer As Integer, upd2ndOrderVals() As Double, timeT22ndOrderMins As Long, timeT22ndOrderMsecs As Long, numScans2ndOrder As Integer) As Boolean
  Dim rc As Boolean
  Dim timeT2Refer As TimeT2
  Dim timeT22ndOrder As TimeT2
  Dim ii As Integer

  timeT2Refer.mins = timeT2ReferMins
  timeT2Refer.msecs = timeT2ReferMsecs
  timeT22ndOrder.mins = timeT22ndOrderMins
  timeT22ndOrder.msecs = timeT22ndOrderMsecs

  ' Replace record in correction file with new spectra data
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.UpdateSpxCorrectionFile(unity_main.m_sys2500, m_numPts, updReferVals(0), VarPtr(timeT2Refer), numScansRefer, upd2ndOrderVals(0), VarPtr(timeT22ndOrder), numScans2ndOrder)
  
  If (SSRCSClientError = 0) Then
    ' Close spectra correction file
    SSRCSClientError = unity_main.SSRCSClient.CloseSpxCorrectionFile
    rc = True

#Else
  If (SpxFileIF1.UpdateCorrectionFile(unity_main.m_sys2500, updReferVals(0), VarPtr(timeT2Refer), numScansRefer, upd2ndOrderVals(0), VarPtr(timeT22ndOrder), numScans2ndOrder) = True) Then
    ' Close spectra correction file
    SpxFileIF1.CloseCorrectionFile
#End If

    ' Replace current sample with new scannned or selected history sample
    For ii = 0 To m_numPts - 1
      m_corrReferVals(ii) = updReferVals(ii)
      m_corr2ndOrderVals(ii) = upd2ndOrderVals(ii)
    Next ii

    m_corrReferTimeT2 = timeT2Refer
    m_corr2ndOrderTimeT2 = timeT22ndOrder

    ' Flag that a calibration reference has been updated
    m_refCalUpdatedFlg = True

    ' Flag take new reference scan
    m_refChkTaken = False
    m_refChkTimeStamp = 0
    m_refUpdTaken = False
    m_refUpdTimeStamp = 0

    ' Get configuration data to force MS11srv to reload spectra correction file
    rc = Get_MS11srv_Inst_Cfg()
    
    If (rc = False) Then
      m_errMsg = "Error getting SpectraStar configuration data"
      m_uniErrMsg = MLSupport.GSS("MS11srv", "errMsg6", "Error getting SpectraStar configuration data")
    Else
      rc = load_correction_file    ' reload spectra correction file
    End If
  Else
    build_spx_err_msg
  End If
  
  update_correction_file = rc
End Function

Private Function load_history_file() As Boolean
  Dim rc As Boolean
  Dim dirname As String
  Dim fileName As String
  Dim ii As Integer
  
  dirname = MS11_DIR
  fileName = "History_" & m_spxCorFile
  m_histNumRecs = -1
  
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.LoadSpxHistoryFile(dirname, fileName, m_scanStartWvln, m_scanEndWvln, MS11CfgData.wvlnIncr, unity_main.m_sys2500, m_histNumRecs)
  
  If (SSRCSClientError = 0) Then
    Dim timeStrg As String
    lst_spxHistory.Clear

    For ii = 0 To m_histNumRecs - 1
      SSRCSClientError = unity_main.SSRCSClient.GetSpxHistoryTimeT2(ii, unity_main.m_sys2500, VarPtr(m_histReferTimeT2), VarPtr(m_hist2ndOrderTimeT2))
      
      If (SSRCSClientError = 0) Then
        SSRCSClientError = unity_main.SSRCSClient.CnvtSpxTimeT2Strg(VarPtr(m_histReferTimeT2), timeStrg)
        
        If (SSRCSClientError = 0) Then
          lst_spxHistory.AddItem timeStrg
          rc = True
        Else
          rc = False
          Exit For
        End If
      Else
        rc = False
        Exit For
      End If
    Next ii
  End If

#Else
  If (SpxFileIF1.LoadHistoryFile(dirname, fileName, m_scanStartWvln, m_scanEndWvln, MS11CfgData.wvlnIncr, unity_main.m_sys2500, m_histNumRecs) = True) Then
    Dim timeStrg As String
    lst_spxHistory.Clear

    For ii = 0 To m_histNumRecs - 1
      If (SpxFileIF1.GetHistoryTimeT2(ii, unity_main.m_sys2500, VarPtr(m_histReferTimeT2), VarPtr(m_hist2ndOrderTimeT2)) = True) Then
        If (SpxFileIF1.CnvtTimeT2Strg(VarPtr(m_histReferTimeT2), timeStrg) = True) Then
          lst_spxHistory.AddItem timeStrg
          rc = True
        Else
          rc = False
          Exit For
        End If
      Else
        rc = False
        Exit For
      End If
    Next ii
  End If
#End If

  If (rc = False) Then
    build_spx_err_msg
  End If
  
  load_history_file = rc
End Function

Private Function update_history_file(updReferVals() As Double, vpTimeT2Refer As Long, numScansRefer As Integer, upd2ndOrderVals() As Double, vpTimeT22ndOrder As Long, numScans2ndOrder As Integer) As Boolean
  Dim rc As Boolean
  
  ' Replace record in correction file with new spectra data
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.UpdateSpxHistoryFile(unity_main.m_sys2500, m_numPts, updReferVals(0), vpTimeT2Refer, numScansRefer, upd2ndOrderVals(0), vpTimeT22ndOrder, numScans2ndOrder)
  
  If (SSRCSClientError = 0) Then
    rc = True
  End If
  
#Else
  rc = SpxFileIF1.UpdateHistoryFile(unity_main.m_sys2500, updReferVals(0), vpTimeT2Refer, numScansRefer, upd2ndOrderVals(0), vpTimeT22ndOrder, numScans2ndOrder)
#End If
  
  If (rc = False) Then
    build_spx_err_msg
  End If

  update_history_file = rc
End Function

Private Function setup_ref_scan() As Boolean

  ' Check if recalibrating system 2500
  If ((m_calibFunc = CF_RECALIBRATE) And (unity_main.m_sys2500 = True)) Then
    ' Create SPX dat file to store 2500 sampled calibration spectra
    If (create_2500_cal_file() = False) Then
      Exit Function
    End If
  End If

  ' Flag take new reference scan
  m_refChkTaken = False
  m_refChkTimeStamp = 0
  m_refUpdTaken = False
  m_refUpdTimeStamp = 0
  m_refFlag = True
  prg_refScan.percent = 0

  ' Setup scan configuration variables
  unity_main.m_smplEndWvln = m_scanEndWvln
  unity_main.m_smplStartWvln = m_scanStartWvln
  unity_main.m_intRefNScans = m_refNumScans
  unity_main.m_intRefPPT = 0
  
  ' Flag that a calibration reference has been taken
  m_refCalTakenFlg = True
  
  ' Start reference scan
  m_scanFlg = True
  unity_main.m_calibFunc = m_calibFunc
  unity_main.m_scanDataType = SDT_PRODINTREF
  unity_main.m_scanTmrState = STS_SETUP
  unity_main.tmr_ref.enabled = True
  
  ' Display Abort button
  cmd_abortScan.Visible = True
  m_scanAbortFlg = False
  
  setup_ref_scan = True
End Function

Private Sub setup_smpl_scan()

  m_refFlag = False
  
  ' Prompt for proper standard sample
  If (m_smplNum = 1) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_intRefCalMgmt", "prompt1", "Make sure the Unity R99 Standard sample is in place"), vbOKOnly
  Else
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_intRefCalMgmt", "prompt2", "Make sure the Unity SMP01 Standard sample is in place"), vbOKOnly
  End If

  ' Setup sample scan configuration variables
  unity_main.m_smplEndWvln = m_scanEndWvln
  unity_main.m_smplStartWvln = m_scanStartWvln
  unity_main.m_smplPPT = 0
  
  If (m_smplNum = 1) Then
    unity_main.m_smplNScans = m_smplNumReferScans
  Else
    unity_main.m_smplNScans = m_smplNum2ndOrderScans
  End If
  
  ' Setup tray configuration variables
  unity_main.m_trayNum = m_smplTrayNum
  unity_main.m_rotateDir = m_smplRotateDir
  unity_main.m_rotateIndexSteps = m_smplRotateIndexSteps
  unity_main.m_rotateMoveMode = m_smplRotateMoveMode
  unity_main.m_rotateSpeed = m_smplRotateSpeed
  unity_main.m_rotateStepSteps = m_smplRotateStepSteps
    
  ' Start sample scan
  m_scanFlg = True
  unity_main.m_calibFunc = m_calibFunc
  unity_main.m_smplRepacks = 1
  unity_main.m_scanDataType = SDT_PRODSMPL
  unity_main.m_scanTmrState = STS_SETUP
  unity_main.tmr_sample.enabled = True
  
  ' Display Abort button
  cmd_abortScan.Visible = True
  m_scanAbortFlg = False
End Sub

Private Function get_scan_data() As Boolean
  Dim ii As Integer
   
  ' Check if internal reference scan
  If (m_refFlag = True) Then
    ' Get reference scan data
    m_refIntenVals = ProdRefYVals
    
    ' Check if performing check reference function
    If (m_calibFunc = CF_CHECK) Then
      m_refChkTimeStamp = Timer
    Else    ' performing update calibration function
      ' Get reference timestamp
      m_refUpdTimeStamp = Timer
      
#If SSRCS Then
      SSRCSClientError = unity_main.SSRCSClient.GetSpxTimeT2(VarPtr(m_refUpdTimeT2))
#Else
      SpxFileIF1.GetTimeT2 VarPtr(m_refUpdTimeT2)
#End If

      ' Check if recalibrating system 2500
      If ((m_calibFunc = CF_RECALIBRATE) And (unity_main.m_sys2500 = True)) Then
        ' Save reference spectrum for 2500 math treatment operation
        If (update_2500_cal_file(0, m_refIntenVals, m_refUpdTimeT2) = False) Then
          Exit Function
        End If
      End If
    End If
  Else      ' get sample scan
    ' Check if performing check reference function
    If (m_calibFunc = CF_CHECK) Then
      ' Get sample scan data
      m_smplChkReferVals = ProdSmplYVals
      
      ' Get sample timestamp
#If SSRCS Then
      SSRCSClientError = unity_main.SSRCSClient.GetSpxTimeT2(VarPtr(m_smplChkTimeT2))
#Else
      SpxFileIF1.GetTimeT2 VarPtr(m_smplChkTimeT2)
#End If
    Else    ' performing update calibration function
      ' Check if scanned Fluorion standard
      If (m_smplNum = 1) Then
        ' Get sample scan data
        m_smplUpdReferVals = ProdSmplYVals
      
        ' Get sample timestamp
#If SSRCS Then
        SSRCSClientError = unity_main.SSRCSClient.GetSpxTimeT2(VarPtr(m_smplReferTimeT2))
#Else
        SpxFileIF1.GetTimeT2 VarPtr(m_smplReferTimeT2)
#End If
      
        ' Check if recalibrating system 2500
        If (unity_main.m_sys2500 = True) Then
          ' Calculate sample intensity from reflectance value
          For ii = 0 To m_numPts - 1
            m_smplIntenVals(ii) = m_smplUpdReferVals(ii) * m_refIntenVals(ii)
          Next ii

          ' Save Fluorion R99 intensity spectrum for 2500 math treatment operation
          If (update_2500_cal_file(1, m_smplIntenVals, m_smplReferTimeT2) = False) Then
            Exit Function
          End If
        End If
      Else  ' scanned Flour SMP01 standard
        ' Get sample scan data
        m_smplUpd2ndOrderVals = ProdSmplYVals
      
        ' Get sample timestamp
#If SSRCS Then
        SSRCSClientError = unity_main.SSRCSClient.GetSpxTimeT2(VarPtr(m_smpl2ndOrderTimeT2))
#Else
        SpxFileIF1.GetTimeT2 VarPtr(m_smpl2ndOrderTimeT2)
#End If
      
        ' Check if recalibrating system 2500
        If (unity_main.m_sys2500 = True) Then
          ' Calculate sample intensity from reflectance value
          For ii = 0 To m_numPts - 1
            m_smplIntenVals(ii) = m_smplUpd2ndOrderVals(ii) * m_refIntenVals(ii)
          Next ii

          ' Save Flour SMP01 intensity spectrum for 2500 math treatment operation
          If (update_2500_cal_file(2, m_smplIntenVals, m_smpl2ndOrderTimeT2) = False) Then
            Exit Function
          End If
        End If
      End If
    End If
  End If
  
  get_scan_data = True
End Function

Private Function create_2500_cal_file()
  Dim rc As Boolean
  Dim filePath As String
  Dim fileName As String
  
  filePath = MS11_DIR
  fileName = SPX_2500_CAL_FILE_NAME

#If SSRCS Then
  ' Create SPX dat file to store 2500 sampled calibration spectra
  SSRCSClientError = unity_main.SSRCSClient.CreateSpxDatFile(filePath, fileName, m_scanStartWvln, m_scanEndWvln, MS11CfgData.wvlnIncr, -1)
  
  If (SSRCSClientError = 0) Then
    rc = True
  End If

#Else
  ' Create SPX dat file to store 2500 sampled calibration spectra
  rc = SpxFileIF1.CreateDatFile(filePath, fileName, m_scanStartWvln, m_scanEndWvln, MS11CfgData.wvlnIncr, -1)
#End If
  
  If (rc = False) Then
    build_spx_err_msg
  End If

  create_2500_cal_file = rc
End Function

Private Function update_2500_cal_file(recIndx As Integer, updVals() As Double, timT2 As TimeT2)
  Dim rc As Boolean
  Dim filePath As String
  Dim fileName As String
  
  filePath = MS11_DIR
  fileName = SPX_2500_CAL_FILE_NAME

#If SSRCS Then
  ' Update SPX dat file with 2500 sampled calibration spectra
  SSRCSClientError = unity_main.SSRCSClient.UpdateSpxDatFile(filePath, fileName, recIndx, m_numPts, updVals(0), VarPtr(timT2), NUM_RECAL_SCANS)
  
  If (SSRCSClientError = 0) Then
    rc = True
  End If

#Else
  ' Update SPX dat file with 2500 sampled calibration spectra
  rc = SpxFileIF1.UpdateDatFile(filePath, fileName, recIndx, updVals(0), VarPtr(timT2), NUM_RECAL_SCANS)
#End If

  If (rc = False) Then
    build_spx_err_msg
  End If
  
  update_2500_cal_file = rc
End Function
  
Private Function apply_2500_math_treatment() As Boolean
  Dim rc As Boolean
  Dim filePath As String
  Dim batFileName As String
  Dim spxFileName As String
  Dim recalFileName As String
  
  filePath = MS11_DIR
  batFileName = MATH_TREAT_2500_BAT_FILE_NAME
  spxFileName = SPX_2500_CAL_FILE_NAME
  recalFileName = SPX_2500_RECAL_FILE_NAME

#If SSRCS Then
  ' Create batch file to perfrom math treatment on 2500 spectral file
  SSRCSClientError = unity_main.SSRCSClient.CreateSpx2500MathTreatBatFile(filePath, batFileName, spxFileName, m_spxCorFile, recalFileName)
  
  If (SSRCSClientError = 0) Then
    ' Perform math treatment on 2500 spectral file
    SSRCSClientError = unity_main.SSRCSClient.ApplySpx2500MathTreat(filePath, batFileName, recalFileName)
    
    If (SSRCSClientError = 0) Then
      rc = True
    End If
  End If

#Else
  ' Create batch file to perfrom math treatment on 2500 spectral file
  rc = SpxFileIF1.Create2500MathTreatBatFile(filePath, batFileName, spxFileName, m_spxCorFile, recalFileName)
  
  If (rc = True) Then
    ' Perform math treatment on 2500 spectral file
    rc = SpxFileIF1.Apply2500MathTreat(filePath, batFileName, recalFileName)
  End If
#End If

  If (rc = False) Then
    build_spx_err_msg
  End If
  
  apply_2500_math_treatment = rc
End Function

Private Function load_2500_recal_file() As Boolean
  Dim rc As Boolean
  Dim filePath As String
  Dim fileName As String
  Dim t2Refer As TimeT2
  Dim t22ndOrder As TimeT2
  
  filePath = MS11_DIR
  fileName = SPX_2500_RECAL_FILE_NAME

#If SSRCS Then
  ' Get 2500 math treated spectra data
  SSRCSClientError = unity_main.SSRCSClient.GetSpxDatSpectra(filePath, fileName, 0, unity_main.m_sys2500, m_numPts, m_smplUpdReferVals(0), VarPtr(t2Refer), m_smplUpd2ndOrderVals(0), VarPtr(t22ndOrder))
  
  If (SSRCSClientError = 0) Then
    rc = True
  End If

#Else
  ' Get 2500 math treated spectra data
  rc = SpxFileIF1.GetDatSpectra(filePath, fileName, 0, unity_main.m_sys2500, m_smplUpdReferVals(0), VarPtr(t2Refer), m_smplUpd2ndOrderVals(0), VarPtr(t22ndOrder))
#End If

  If (rc = False) Then
    build_spx_err_msg
  End If
  
  load_2500_recal_file = rc
End Function

Private Function delete_2500_files() As Boolean
  Dim rc As Boolean
  Dim filePath As String
  Dim batFileName As String
  Dim spxFileName As String
  Dim recalFileName As String
  
  filePath = MS11_DIR
  batFileName = MATH_TREAT_2500_BAT_FILE_NAME
  spxFileName = SPX_2500_CAL_FILE_NAME
  recalFileName = SPX_2500_RECAL_FILE_NAME

#If SSRCS Then
  ' Delete 2500 batch file
  SSRCSClientError = unity_main.SSRCSClient.DeleteSpxFile(filePath, batFileName)
  
  If (SSRCSClientError = 0) Then
    ' Delete 2500 calibration spectra file
    SSRCSClientError = unity_main.SSRCSClient.DeleteSpxFile(filePath, spxFileName)
    
    If (SSRCSClientError = 0) Then
      ' Delete 2500 recalibration spectra file
      SSRCSClientError = unity_main.SSRCSClient.DeleteSpxFile(filePath, recalFileName)
    
      If (SSRCSClientError = 0) Then
        rc = True
      End If
    End If
  End If

#Else
  ' Delete 2500 batch file
  SpxFileIF1.DeleteFile filePath, batFileName

  ' Delete 2500 calibration spectra file
  SpxFileIF1.DeleteFile filePath, spxFileName

  ' Delete 2500 recalibration spectra file
  SpxFileIF1.DeleteFile filePath, recalFileName
  rc = True
#End If

  If (rc = False) Then
    build_spx_err_msg
  End If
  
  delete_2500_files = rc
End Function

Private Sub check_reference()
  Dim ii As Integer
  Dim reflectVal As Double
  
  ' Init max/min values
  m_chkMaxDevVal = 0#
  m_chkMaxReflectVal = 0#
  m_chkMinReflectVal = 65535#

  For ii = m_chkStartIndx To m_chkEndIndx - 1
    ' Get spectrum reflectance value
    reflectVal = m_smplChkReferVals(ii)

    ' Check if reflectance value greater than max limit or less than min limit
    If ((reflectVal > m_chkMaxReflectValLmt) Or (reflectVal < m_chkMinReflectValLmt)) Then
      m_chkRefState = -1     ' mark spectrum failed limits
      Exit For
    End If

    ' Check if reflectance value is maximum value
    If (reflectVal > m_chkMaxReflectVal) Then
      m_chkMaxReflectVal = reflectVal
    End If

    ' Check if reflectance value is minimum value
    If (reflectVal < m_chkMinReflectVal) Then
      m_chkMinReflectVal = reflectVal
    End If
  Next ii

  If (m_chkRefState = 0) Then
    ' Check if maximum reflectance deviation is greater than limit
    m_chkMaxDevVal = m_chkMaxReflectVal - m_chkMinReflectVal

    If (m_chkMaxDevVal > m_chkMaxDevValLmt) Then
      m_chkRefState = -1     ' mark spectrum failed limits
    Else
      m_chkRefState = 1      ' mark spectrum passed limits
    End If
  End If
End Sub

Private Sub load_chk_ref_lmts_file()
  Dim fileName As String
  Dim numPts As Integer
  Dim ii As Integer

  m_badIniVal = False

  ' Initialize default setting values
  m_chkEndWvlnLmt = MS11DfltScanCfgData.endWvln
  m_chkMaxDevValLmt = CHK_MAX_DEV_VAL_LMT
  m_chkMaxReflectValLmt = CHK_MAX_REFLECT_VAL_LMT
  m_chkMinReflectValLmt = CHK_MIN_REFLECT_VAL_LMT
  m_chkStartWvlnLmt = MS11DfltScanCfgData.startWvln
  
  ' Check if configuration file exists
  fileName = (MS11_DIR & CHK_REF_LMTS_FILE)
  
  If (CFile.st_FileExist(fileName) = True) Then
    process_chk_ref_lmts fileName
  End If
  
  ' Check ending wavelength limit
  If (m_chkEndWvlnLmt > MS11CfgData.maxWvln) Or (m_chkEndWvlnLmt <= m_chkStartWvlnLmt) Then
    unity_main.errorstring = (fileName & " had incompatible value. EndWvlnLmt was " & m_chkEndWvlnLmt & "; updated to " & MS11DfltScanCfgData.endWvln)
    unity_main.write_error
    m_chkEndWvlnLmt = MS11DfltScanCfgData.endWvln
    m_badIniVal = True
  End If
  
  ' Check deviation limit
  If (m_chkMaxDevValLmt <= 0#) Then
    unity_main.errorstring = (fileName & " had incompatible value. MaxDevValLmt was " & m_chkMaxDevValLmt & "; updated to " & CHK_MAX_DEV_VAL_LMT)
    unity_main.write_error
    m_chkMaxDevValLmt = CHK_MAX_DEV_VAL_LMT
    m_badIniVal = True
  End If
  
  ' Check max reflectance limit
  If (m_chkMaxReflectValLmt <= CHK_TARGET_VAL) Then
    unity_main.errorstring = (fileName & " had incompatible value. MaxReflectValLmt was " & m_chkMaxReflectValLmt & "; updated to " & CHK_MAX_REFLECT_VAL_LMT)
    unity_main.write_error
    m_chkMaxReflectValLmt = CHK_MAX_REFLECT_VAL_LMT
    m_badIniVal = True
  End If
  
  ' Check min reflectance limit
  If (m_chkMinReflectValLmt <= 0#) Or (m_chkMinReflectValLmt >= CHK_TARGET_VAL) Then
    unity_main.errorstring = (fileName & " had incompatible value. MinReflectValLmt was " & m_chkMinReflectValLmt & "; updated to " & CHK_MIN_REFLECT_VAL_LMT)
    unity_main.write_error
    m_chkMinReflectValLmt = CHK_MIN_REFLECT_VAL_LMT
    m_badIniVal = True
  End If
  
  ' Check starting wavelength limit
  If (m_chkStartWvlnLmt < MS11CfgData.minWvln) Or (m_chkStartWvlnLmt >= m_chkEndWvlnLmt) Then
    unity_main.errorstring = (fileName & " had incompatible value. StartWvlnLmt was " & m_chkStartWvlnLmt & "; updated to " & MS11DfltScanCfgData.startWvln)
    unity_main.write_error
    m_chkStartWvlnLmt = MS11DfltScanCfgData.startWvln
    m_badIniVal = True
  End If
  
  ' Check if ini file had bad value
  If (m_badIniVal = True) Then
    unity_main.errorstring = (fileName & " had incompatible value(s). Updated file with proper/default value(s).")
    unity_main.write_error
    m_uniErrMsg = MLSupport.GGS_Params("fileErrMsg7", "%1 had incompatible value(s). Updated file with proper/default value(s)", fileName)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg3", "%1. Refer to system log for more details", m_uniErrMsg), vbOKOnly
    save_chk_ref_lmts fileName
  End If
  
  numPts = (m_chkEndWvlnLmt - m_chkStartWvlnLmt) / MS11CfgData.wvlnIncr
  m_chkStartIndx = (m_chkStartWvlnLmt - m_scanStartWvln) / MS11CfgData.wvlnIncr
  m_chkEndIndx = m_chkStartIndx + numPts
  ReDim m_chkMaxReflectLmts(numPts)
  ReDim m_chkMinReflectLmts(numPts)
  ReDim m_chkTargets(numPts)

  For ii = 0 To numPts
    m_chkMaxReflectLmts(ii) = m_chkMaxReflectValLmt
    m_chkMinReflectLmts(ii) = m_chkMinReflectValLmt
    m_chkTargets(ii) = CHK_TARGET_VAL
  Next ii
End Sub

Private Sub process_chk_ref_lmts(fileName As String)
  Dim iniString As String
  Dim lineCnt As Integer
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim rc As Boolean
  Dim pos As Integer
  Dim tmpStrg As String
  Dim cfgVar As String
  Dim varVal As String
  Dim strlen As Integer

  ' Load check reference limits configuration file
  If (uniFile.OpenFileRead(fileName) = False) Then GoTo FILE_ERROR
  
  fEncoding = uniFile.ReadBOM
    
  ' Process each line in .ini file
  While Not (uniFile.EOF())
    ' Read line from file
    On Error GoTo FILE_ERROR
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
      Case "endwvlnlmt"
        m_chkEndWvlnLmt = CDbl(varVal)
      Case "maxdevvallmt"
        m_chkMaxDevValLmt = CDbl(varVal)
      Case "maxreflectvallmt"
        m_chkMaxReflectValLmt = CDbl(varVal)
      Case "minreflectvallmt"
        m_chkMinReflectValLmt = CDbl(varVal)
      Case "startwvlnlmt"
        m_chkStartWvlnLmt = CDbl(varVal)
    End Select
  Wend
  
  uniFile.CloseFile
  Exit Sub
  
BAD_INI_VALUE:
  unity_main.errorstring = (fileName & " had incompatible value. " & cfgVar & " = " & varVal & "; will use default value")
  unity_main.write_error
  m_badIniVal = True
  Resume Next

FILE_ERROR:
  If (lineCnt = 0) Then
    m_errMsg = (fileName & " file open error." & Error$)
    m_uniErrMsg = MLSupport.GGS_Params("fileErrMsg2", "%1 file open error. %2", fileName, Error$)
  Else
    m_errMsg = (fileName & " file has error on line " & CStr(lineCnt) & ". " & Error$)
    m_uniErrMsg = MLSupport.GGS_Params("fileErrMsg3", "%1 file has error on line %2. %3", fileName, CStr(lineCnt), Error$)
  End If
  
  unity_main.errorstring = m_errMsg
  unity_main.write_error
  
LOAD_ERROR:
  m_uniErrMsg = MLSupport.GGS_Params("frm_intRefCalMgmt.errMsg4", "%1. Using default limit values", m_uniErrMsg)
  CWrap.ShowMessageBoxW m_uniErrMsg, vbCritical
  uniFile.CloseFile
End Sub

Private Sub save_chk_ref_lmts(fileName As String)
  Dim uniFile As New clsUniFile

  ' Check if file can be created
  If (uniFile.OpenFileWrite(fileName) = True) Then
    On Error GoTo FILE_ERROR
    uniFile.WriteBOM fe_ANSI
    uniFile.WriteAnsiLine ("EndWvlnLmt=" & m_chkEndWvlnLmt)
    uniFile.WriteAnsiLine ("MaxDevValLmt=" & m_chkMaxDevValLmt)
    uniFile.WriteAnsiLine ("MaxReflectValLmt=" & m_chkMaxReflectValLmt)
    uniFile.WriteAnsiLine ("MinReflectValLmt=" & m_chkMinReflectValLmt)
    uniFile.WriteAnsiLine ("StartWvlnLmt=" & m_chkStartWvlnLmt)
    uniFile.Flush
  Else
FILE_ERROR:
    m_errMsg = (fileName & " file write error. " & Error$)
    unity_main.errorstring = m_errMsg
    unity_main.write_error
    m_uniErrMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", fileName, Error$)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", m_uniErrMsg), vbCritical
  End If
  
  uniFile.CloseFile
End Sub

Private Sub init_scan_info()
  Dim ii As Integer

  m_scanAbortFlg = False
  m_scanFlg = False

  m_refCalTakenFlg = False
  m_refCalUpdatedFlg = False
  m_refChkTaken = False
  m_refChkTimeStamp = 0
  m_refFlag = False
  m_refNumScans = NUM_CHECK_SCANS
  m_refPPT = 0
  m_refUpdTaken = False
  m_refUpdTimeStamp = 0

  m_smplChkTimeT2.msecs = 0
  m_smplChkTimeT2.mins = 0
  m_smplNumReferScans = NUM_CHECK_SCANS
  m_smplNum2ndOrderScans = NUM_CHECK_SCANS
  m_smplReady2Save = False
  m_smplRotateIndexSteps = 0
  m_smplRotateDir = TRD_NONE
  m_smplRotateMoveMode = TRM_NONE
  m_smplRotateSpeed = 0
  m_smplRotateStepSteps = 0
  m_smplTrayNum = 1 ' was 3           ' B&L small cup

  If (MS11CfgData.spxStartWvln = 0) And (MS11CfgData.spxEndWvln = 0) Then
    ' Save min/max spectrum wavelength used by instrument
    m_scanStartWvln = MS11CfgData.minWvln
    m_scanEndWvln = MS11CfgData.maxWvln
  Else
    ' Save min/max spectrum wavelength used by Spectrix file
    m_scanStartWvln = MS11CfgData.spxStartWvln
    m_scanEndWvln = MS11CfgData.spxEndWvln
  End If

  m_numPts = (m_scanEndWvln - m_scanStartWvln) / MS11CfgData.wvlnIncr + 1
  ReDim m_corrReferVals(m_numPts - 1)
  ReDim m_corr2ndOrderVals(m_numPts - 1)
  ReDim m_histReferVals(m_numPts - 1)
  ReDim m_hist2ndOrderVals(m_numPts - 1)
  ReDim m_refIntenVals(m_numPts - 1)
  ReDim m_smplChkReferVals(m_numPts - 1)
  ReDim m_smplChk2ndOrderVals(m_numPts - 1)
  ReDim m_smplIntenVals(m_numPts - 1)
  ReDim m_smplUpdReferVals(m_numPts - 1)
  ReDim m_smplUpd2ndOrderVals(m_numPts - 1)
  ReDim m_wvlnVals(m_numPts - 1)

  ' Setup spectrum wavelength (X) values
  For ii = 0 To m_numPts - 1
    m_wvlnVals(ii) = m_scanStartWvln + ii
  Next ii
End Sub

Private Sub init_plot()
  Dim rc As Boolean
  
  ' Setup sample spectrum plot
  xyPlot.NumSubsets = 4            ' 4 pens
  xyPlot.AxisMinMaxPad = 3         ' This property controls the percent of auto scalling to X/Y axis

  Call xyPlot.Initialize(unity_main.m_minWvln, unity_main.m_maxWvln, MS11CfgData.wvlnIncr, _
                         False, False, True, True, LegendStyles.ONE_LINE, True)
  xyPlot.SubTitle = ""
  xyPlot.XAxisLabel = "Wavelength (nm)"
End Sub

Private Sub setup_plot()
  
  ' Setup sample spectrum plot based on function
  Select Case (m_calibFunc)
    Case CF_CHECK               ' check reference
      xyPlot.MainTitle = "Check Reference"
      xyPlot.YAxisLabel = "Reflectance"
      xyPlot.SubsetLabel(0) = "Scanned"
      xyPlot.SubsetColor(0) = vbBlack
      xyPlot.SubsetLabel(1) = "Upper Limit"
      xyPlot.SubsetColor(1) = RGB(255, 128, 0)   ' orange
      xyPlot.LineType(1) = LineTypes.DASH
      xyPlot.SubsetLabel(2) = "Lower Limit"
      xyPlot.SubsetColor(2) = RGB(255, 128, 0)   ' orange
      xyPlot.LineType(2) = LineTypes.DASH_DOT
      xyPlot.SubsetLabel(3) = "Target"
      xyPlot.SubsetColor(3) = RGB(0, 0, 255)     ' blue
      xyPlot.LineType(3) = LineTypes.DOT
    
    Case CF_RECALIBRATE        ' recalibrate reference
      xyPlot.MainTitle = "Recal Reference"
      xyPlot.YAxisLabel = "Correction Factor"
      xyPlot.SubsetLabel(0) = "Current"
      xyPlot.SubsetColor(0) = RGB(0, 128, 0)     ' dark green
      xyPlot.SubsetLabel(1) = "Scanned"
      xyPlot.SubsetColor(1) = RGB(255, 0, 255)   ' magenta
      xyPlot.LineType(1) = LineTypes.THIN_SOLID
      xyPlot.SubsetLabel(2) = ""
      xyPlot.SubsetColor(2) = vbWhite
      xyPlot.LineType(2) = LineTypes.THIN_SOLID
      xyPlot.SubsetLabel(3) = ""
      xyPlot.SubsetColor(3) = vbWhite
      xyPlot.LineType(3) = LineTypes.THIN_SOLID
      
    Case CF_HISTORY                 ' calibration history
      xyPlot.MainTitle = "Calibration History"
      xyPlot.YAxisLabel = "Correction Factor"
      xyPlot.SubsetLabel(0) = "Current"
      xyPlot.SubsetColor(0) = RGB(0, 128, 0)     ' dark green
      xyPlot.SubsetLabel(1) = "Selected"
      xyPlot.SubsetColor(1) = RGB(0, 0, 255)     ' blue
      xyPlot.LineType(1) = LineTypes.THIN_SOLID
      xyPlot.SubsetLabel(2) = ""
      xyPlot.SubsetColor(2) = vbWhite
      xyPlot.LineType(2) = LineTypes.THIN_SOLID
      xyPlot.SubsetLabel(3) = ""
      xyPlot.SubsetColor(3) = vbWhite
      xyPlot.LineType(3) = LineTypes.THIN_SOLID
  End Select
End Sub

Private Sub plot_scan_data(usDblYVals() As Double, subsetIndx As Integer, vpTimeT2 As Long)
  Dim rc As Boolean
  Dim timeStrg As String
  
  rc = xyPlot.PlotSpectrum2(subsetIndx, m_scanStartWvln, m_scanEndWvln, 1#, 0, m_wvlnVals(0), usDblYVals(0), Null, Null, False, m_errMsg)
  
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.CnvtSpxTimeT2Strg(vpTimeT2, timeStrg)
#Else
  SpxFileIF1.CnvtTimeT2Strg vpTimeT2, timeStrg
#End If

  If (subsetIndx = 0) Then
    lbl_currScanTime.Caption = timeStrg
    lbl_currScanTime.ForeColor = RGB(0, 128, 0)  ' dark green
  Else
    lbl_newScanTime.Caption = timeStrg
    
    If (m_calibFunc = CF_RECALIBRATE) Then
      lbl_newScanTime.ForeColor = RGB(255, 0, 255)  ' magenta
    Else
      lbl_newScanTime.ForeColor = RGB(0, 0, 255)    ' blue
    End If
  End If
End Sub

Private Sub plot_ref_limits()
  Dim rc As Boolean
  
  rc = xyPlot.PlotSpectrum2(1, m_chkStartWvlnLmt, m_chkEndWvlnLmt, 1#, 0, m_wvlnVals(m_chkStartIndx), m_chkMaxReflectLmts(0), vbNull, vbNull, False, m_errMsg)
  rc = xyPlot.PlotSpectrum2(2, m_chkStartWvlnLmt, m_chkEndWvlnLmt, 1#, 0, m_wvlnVals(m_chkStartIndx), m_chkMinReflectLmts(0), vbNull, vbNull, False, m_errMsg)
  rc = xyPlot.PlotSpectrum2(3, m_chkStartWvlnLmt, m_chkEndWvlnLmt, 1#, 0, m_wvlnVals(m_chkStartIndx), m_chkTargets(0), vbNull, vbNull, False, m_errMsg)

  lbl_newScanTime.Caption = ""
End Sub

Private Sub clr_plot_scan_data(subsetIndx As Integer)
  Dim rc As Boolean

  rc = xyPlot.ClearSpectrum(subsetIndx, m_errMsg)
  xyPlot.UnZoom

  Select Case (subsetIndx)
    Case 0:
      lbl_currScanTime.Caption = ""
    Case 1:
      lbl_newScanTime.Caption = ""
  End Select
End Sub

Private Sub build_spx_err_msg()
  Dim errorCode As SPX_CORR_FILE_ERRORS
  Dim lastErrCode As Long
  Dim lastSpxErrCode As Long

#If SSRCS Then
  If (SSRCSClientError = 0) Or (SSRCSClientError = ERR_CMD_METHOD_FAIL) Then
    SSRCSClientError = unity_main.SSRCSClient.GetSpxLastErrorCode(lastErrCode)
    SSRCSClientError = unity_main.SSRCSClient.GetSpxLastSpxErrorCode(lastSpxErrCode)
  Else
    m_errMsg = "SSRCS error code: " & SSRCSClientError
    m_uniErrMsg = MLSupport.GGS_Params("frm_intRefCalMgmt.errMsg12", "SSRCS error code: %1", CStr(SSRCSClientError))
  End If
#Else
  lastErrCode = SpxFileIF1.LastErrorCode
  lastSpxErrCode = SpxFileIF1.LastSpxErrorCode
#End If

  errorCode = lastErrCode
    
  Select Case (errorCode)
    Case SPX_FILE_NULL_ARG_ERR            ' nulled argument passed
      m_errMsg = "Error opening spectra correction file. Null argument passed"
      m_uniErrMsg = MLSupport.GSS("frm_intRefCalMgmt", "errMsg2", "Error opening spectra correction file. Null argument passed")

    Case SPX_CORR_FILE_NOT_LOAD_ERR       ' correction file not loaded
      m_errMsg = "Spectra correction file not loaded"
      m_uniErrMsg = MLSupport.GSS("frm_intRefCalMgmt", "errMsg3", "Spectra correction file not loaded")

    Case SPX_CORR_FILE_ERR                ' correction file error (details stored in m_spxLastErrorCode)
      m_errMsg = "Error " & lastSpxErrCode & " opening spectra correction file"
      m_uniErrMsg = MLSupport.GGS_Params("frm_intRefCalMgmt.errMsg5", "Error %1 opening spectra correction file", CStr(lastSpxErrCode))

    Case SPX_CORR_FILE_INV_WVLN_ERR       ' correction file has wavelength parameters mismatch
      Dim corStartWvln As Double
      Dim corEndWvln As Double
      Dim corWvlnIncr As Double

#If SSRCS Then
      SSRCSClientError = unity_main.SSRCSClient.GetSpxCorrcorStartWvln(corStartWvln)
      SSRCSClientError = unity_main.SSRCSClient.GetSpxCorrEndWvln(corEndWvln)
      SSRCSClientError = unity_main.SSRCSClient.GetSpxCorrWvlnIncr(corWvlnIncr)
#Else
      corStartWvln = SpxFileIF1.CorrStartWvln
      corEndWvln = SpxFileIF1.CorrEndWvln
      corWvlnIncr = SpxFileIF1.CorrWvlnIncr
#End If
      m_errMsg = "Spectra correction file wavelength incompatibility error; File (" & corStartWvln & "-" & corEndWvln & "," & corWvlnIncr & ", MS1100 (" & m_scanStartWvln & "-" & m_scanEndWvln & "," & MS11CfgData.wvlnIncr & ")"
      m_uniErrMsg = MLSupport.GGS_Params("frm_intRefCalMgmt.errMsg7", "Spectra correction file wavelength incompatibility error; File (%1-%2,%3)", CStr(corStartWvln), CStr(corEndWvln), CStr(corWvlnIncr))
      m_uniErrMsg = MLSupport.GGS_Params("frm_intRefCalMgmt.errMsg6", "%1, MS1000 (%2-%3,%4)", m_uniErrMsg, CStr(m_scanStartWvln), CStr(m_scanEndWvln), CStr(MS11CfgData.wvlnIncr))

    Case SPX_CORR_FILE_INV_RECS_ERR       ' correction file has improper number of correction spectra records for system type
      m_errMsg = "Invalid number of records (" & m_corrNumRecs & ") in spectra correction file"
      m_uniErrMsg = MLSupport.GGS_Params("frm_intRefCalMgmt.errMsg8", "Invalid number of records (%1) in spectra correction file", CStr(m_corrNumRecs))
  
    Case SPX_HIST_FILE_NOT_LOAD_ERR       ' history correction file not loaded
      m_errMsg = "History correction file not loaded"
      m_uniErrMsg = MLSupport.GSS("frm_intRefCalMgmt", "errMsg4", "History correction file not loaded")

    Case SPX_HIST_FILE_ERR                ' history correction file error (details stored in m_spxLastErrorCode)
      m_errMsg = "Error " & lastSpxErrCode & " opening history correction file"
      m_uniErrMsg = MLSupport.GGS_Params("frm_intRefCalMgmt.errMsg9", "Error %1 opening history correction file", CStr(lastSpxErrCode))

    Case SPX_HIST_FILE_INV_WVLN_ERR       ' history correction file has wavelength parameters mismatch
      Dim hstStartWvln As Double
      Dim hstEndWvln As Double
      Dim hstWvlnIncr As Double
        
#If SSRCS Then
      SSRCSClientError = unity_main.SSRCSClient.GetSpxHistStartWvln(hstStartWvln)
      SSRCSClientError = unity_main.SSRCSClient.GetSpxHistEndWvln(hstEndWvln)
      SSRCSClientError = unity_main.SSRCSClient.GetSpxHistWvlnIncr(hstWvlnIncr)
#Else
      hstStartWvln = SpxFileIF1.HistStartWvln
      hstEndWvln = SpxFileIF1.HistEndWvln
      hstWvlnIncr = SpxFileIF1.HistWvlnIncr
#End If
      m_errMsg = "Spectra history file wavelength incompatibility error; File (" & hstStartWvln & "-" & hstEndWvln & "," & hstWvlnIncr & ", MS1100 (" & m_scanStartWvln & "-" & m_scanEndWvln & "," & MS11CfgData.wvlnIncr & ")"
      m_uniErrMsg = MLSupport.GGS_Params("frm_intRefCalMgmt.errMsg10", "Spectra history file wavelength incompatibility error; File (%1-%2,%3)", CStr(hstStartWvln), CStr(hstEndWvln), CStr(hstWvlnIncr))
      m_uniErrMsg = MLSupport.GGS_Params("frm_intRefCalMgmt.errMsg6", "%1, MS1000 (%2-%3,%4)", m_uniErrMsg, CStr(m_scanStartWvln), CStr(m_scanEndWvln), CStr(MS11CfgData.wvlnIncr))

    Case SPX_HIST_FILE_INV_REC_INDX_ERR   ' invalid record index for history correction file
      m_errMsg = "Invalid record index for records (" & m_histNumRecs & ") in history correction file"
      m_uniErrMsg = MLSupport.GGS_Params("frm_intRefCalMgmt.errMsg11", "Invalid record index for records (%1) in history correction file", CStr(m_histNumRecs))

    Case SPX_BAT_FILE_CREATE_ERR          ' error creating 2500 math treatment batch file
      m_errMsg = "Error creating 2500 math treatment batch file"
      m_uniErrMsg = MLSupport.GSS("frm_intRefCalMgmt", "errMsg5", "Error creating 2500 math treatment batch file")

    Case SPX_BAT_FILE_EXE_ERR             ' error excuting 2500 math treatment batch file (details stored in m_spxLastErrorCode)
      m_errMsg = "Error " & lastSpxErrCode & " excuting 2500 math treatment batch file"
      m_uniErrMsg = MLSupport.GGS_Params("frm_intRefCalMgmt.errMsg13", "Error %1 excuting 2500 math treatment batch file", CStr(lastSpxErrCode))
  End Select
End Sub

Private Sub display_err_msg()

  ' Disable all menu items
  CheckMenu.enabled = False
  RecalibrateMenu.enabled = False
  HistoryMenu.enabled = False

  ' Hide all buttons except Exit
  cmd_abortScan.Visible = False
  cmd_restoreCalib.Visible = False
  cmd_startScan.Visible = False
  cmd_updateCalib.Visible = False

  ' Display Exit button
  cmd_exit.Visible = True

  ' Display error message
  unity_main.errorstring = m_errMsg
  unity_main.write_error
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", m_uniErrMsg), vbCritical
End Sub

Private Sub CheckMenu_Click()

  unity_main.errorstring = "Internal Reference Calibration Management screen Check Reference menu option selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  ' Disable menu items
  CheckMenu.enabled = False

  ' Hide buttons
  cmd_restoreCalib.Visible = False

  ' Hide List Box
  lst_spxHistory.Visible = False

  ' Enable menu items
  RecalibrateMenu.enabled = True
  HistoryMenu.enabled = True

  ' Display buttons
  cmd_startScan.Visible = True

  ' Display Frames
  frame_ref.Visible = True
  frame_smpl.Visible = True

  m_calibFunc = CF_CHECK   ' Check reference function
  setup_plot
  plot_ref_limits

  If ((m_smplChkTimeT2.msecs <> 0) Or (m_smplChkTimeT2.mins <> 0)) Then
    Select Case (m_chkRefState)
      Case -1                ' check reference failed
        xyPlot.SubsetColor(0) = vbRed
        lbl_newScanTime.ForeColor = vbRed
        lbl_newScanTime.Caption = Chr(34) & "FAILED" & Chr(34)

      Case 0                 ' no check reference
        xyPlot.SubsetColor(0) = vbBlack
        lbl_newScanTime.ForeColor = vbBlack
        lbl_newScanTime.Caption = ""

      Case 1                 ' check reference passed
        xyPlot.SubsetColor(0) = RGB(0, 128, 0)  ' dark green
        lbl_newScanTime.ForeColor = RGB(0, 128, 0)  ' dark green
        lbl_newScanTime.Caption = Chr(34) & "PASSED" & Chr(34)
    End Select

    plot_scan_data m_smplChkReferVals, 0, VarPtr(m_smplChkTimeT2)
  Else
    If ((m_refChkTaken = True) And (unity_main.chk_timeout(m_refChkTimeStamp, REF_EXPIRE_TIME) = False)) Then
      prg_refScan.percent = 100
    Else
      prg_refScan.percent = 0
    End If
    
    clr_plot_scan_data 0
  End If

  m_refNumScans = NUM_CHECK_SCANS
  m_smplNumReferScans = NUM_CHECK_SCANS
  m_smplNum2ndOrderScans = NUM_CHECK_SCANS
  lbl_numRefScans.Caption = m_refNumScans
  lbl_numSmplScans.Caption = m_smplNumReferScans
End Sub

Private Sub cmd_abortScan_Click()
  Dim StartTime As Single

  unity_main.errorstring = "Internal Reference Calibration Management screen Abort Scan button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  ' Hide Abort button
  cmd_abortScan.Visible = False
  
  ' Exit if not scanning
  If (m_scanFlg = False) Then
    Exit Sub
  End If
  
  ' Wait if scan is in setup mode
  While (unity_main.m_scanTmrState = STS_SETUP)
    DoEvents
  Wend
  
  ' Check if scan is running
  If (unity_main.m_scanTmrState = STS_WAIT_CMP) Then
    DoEvents
  
    ' Clear any previous errors
    Clear_MS11_Error_Codes
  
#If SSRCS Then
    SSRCSClientError = unity_main.SSRCSClient.scanStop
    
    If (SSRCSClientError = 0) Then
#Else
    If (unity_main.MS11srv.scanStop() = True) Then
#End If
      unity_main.m_scanState = SS_ABORT
      StartTime = Timer
      
      ' Delay to allow instrument to actually stop scan
      Do While (True)
        ' Check scan state to see if not actively scanning
#If SSRCS Then
        Dim scanState As Long
        SSRCSClientError = unity_main.SSRCSClient.GetScanState(scanState)
        
        If (scanState <> 5) Then GoTo CLEAR_SCAN

        If (unity_main.chk_timeout(StartTime, 8) = True) Then
#Else
        If (unity_main.MS11srv.scanState <> 5) Then GoTo CLEAR_SCAN

        If (unity_main.chk_timeout(StartTime, 5) = True) Then
#End If
          Exit Do
        End If
    
        DoEvents
      Loop
      
CLEAR_SCAN:
#If SSRCS Then
      SSRCSClientError = unity_main.SSRCSClient.scanDataClr
      
      If (SSRCSClientError <> 0) Then
#Else
      If (unity_main.MS11srv.scanDataClr() = False) Then
#End If
        ' Report error codes
        m_errMsg = "Error clearing scan data"
        m_uniErrMsg = MLSupport.GSS("OperStatus", "status27", "Error clearing scan data")
        m_scanState = CSS_ERROR
      End If
    End If
  End If
  
  ' Wait until scan is fully completed or stopped due to error
  While ((unity_main.m_scanTmrState <> STS_COMPLETED) And (unity_main.m_scanTmrState <> STS_ABORT))
    DoEvents
  Wend
  
  m_scanFlg = 0
End Sub

Private Sub cmd_exit_Click()

  If (m_restartFlg = True) Then
    Call unity_main.restart_loop(LOG_DBG_LEVEL3, "Internal Reference Calibration Management screen Exit button selected")
  Else
    unity_main.errorstring = "Internal Reference Calibration Management screen Exit button selected"
    unity_main.write_error (LOG_DBG_LEVEL3)
  End If
  
  ' Check if any calibration reference has been taken
  If (m_refCalTakenFlg = True) Then
    ' Recover current internal reference scan parameters
    unity_main.m_intRefNScans = m_intRefNScans
    unity_main.m_intRefPPT = m_intRefPPT
  
    ' Reload current product configuration file
    Call unity_main.load_prod_file("", True)
  
    ' Process according to product reference type
    Select Case (unity_main.m_bType)
      Case "internal"
        ' Flag to reload internal reference spectrum from file
        unity_main.m_intRefFileLoadFlg = True

        ' Check if qualifying reference
        If (unity_main.m_intRefPPT <> 0) Then
          ' Flag to reload internal reference PPT spectrum from file
          unity_main.m_intRefPPTFileSetup = False
        End If
        
        Dim refTime As Long
#If SSRCS Then
        SSRCSClientError = unity_main.SSRCSClient.GetRefTimeout(refTime)
#Else
        refTime = unity_main.MS11srv.refTimeout
#End If

        ' Check if reference timeout configured and has timed out
        If ((unity_main.m_intRefTimeout > 0) And (refTime = 0)) Then
          unity_main.m_intRefAutoScan = True
        End If
        
      Case "external"
        ' Flag to reload external reference spectrum from file
        unity_main.m_extRefFileSetup = False
        
        ' Check if qualifying reference
        If (unity_main.m_extRefPPT <> 0) Then
          ' Flag to reload external reference PPT spectrum from file
          unity_main.m_extRefPPTFileSetup = False
        End If
        
      Case "file"
        ' Flag to reload off-line reference spectrum from file
        unity_main.m_olRefFileSetup = False
    End Select
    
    unity_main.clear_GN_eventQ
  End If
  
  ' Check if any calibration reference has been updated
  If (m_refCalUpdatedFlg = True) Then
    ' Check if product using internal reference
    If (unity_main.m_bType = "internal") And (unity_main.m_intRefPPT <> 0) Then
      unity_main.m_intRefPPTScan = True
    End If
  End If
  
  unity_main.m_intRefCalFlg = False
  Unload frm_intRefCalMgmt
End Sub

Private Sub cmd_restoreCalib_Click()

  unity_main.errorstring = "Internal Reference Calibration Management screen Restore Calibration button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  ' Hide Restore Calibration button
  cmd_restoreCalib.Visible = False

  If (update_correction_file(m_histReferVals, m_histReferTimeT2.mins, m_histReferTimeT2.msecs, m_histNumReferScans, m_hist2ndOrderVals, m_hist2ndOrderTimeT2.mins, m_hist2ndOrderTimeT2.msecs, m_histNum2ndOrderScans) = True) Then
    clr_plot_scan_data 1
    plot_scan_data m_histReferVals, 0, VarPtr(m_histReferTimeT2)
    CWrap.ShowMessageBoxW "Spectra correction file restored!"
    lst_spxHistory.ListIndex = -1
  Else
    display_err_msg
  End If
End Sub

Private Sub cmd_startScan_Click()

  unity_main.errorstring = "Internal Reference Calibration Management screen Start Scan button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  ' Disable menu items
  CheckMenu.enabled = False
  RecalibrateMenu.enabled = False
  HistoryMenu.enabled = False

  ' Hide buttons
  cmd_exit.Visible = False
  cmd_updateCalib.Visible = False
  cmd_startScan.Visible = False

  ' Check if performing reference check function
  If (m_calibFunc = CF_CHECK) Then
    m_chkRefState = 0
    m_smplChkTimeT2.msecs = 0
    m_smplChkTimeT2.mins = 0
    xyPlot.SubsetColor(0) = vbBlack
    lbl_newScanTime.Caption = ""
    clr_plot_scan_data 0
  Else    ' performing update calibration function
    m_smplReferTimeT2.msecs = 0
    m_smplReferTimeT2.mins = 0
    clr_plot_scan_data 1
  End If

  prg_smplScan.percent = 0

  ' Check if reference required
  If (((m_calibFunc = CF_CHECK) And ((m_refChkTaken = False) Or (unity_main.chk_timeout(m_refChkTimeStamp, REF_EXPIRE_TIME) = True))) Or _
      ((m_calibFunc = CF_RECALIBRATE) And ((m_refUpdTaken = False) Or (unity_main.chk_timeout(m_refUpdTimeStamp, REF_EXPIRE_TIME) = True)))) Then
    ' Setup reference scan configuration variables
    If (setup_ref_scan = False) Then
      Exit Sub
    End If
  Else
    ' Setup sample scan configuration variables
    setup_smpl_scan
  End If
  
  ' Start scanning timer
  m_scanState = CSS_CHK_COMP
  tmr_scan.enabled = True
End Sub

Private Sub cmd_updateCalib_Click()

  unity_main.errorstring = "Internal Reference Calibration Management screen Update Calibration button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  ' Hide Update Calibration button
  cmd_updateCalib.Visible = False

  ' Replace spectra in correction file with new sampled spectra
  If (update_correction_file(m_smplUpdReferVals, m_smplReferTimeT2.mins, m_smplReferTimeT2.msecs, m_smplNumReferScans, m_smplUpd2ndOrderVals, m_smpl2ndOrderTimeT2.mins, m_smpl2ndOrderTimeT2.msecs, m_smplNum2ndOrderScans) = True) Then
    m_smplReady2Save = False

    clr_plot_scan_data 1
    plot_scan_data m_smplUpdReferVals, 0, VarPtr(m_smplReferTimeT2)
    CWrap.ShowMessageBoxW "Spectra correction file updated!"

    If (update_history_file(m_smplUpdReferVals, VarPtr(m_smplReferTimeT2), m_smplNumReferScans, m_smplUpd2ndOrderVals, VarPtr(m_smpl2ndOrderTimeT2), m_smplNum2ndOrderScans) = True) Then
      Exit Sub
    End If
  End If

  display_err_msg
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me

  ' Assign menu option text and attach menu to form
  CheckMenu.Caption = CWrap.StrToHex(MLSupport.GSS("frm_intRefCalMgmt", "CheckMenuOpt", "Check Reference"))
  HistoryMenu.Caption = CWrap.StrToHex(MLSupport.GSS("frm_intRefCalMgmt", "HistoryMenuOpt", "Calibration History"))
  RecalibrateMenu.Caption = CWrap.StrToHex(MLSupport.GSS("frm_intRefCalMgmt", "RecalibrateMenuOpt", "Recalibrate Reference"))
  ctlUTF8Menu.CustomFontFace = "Arial Unicode MS"
  ctlUTF8Menu.Attach Me.hWnd

  CheckMenu.enabled = False
  RecalibrateMenu.enabled = False
  HistoryMenu.enabled = False

  cmd_abortScan.Visible = False
  cmd_restoreCalib.Visible = False
  cmd_updateCalib.Visible = False
  cmd_startScan.Visible = False
  lst_spxHistory.Visible = False
  
  m_spxCorFile = StrConv(MS11CfgData.spxCorFile, vbFromUnicode)
  
  ' Save backup of current internal reference scan parameters
  m_intRefNScans = unity_main.m_intRefNScans
  m_intRefPPT = unity_main.m_intRefPPT
  
  unity_main.m_intRefCalFlg = True
  
  If (setup_spx_if = True) Then
    CheckMenu_Click
    
    ' Display SpectraStar info
    lbl_serNum.Caption = MS11CfgData.sysSerialNum
    lbl_minWvln.Caption = Format(unity_main.m_minWvln, "####.0")
    lbl_maxWvln.Caption = Format(unity_main.m_maxWvln, "####.0")
    lbl_spxCorrFile.Caption = m_spxCorFile
  End If
End Sub

Private Sub HistoryMenu_Click()

  unity_main.errorstring = "Internal Reference Calibration Management screen Calibration History menu option selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  ' Disable menu items
  HistoryMenu.enabled = False

  ' Hide buttons
  cmd_restoreCalib.Visible = False
  cmd_updateCalib.Visible = False
  cmd_startScan.Visible = False

  ' Hide Frames
  frame_ref.Visible = False
  frame_smpl.Visible = False

  ' Enable menu items
  CheckMenu.enabled = True
  RecalibrateMenu.enabled = True

  ' Display List Box
  lst_spxHistory.Visible = True

  ' Check if any error loading in spectra history file
  If (load_history_file() = False) Then
    display_err_msg
  Else
    m_calibFunc = CF_HISTORY   ' calibration history function
    setup_plot
    plot_scan_data m_corrReferVals, 0, VarPtr(m_corrReferTimeT2)
    clr_plot_scan_data 1
    clr_plot_scan_data 2
    clr_plot_scan_data 3
  End If
End Sub

Private Sub lst_spxHistory_Click()
  Dim indx As Integer

  indx = lst_spxHistory.ListIndex

  If (indx >= 0) Then
    ' Display Restore Calibration button
    cmd_restoreCalib.Visible = True

#If SSRCS Then
    SSRCSClientError = unity_main.SSRCSClient.GetSpxHistorySpectra(indx, unity_main.m_sys2500, m_numPts, m_histReferVals(0), VarPtr(m_histReferTimeT2), m_histNumReferScans, m_hist2ndOrderVals(0), VarPtr(m_hist2ndOrderTimeT2), m_histNum2ndOrderScans)
    
    If (SSRCSClientError = 0) Then
#Else
    If (SpxFileIF1.GetHistorySpectra(indx, unity_main.m_sys2500, m_histReferVals(0), VarPtr(m_histReferTimeT2), m_histNumReferScans, m_hist2ndOrderVals(0), VarPtr(m_hist2ndOrderTimeT2), m_histNum2ndOrderScans) = True) Then
#End If
      plot_scan_data m_histReferVals, 1, VarPtr(m_histReferTimeT2)
    Else
      build_spx_err_msg
      display_err_msg
    End If
  End If
End Sub

Private Sub RecalibrateMenu_Click()

  unity_main.errorstring = "Internal Reference Calibration Management screen Recalibrate Reference menu option selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  ' Disable menu items
  RecalibrateMenu.enabled = False

  ' Hide buttons
  cmd_restoreCalib.Visible = False

  ' Hide List Box
  lst_spxHistory.Visible = False

  ' Enable menu items
  CheckMenu.enabled = True
  HistoryMenu.enabled = True

  ' Display buttons
  cmd_startScan.Visible = True

  ' Display Frames
  frame_ref.Visible = True
  frame_smpl.Visible = True

  m_calibFunc = CF_RECALIBRATE    ' recalibrate reference function
  setup_plot
  plot_scan_data m_corrReferVals, 0, VarPtr(m_corrReferTimeT2)
  clr_plot_scan_data 2
  clr_plot_scan_data 3

  m_refNumScans = NUM_RECAL_SCANS
  m_smplNumReferScans = NUM_RECAL_SCANS
  m_smplNum2ndOrderScans = NUM_RECAL_SCANS
  lbl_numRefScans.Caption = m_refNumScans
  lbl_numSmplScans.Caption = m_smplNumReferScans

  ' Check if have new sample scan not saved
  If (m_smplReady2Save = True) Then
    plot_scan_data m_smplUpdReferVals, 1, VarPtr(m_smplReferTimeT2)
    
    ' Display Update Calibration button
    cmd_updateCalib.Visible = True
  Else
    If ((m_refUpdTaken = True) And (unity_main.chk_timeout(m_refUpdTimeStamp, REF_EXPIRE_TIME) = False)) Then
      prg_refScan.percent = 100
    Else
      prg_refScan.percent = 0
    End If

    clr_plot_scan_data 1
  End If
End Sub

Private Sub tmr_scan_Timer()
    
    ' Process scanning state
    Select Case (m_scanState)
      Case CSS_CHK_COMP             ' check if scan completed
        Select Case (unity_main.m_scanTmrState)
          Case STS_COMPLETED          ' reference/sample scan completed
            m_scanFlg = False

            ' Hide Abort button
            cmd_abortScan.Visible = False
            m_scanState = CSS_GET_SPECTRUM

          Case STS_ABORT              ' scan was aborted due to user request or error
            unity_main.tmr_ref.enabled = False
            unity_main.tmr_sample.enabled = False
            m_scanFlg = False
            
            If (m_scanAbortFlg = True) Then
              m_scanState = CSS_ABORT
            Else
              m_scanState = CSS_ERROR
            End If
        End Select

      Case CSS_GET_SPECTRUM         ' get scan spectrum
        ' Get sample scan data
        If (get_scan_data = True) Then
          ' Check if reference scan
          If (m_refFlag = True) Then
            ' Check if performing reference function
            If (m_calibFunc = CF_CHECK) Then
              m_refChkTaken = True
            Else    ' performing update calibration function
              m_refUpdTaken = True
            End If

            ' Setup and start sample scan of Fluorion
            m_smplNum = 1
            setup_smpl_scan
            m_scanState = CSS_CHK_COMP
          Else
            ' Check if 2500 system and performing recalibrate reference function
            If ((unity_main.m_sys2500 = True) And (m_calibFunc = CF_RECALIBRATE)) Then
              m_smplNum = m_smplNum + 1

              ' Check if need to sample flour standard
              If (m_smplNum = 2) Then
                setup_smpl_scan
                m_scanState = CSS_CHK_COMP
              Else
                m_scanState = CSS_TREAT_SPECTRUM
              End If
            Else
              m_scanState = CSS_PLOT_SPECTRUM
            End If
          End If
        Else
          m_scanState = CSS_ERROR
        End If

      Case CSS_TREAT_SPECTRUM        ' apply 2500 math treatment to spectrum
        Dim rc1 As Boolean
        Dim rc2 As Boolean

        ' Apply math treatment to sampled 2500 spectra
        rc1 = apply_2500_math_treatment
    
        If (rc1 = True) Then
          ' Get math treated spectra
          rc1 = load_2500_recal_file
        End If
        
        ' Delete all temp files used to perform 2500 math treatment
        rc2 = delete_2500_files
        
        If ((rc1 = False) Or (rc2 = False)) Then
          m_scanState = CSS_ERROR
        Else
          m_scanState = CSS_PLOT_SPECTRUM
        End If

      Case CSS_PLOT_SPECTRUM         ' plot scan spectrum
        ' Check if performing check reference function
        If (m_calibFunc = CF_CHECK) Then
          check_reference

          Select Case (m_chkRefState)
            Case -1                ' check reference failed
              xyPlot.SubsetColor(0) = vbRed
              lbl_newScanTime.ForeColor = vbRed
              lbl_newScanTime.Caption = Chr(34) & "FAILED" & Chr(34)

            Case 0                 ' no check reference
              xyPlot.SubsetColor(0) = vbBlack
              lbl_newScanTime.ForeColor = vbBlack
              lbl_newScanTime.Caption = ""

            Case 1                 ' check reference passed
              xyPlot.SubsetColor(0) = RGB(0, 128, 0)  ' dark green
              lbl_newScanTime.ForeColor = RGB(0, 128, 0)  ' dark green
              lbl_newScanTime.Caption = Chr(34) & "PASSED" & Chr(34)
          End Select

          ' Plot scanned spectrum
          plot_scan_data m_smplChkReferVals, 0, VarPtr(m_smplChkTimeT2)
        Else    ' performing update calibration function
          ' Plot scanned spectrum
          plot_scan_data m_smplUpdReferVals, 1, VarPtr(m_smplReferTimeT2)
        End If

        m_scanState = CSS_DONE

      Case CSS_DONE                 ' scanning completed
        tmr_scan.enabled = False
        prg_refScan.percent = 0
        prg_smplScan.percent = 0
        m_smplNum = 1

        ' Enable menu items
        HistoryMenu.enabled = True

        ' Check if performing check reference function
        If (m_calibFunc = CF_CHECK) Then
          ' Check if to reset internal reference verification timer
          If (unity_main.m_intRefVerifyTimeout > 0) Then
            unity_main.m_intRefVerifyTimer = Timer
            unity_main.m_intRefVerReminderCtr = 0
            unity_main.m_dblAccumTime = 0#
          End If
          
          ' Hide Verify Ref button
          unity_main.disp_verify_ref_button False
          
          RecalibrateMenu.enabled = True
        Else    ' performing update calibration function
          CheckMenu.enabled = True
        End If

        ' Display buttons
        cmd_exit.Visible = True
        cmd_startScan.Visible = True

        ' Check if performing recalibrate reference function
        If (m_calibFunc = CF_RECALIBRATE) Then
          ' Flag sample ready to be save to correction file
          m_smplReady2Save = True

          ' Display Update Calibration button
          cmd_updateCalib.Visible = True
        End If

      Case CSS_ERROR                ' scan error
        tmr_scan.enabled = False
        m_smplNum = 1

        ' Hide Abort button
        cmd_abortScan.Visible = False

        ' Enable menu items
        HistoryMenu.enabled = True

        ' Check if performing check reference function
        If (m_calibFunc = CF_CHECK) Then
          RecalibrateMenu.enabled = True
        Else    ' performing update calibration function
          CheckMenu.enabled = True
        End If

        ' Display buttons
        cmd_exit.Visible = True
        cmd_startScan.Visible = True
        CWrap.ShowMessageBoxW unity_main.m_uniErrMsg, vbCritical

      Case CSS_ABORT                ' scan aborted
        tmr_scan.enabled = False
        m_smplNum = 1

        ' Enable menu items
        HistoryMenu.enabled = True

        ' Check if performing check reference function
        If (m_calibFunc = CF_CHECK) Then
          RecalibrateMenu.enabled = True
        Else    ' performing update calibration function
          CheckMenu.enabled = True
        End If

        ' Display buttons
        cmd_exit.Visible = True
        cmd_startScan.Visible = True
        CWrap.ShowMessageBoxW m_uniErrMsg

      Case CSS_SHUTDOWN             ' major scan operation error, shutdown
        tmr_scan.enabled = False
        m_smplNum = 1
        display_err_msg
    End Select
End Sub
