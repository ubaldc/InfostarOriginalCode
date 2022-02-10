VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_Inst 
   Caption         =   "Global System Configuration"
   ClientHeight    =   9075
   ClientLeft      =   360
   ClientTop       =   750
   ClientWidth     =   10905
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   8.25
      Charset         =   0
      Weight          =   900
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Inst.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniLabel lbl_refTimeout 
      Height          =   590
      Left            =   5580
      Top             =   840
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_Inst.frx":058A
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_Inst.frx":0602
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":0622
   End
   Begin HexUniControls.ctlUniLabel lbl_refPPT 
      Height          =   590
      Left            =   5580
      Top             =   1440
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_Inst.frx":063E
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_Inst.frx":06CE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":06EE
   End
   Begin HexUniControls.ctlUniFrameXP frame_autoSmplr 
      Height          =   2535
      Left            =   6120
      Top             =   4800
      Width           =   4455
      _ExtentX        =   7858
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
      Caption         =   "frm_Inst.frx":070A
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_Inst.frx":075E
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":077E
      Begin HexUniControls.ctlUniCheckXP chk_useAutoSmplr 
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "frm_Inst.frx":079A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_Inst.frx":07E2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":0802
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniLabel lbl_batchRptPath 
         Height          =   300
         Left            =   120
         Top             =   840
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":081E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_Inst.frx":0884
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":08A4
      End
      Begin HexUniControls.ctlUniTextBoxXP txt_batchRptPath 
         Height          =   450
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   794
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frm_Inst.frx":08C0
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
         Tip             =   "frm_Inst.frx":0902
         NoHideSel       =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":0922
      End
      Begin HexUniControls.ctlNumIncXP numInc_commPort 
         Height          =   600
         Left            =   1440
         TabIndex        =   38
         Top             =   1800
         Width           =   1095
         _ExtentX        =   1931
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
         Text            =   "3"
         Min             =   1
         Max             =   4
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
         MouseIcon       =   "frm_Inst.frx":093E
         TrapTabKey      =   0   'False
      End
      Begin HexUniControls.ctlUniLabel lbl_commPort 
         Height          =   600
         Left            =   120
         Top             =   1800
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1058
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":095A
         BackColor       =   -2147483633
         ForeColor       =   0
         Alignment       =   0
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_Inst.frx":098C
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":09AC
      End
   End
   Begin HexUniControls.ctlUniFrameXP frame_reference 
      Height          =   1575
      Left            =   240
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2778
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_Inst.frx":09C8
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_Inst.frx":09FA
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":0A1A
      Begin HexUniControls.ctlUniLabel lbl_refTempDiff 
         Height          =   945
         Left            =   120
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1667
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":0A36
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_Inst.frx":0AA6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":0AC6
      End
      Begin HexUniControls.ctlUniLabel ctlUniLabel8 
         Height          =   345
         Left            =   3120
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":0AE2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_Inst.frx":0B04
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":0B24
      End
      Begin HexUniControls.ctlUniLabel ctlUniLabel1 
         Height          =   345
         Left            =   3000
         Top             =   600
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":0B40
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_Inst.frx":0B62
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":0B82
      End
      Begin HexUniControls.ctlNumIncXP numInc_refTempDiff 
         Height          =   600
         Left            =   1800
         TabIndex        =   39
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
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
         Text            =   "5"
         Min             =   1
         Max             =   10
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
         MouseIcon       =   "frm_Inst.frx":0B9E
         TrapTabKey      =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cfgSpectTreat 
      Height          =   645
      Left            =   6600
      TabIndex        =   35
      Top             =   8160
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
      Caption         =   "frm_Inst.frx":0BBA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_Inst.frx":0BFE
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":0C1E
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cfgTkt 
      Height          =   645
      Left            =   120
      TabIndex        =   27
      Top             =   8160
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
      Caption         =   "frm_Inst.frx":0C3A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_Inst.frx":0C76
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":0C96
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cfgLIMS 
      Height          =   645
      Left            =   120
      TabIndex        =   26
      Top             =   7440
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
      Caption         =   "frm_Inst.frx":0CB2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_Inst.frx":0CDA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":0CFA
   End
   Begin HexUniControls.ctlUniCheckXP chk_darkSub 
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_Inst.frx":0D16
      Enabled         =   0   'False
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_Inst.frx":0D5A
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":0D7A
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniCheckXP chk_allowbias 
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   4335
      _ExtentX        =   7646
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
      Caption         =   "frm_Inst.frx":0D96
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_Inst.frx":0DEA
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":0E0A
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cfgInputs 
      Height          =   645
      Left            =   2280
      TabIndex        =   28
      Top             =   7440
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
      Caption         =   "frm_Inst.frx":0E26
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_Inst.frx":0E5C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":0E7C
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_superPW 
      Height          =   645
      Left            =   2280
      TabIndex        =   29
      Top             =   8160
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
      Caption         =   "frm_Inst.frx":0E98
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_Inst.frx":0EDE
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":0EFE
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8760
      Top             =   8760
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9075
      FormDesignWidth =   10905
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   645
      Left            =   8760
      TabIndex        =   33
      Top             =   7440
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
      Caption         =   "frm_Inst.frx":0F1A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_Inst.frx":0F52
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":0F72
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   645
      Left            =   8760
      TabIndex        =   0
      Top             =   8160
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
      Caption         =   "frm_Inst.frx":0F8E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_Inst.frx":0FBA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":0FDA
   End
   Begin HexUniControls.ctlUniFrameXP Frame1 
      Height          =   735
      Left            =   1080
      Top             =   60
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_Inst.frx":0FF6
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_Inst.frx":1032
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":1052
      Begin HexUniControls.ctlUniRadioXP opt_scanNone 
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   2000
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":106E
         Enabled         =   0   'False
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_Inst.frx":1096
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":10B6
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniRadioXP opt_scanBoth 
         Height          =   300
         Left            =   6420
         TabIndex        =   4
         Top             =   330
         Width           =   2000
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":10D2
         Enabled         =   0   'False
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_Inst.frx":10FA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":111A
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniRadioXP opt_scanDown 
         Height          =   300
         Left            =   2220
         TabIndex        =   2
         Top             =   330
         Width           =   2000
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":1136
         Enabled         =   0   'False
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_Inst.frx":115E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":117E
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniRadioXP opt_scanUp 
         Height          =   300
         Left            =   4320
         TabIndex        =   3
         Top             =   330
         Width           =   2000
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":119A
         Enabled         =   0   'False
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_Inst.frx":11BE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":11DE
         ShowFocus       =   -1  'True
      End
   End
   Begin HexUniControls.ctlNumIncXP numInc_refTimeout 
      Height          =   600
      Left            =   9345
      TabIndex        =   22
      Top             =   840
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
      Text            =   "2880"
      Min             =   0
      Max             =   2880
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
      MouseIcon       =   "frm_Inst.frx":11FA
      TrapTabKey      =   0   'False
   End
   Begin HexUniControls.ctlNumIncXP numInc_refPPT 
      Height          =   600
      Left            =   9360
      TabIndex        =   23
      Top             =   1440
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
      Text            =   "300"
      Min             =   0
      Max             =   300
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
      MouseIcon       =   "frm_Inst.frx":1216
      TrapTabKey      =   0   'False
   End
   Begin HexUniControls.ctlNumIncXP numInc_refNScans 
      Height          =   600
      Left            =   9360
      TabIndex        =   24
      Top             =   2040
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
      Text            =   "24"
      Min             =   1
      Max             =   512
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
      MouseIcon       =   "frm_Inst.frx":1232
      TrapTabKey      =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chk_darkScan 
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_Inst.frx":124E
      Enabled         =   0   'False
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_Inst.frx":128E
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":12AE
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniCheckXP chk_enableRunMode 
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_Inst.frx":12CA
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_Inst.frx":1308
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":1328
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniLabel lbl_refNScans 
      Height          =   590
      Left            =   5580
      Top             =   2040
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_Inst.frx":1344
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_Inst.frx":13B0
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":13D0
   End
   Begin HexUniControls.ctlUniLabel lbl_dbgLevel 
      Height          =   590
      Left            =   5985
      Top             =   2640
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_Inst.frx":13EC
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_Inst.frx":1422
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":1442
   End
   Begin HexUniControls.ctlNumIncXP numInc_dbgLevel 
      Height          =   600
      Left            =   9360
      TabIndex        =   25
      Top             =   2640
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
      Text            =   "3"
      Min             =   0
      Max             =   3
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
      MouseIcon       =   "frm_Inst.frx":145E
      TrapTabKey      =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cfgCSV 
      Height          =   645
      Left            =   4440
      TabIndex        =   30
      Top             =   7440
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
      Caption         =   "frm_Inst.frx":147A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_Inst.frx":14AE
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":14CE
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_extRef 
      Height          =   645
      Left            =   4440
      TabIndex        =   31
      Top             =   8160
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
      Caption         =   "frm_Inst.frx":14EA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_Inst.frx":1530
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":1550
   End
   Begin HexUniControls.ctlUniFrameXP frame_globalSampName 
      Height          =   3825
      Left            =   120
      Top             =   3520
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   6747
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_Inst.frx":156C
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_Inst.frx":15C6
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":15E6
      Begin HexUniControls.ctlUniTextBoxXP txt_GlobalScanName 
         Height          =   450
         Left            =   120
         TabIndex        =   34
         Top             =   3240
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   794
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         Locked          =   0   'False
         Text            =   "frm_Inst.frx":1602
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
         Tip             =   "frm_Inst.frx":1622
         NoHideSel       =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":1642
      End
      Begin HexUniControls.ctlUniLabel Label3 
         Height          =   300
         Left            =   120
         Top             =   2880
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":165E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_Inst.frx":169E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":16BE
      End
      Begin HexUniControls.ctlNumIncXP numInc_globalBaseCtr 
         Height          =   600
         Left            =   3555
         TabIndex        =   13
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
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
         Text            =   "999999999"
         Min             =   0
         Max             =   999999999
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
         MouseIcon       =   "frm_Inst.frx":16DA
         TrapTabKey      =   0   'False
      End
      Begin HexUniControls.ctlNumIncXP numInc_globalDateCtr 
         Height          =   600
         Left            =   3555
         TabIndex        =   11
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
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
         Text            =   "999999"
         Min             =   1
         Max             =   999999
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
         MouseIcon       =   "frm_Inst.frx":16F6
         TrapTabKey      =   0   'False
      End
      Begin HexUniControls.ctlUniTextBoxXP txt_globalSampNameBase 
         Height          =   450
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   794
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frm_Inst.frx":1712
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
         Tip             =   "frm_Inst.frx":174A
         NoHideSel       =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":176A
      End
      Begin HexUniControls.ctlUniLabel Label4 
         Height          =   300
         Left            =   120
         Top             =   2040
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":1786
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_Inst.frx":17C6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":17E6
      End
      Begin HexUniControls.ctlUniRadioXP opt_globalNameBase 
         Height          =   450
         Left            =   120
         TabIndex        =   12
         Top             =   860
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "frm_Inst.frx":1802
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_Inst.frx":1854
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":1874
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniRadioXP opt_globalNameDate 
         Height          =   450
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "frm_Inst.frx":1890
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_Inst.frx":18D8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":18F8
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniRadioXP opt_globalNameCtr 
         Height          =   450
         Left            =   120
         TabIndex        =   14
         Top             =   1360
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "frm_Inst.frx":1914
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_Inst.frx":195E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":197E
         ShowFocus       =   -1  'True
      End
   End
   Begin HexUniControls.ctlUniCheckXP chk_enableGlobalSampName 
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   4335
      _ExtentX        =   7646
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
      Caption         =   "frm_Inst.frx":199A
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_Inst.frx":19EC
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":1A0C
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cfgDynRpt 
      Height          =   645
      Left            =   6600
      TabIndex        =   32
      Top             =   7440
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
      Caption         =   "frm_Inst.frx":1A28
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_Inst.frx":1A64
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":1A84
   End
   Begin HexUniControls.ctlUniLabel lbl_refVerifyTimeout 
      Height          =   590
      Left            =   6000
      Top             =   3240
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_Inst.frx":1AA0
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_Inst.frx":1B18
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":1B38
   End
   Begin HexUniControls.ctlNumIncXP numInc_refVerifyTimeout 
      Height          =   600
      Left            =   9360
      TabIndex        =   40
      Top             =   3240
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
      Text            =   "168"
      Min             =   0
      Max             =   8760
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
      MouseIcon       =   "frm_Inst.frx":1B54
      TrapTabKey      =   0   'False
   End
   Begin HexUniControls.ctlUniLabel lbl_refVerifyRemindTime 
      Height          =   590
      Left            =   6000
      Top             =   3840
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_Inst.frx":1B70
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_Inst.frx":1C08
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":1C28
   End
   Begin HexUniControls.ctlNumIncXP numInc_refVerifyRemindTime 
      Height          =   600
      Left            =   9360
      TabIndex        =   41
      Top             =   3840
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
      Text            =   "5"
      Min             =   1
      Max             =   60
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
      MouseIcon       =   "frm_Inst.frx":1C44
      TrapTabKey      =   0   'False
   End
   Begin HexUniControls.ctlUniFrameXP frame_network 
      Height          =   2295
      Left            =   6840
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4048
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_Inst.frx":1C60
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_Inst.frx":1C8E
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_Inst.frx":1CAE
      Begin HexUniControls.ctlUniLabel lbl_ipAddr 
         Height          =   345
         Left            =   120
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":1CCA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_Inst.frx":1CFE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":1D1E
      End
      Begin HexUniControls.ctlUniLabel lbl_portNum 
         Height          =   345
         Left            =   120
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":1D3A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_Inst.frx":1D70
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":1D90
      End
      Begin HexUniControls.ctlUniLabel lbl_rspTimeout 
         Height          =   600
         Left            =   120
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1058
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":1DAC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   1
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_Inst.frx":1DE4
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":1E04
      End
      Begin HexUniControls.ctlUniLabel lbl_sec 
         Height          =   345
         Left            =   2880
         Top             =   1560
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":1E20
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_Inst.frx":1E48
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":1E68
      End
      Begin HexUniControls.ctlUniTextBoxXP txt_portNum 
         Height          =   390
         Left            =   1680
         TabIndex        =   20
         Top             =   960
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   688
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frm_Inst.frx":1E84
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
         Alignment       =   1
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frm_Inst.frx":1EA4
         NoHideSel       =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":1EC4
      End
      Begin HexUniControls.ctlUniTextBoxXP txt_ipAddr 
         Height          =   390
         Index           =   1
         Left            =   1680
         TabIndex        =   16
         Top             =   480
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frm_Inst.frx":1EE0
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
         Alignment       =   1
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frm_Inst.frx":1F00
         NoHideSel       =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":1F20
      End
      Begin HexUniControls.ctlUniTextBoxXP txt_ipAddr 
         Height          =   390
         Index           =   2
         Left            =   2160
         TabIndex        =   17
         Top             =   480
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frm_Inst.frx":1F3C
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
         Alignment       =   1
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frm_Inst.frx":1F5C
         NoHideSel       =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":1F7C
      End
      Begin HexUniControls.ctlUniTextBoxXP txt_ipAddr 
         Height          =   390
         Index           =   3
         Left            =   2640
         TabIndex        =   18
         Top             =   480
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frm_Inst.frx":1F98
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
         Alignment       =   1
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frm_Inst.frx":1FB8
         NoHideSel       =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":1FD8
      End
      Begin HexUniControls.ctlUniTextBoxXP txt_ipAddr 
         Height          =   390
         Index           =   4
         Left            =   3120
         TabIndex        =   19
         Top             =   480
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frm_Inst.frx":1FF4
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
         Alignment       =   1
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frm_Inst.frx":2014
         NoHideSel       =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":2034
      End
      Begin HexUniControls.ctlNumIncXP numInc_rspTimeout 
         Height          =   600
         Left            =   1680
         TabIndex        =   21
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
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
         Text            =   "12"
         Min             =   5
         Max             =   20
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
         MouseIcon       =   "frm_Inst.frx":2050
         TrapTabKey      =   0   'False
      End
      Begin HexUniControls.ctlUniLabel Label1 
         Height          =   255
         Index           =   2
         Left            =   3000
         Top             =   600
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":206C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_Inst.frx":208E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":20AE
      End
      Begin HexUniControls.ctlUniLabel Label1 
         Height          =   255
         Index           =   1
         Left            =   2520
         Top             =   600
         Width           =   195
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":20CA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_Inst.frx":20EC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":210C
      End
      Begin HexUniControls.ctlUniLabel Label1 
         Height          =   255
         Index           =   0
         Left            =   2040
         Top             =   600
         Width           =   200
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_Inst.frx":2128
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_Inst.frx":214A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_Inst.frx":216A
      End
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   7080
      Top             =   8760
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
      Left            =   8040
      Top             =   8760
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_Inst.frx":2186
   End
End
Attribute VB_Name = "frm_Inst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_saveGlobalIni As Boolean

Private m_badIniVal As Boolean

Private m_allowBias As Boolean
#If SSTAR Then
Private m_autoSmplrPort As Integer
Private m_batchRptPath As String
#End If
Private m_darkScan As Boolean
Private m_darkSub As Boolean
Private m_dbgLevel As LOG_DBG_LEVELS
Private m_enableGlobalName As Boolean
Private m_enableRunMode As Boolean
Private m_globalBaseCounter As Long
Private m_globalDate As String
Private m_globalDateCounter As Long
Private m_globalNameBase As String
Private m_globalNameMode As String
Private m_intRefNScans As Integer
Private m_intRefPPT As Integer
Private m_intRefTimeout As Integer
Private m_intRefVerifyRemindTime As Integer
Private m_intRefVerifyTimeout As Integer
Private m_intRefVerifyAccumTime As Integer
Private m_ipAddr As String
Private m_portNum As Long
Private m_refTempDiff As Integer
Private m_rspTimeout As Long
Private m_saveInstsFail As Boolean

#If SSTAR Then
Private m_scanDir As SCAN_DIRECTIONS
Private m_useAutoSmplr As Boolean
#End If

#If ABBFT Then
Private m_commChanged As Boolean
#End If

#If SSTAR Then
Private Const DFLT_INT_REF_VERIFY_REMIND_TIME = 5       ' 5 minutes
Private Const DFLT_INT_REF_VERIFY_TIME = (24 * 7)       ' 1 week in hours
#End If

Public Sub savemyinsts(saveBtn As Boolean, chkPPT As Boolean)
  Dim ii As Integer
  Dim ipAddr(1 To 4) As Integer
  Dim filePathName As String
  Dim errMsg As String
  Dim uniFile As New clsUniFile
  Dim uniMsg As String
  
  ' Check if config files needs to be saved due to user request or global name counter update
  If (saveBtn = True) Or (m_saveGlobalIni = True) Then
    m_saveGlobalIni = False
    
#If ABBFT Then
    ' Save ABB Interferometer variables
    m_refTempDiff = numInc_refTempDiff.Text
    
    For ii = 1 To 4
      ipAddr(ii) = txt_ipAddr(ii).Text
    Next ii
  
    m_ipAddr = (ipAddr(1) & "." & ipAddr(2) & "." & ipAddr(3) & "." & ipAddr(4))
    m_portNum = txt_portNum.Text
    m_rspTimeout = numInc_rspTimeout.Text
#Else
    ' Save SpectraStar variables
    m_darkScan = CBool(chk_darkScan.Value)
    m_darkSub = CBool(chk_darkSub.Value)
    m_intRefNScans = numInc_refNScans.Text
    m_intRefPPT = numInc_refPPT.Text
    m_intRefTimeout = numInc_refTimeout.Text
    m_intRefVerifyRemindTime = numInc_refVerifyRemindTime.Text
    m_intRefVerifyTimeout = numInc_refVerifyTimeout.Text
    
    m_saveInstsFail = False
   
    If CDbl(m_intRefVerifyTimeout) > 744# Then
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_Inst.errMsg5", "Value must be between -1 and 744"), vbOKOnly
      m_saveInstsFail = True
      Exit Sub
    End If
    ' Save auto-sampler comm port
    m_autoSmplrPort = numInc_commPort.Text
    
    ' Confirm report filepath contains '\' instead of '/'
    filePathName = Trim(txt_batchRptPath.Text)
    check_filepathname_delimiters filePathName
    txt_batchRptPath.Text = filePathName
    
    ' Append "\" to report file path if not present
    If (Right(txt_batchRptPath.Text, 1) <> "\") Then
      txt_batchRptPath.Text = txt_batchRptPath.Text & "\"
    End If
    
    ' Save batch report folder
    m_batchRptPath = txt_batchRptPath.Text
    
    ' Save scan direction
    If (opt_scanNone.Value = True) Then
      m_scanDir = SD_NONE
    Else
      If (opt_scanDown.Value = True) Then
        m_scanDir = SD_DOWN
      Else
        If (opt_scanUp.Value = True) Then
          m_scanDir = SD_UP
        Else
          If (opt_scanBoth.Value = True) Then
            m_scanDir = SD_BOTH
          End If
        End If
      End If
    End If
    
    ' Save utilize auto-sampler
    If (chk_useAutoSmplr.Value = 1) Then
      m_useAutoSmplr = True
    Else
      m_useAutoSmplr = False
    End If
#End If

    ' Save InfoStar variables
    m_allowBias = CBool(chk_allowbias.Value)
    m_dbgLevel = numInc_dbgLevel.Text
    m_enableGlobalName = CBool(chk_enableGlobalSampName.Value)
    m_enableRunMode = CBool(chk_enableRunMode.Value)
    m_globalBaseCounter = numInc_globalBaseCtr.Text
    m_globalDate = unity_main.m_globalDate
    m_globalDateCounter = numInc_globalDateCtr.Text

    ' Save sample base file name
    m_globalNameBase = Trim(txt_globalSampNameBase.Text)
  
#If SSTAR Then
    ' Check if current product using internal reference
    If (unity_main.m_bType = "internal") Then
      ' Check if reference performed on demand
      If (unity_main.m_backFreq = REF_FREQ_ON_DEMAND) Then
        ' Restart instrument's reference timer
#If SSRCS Then
        SSRCSClientError = unity_main.SSRCSClient.SetRefTimeout(m_intRefTimeout)
#Else
        unity_main.MS11srv.refTimeout = m_intRefTimeout
#End If
      Else
        ' Stop instrument's reference timer
#If SSRCS Then
        SSRCSClientError = unity_main.SSRCSClient.SetRefTimeout(0)
#Else
        unity_main.MS11srv.refTimeout = 0
#End If
      End If
  
      ' Check if internal reference qualification required
      If (chkPPT = True) And (m_intRefPPT <> 0) Then
        Dim spcFilename As String
      
        ' Check if do not have internal reference qualification file for wavelength range
        If (unity_main.check_int_ref_ppt_file(unity_main.m_smplStartWvln, unity_main.m_smplEndWvln, spcFilename) = False) Then
          unity_main.errorstring = (spcFilename & " file cannot be found. Need to collect reference qualification data")
          unity_main.write_error
          CWrap.ShowMessageBoxW MLSupport.GGS_Params("unity_main.errMsg4", "%1 file cannot be found. Need to collect internal reference qualification data over wavelength range %2 - %3", spcFilename, unity_main.m_smplStartWvln, unity_main.m_smplEndWvln), vbOKOnly
          unity_main.m_intRefPPTScan = True
          unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status6", "Reference Qualification Required")
        End If
      End If
    End If
#End If

    If (uniFile.OpenFileWrite(CFG_DIR & INST_CFG_FILE) = True) Then
      On Error GoTo FILE_ERROR
      uniFile.WriteBOM fe_UTF16LE
      uniFile.WriteUnicodeLine "[signature settings]"
#If ABBFT Then
      uniFile.WriteUnicodeLine ("DevID=" & DTID_ABBFT)
      uniFile.WriteUnicodeLine ("Version=" & INFOSTAR_VER)
      uniFile.WriteUnicodeLine "[global settings]"
      uniFile.WriteUnicodeLine ("DbgLevel=" & m_dbgLevel)
      uniFile.WriteUnicodeLine ("EnableGlobalName=" & m_enableGlobalName)
      uniFile.WriteUnicodeLine ("GlobalBaseCtr=" & m_globalBaseCounter)
      uniFile.WriteUnicodeLine ("GlobalDate=" & m_globalDate)
      uniFile.WriteUnicodeLine ("GlobalDateCtr=" & m_globalDateCounter)
      uniFile.WriteUnicodeLine ("GlobalNameBase=" & m_globalNameBase)
      uniFile.WriteUnicodeLine ("GlobalNameMode=" & m_globalNameMode)
      uniFile.WriteUnicodeLine ("IPAddr=" & m_ipAddr)
      uniFile.WriteUnicodeLine ("MinGui=" & chk_enableRunMode.Value)
      uniFile.WriteUnicodeLine ("PortNum=" & m_portNum)
      uniFile.WriteUnicodeLine ("RefTempDiff=" & m_refTempDiff)
      uniFile.WriteUnicodeLine ("RspTimeout=" & m_rspTimeout)
      uniFile.WriteUnicodeLine ("RunTimeBias=" & chk_allowbias.Value)
#Else
      uniFile.WriteUnicodeLine ("DevID=" & MS11CfgData.devID)
      uniFile.WriteUnicodeLine ("SmplTable=" & unity_main.m_smplTable)
      uniFile.WriteUnicodeLine ("ScanMode=" & unity_main.m_sysScanMode)
      uniFile.WriteUnicodeLine ("Version=" & INFOSTAR_VER)
      uniFile.WriteUnicodeLine "[global settings]"
      uniFile.WriteUnicodeLine ("AutoSmplrPort=" & m_autoSmplrPort)
      uniFile.WriteUnicodeLine ("BatchRptPath=" & m_batchRptPath)
      uniFile.WriteUnicodeLine ("DbgLevel=" & m_dbgLevel)
      uniFile.WriteUnicodeLine ("DarkScan=" & chk_darkScan.Value)
      uniFile.WriteUnicodeLine ("DarkSub=" & chk_darkSub.Value)
      uniFile.WriteUnicodeLine ("EnableGlobalName=" & m_enableGlobalName)
      uniFile.WriteUnicodeLine ("GlobalBaseCtr=" & m_globalBaseCounter)
      uniFile.WriteUnicodeLine ("GlobalDate=" & m_globalDate)
      uniFile.WriteUnicodeLine ("GlobalDateCtr=" & m_globalDateCounter)
      uniFile.WriteUnicodeLine ("GlobalNameBase=" & m_globalNameBase)
      uniFile.WriteUnicodeLine ("GlobalNameMode=" & m_globalNameMode)
      uniFile.WriteUnicodeLine ("MinGui=" & chk_enableRunMode.Value)
      uniFile.WriteUnicodeLine ("RefNScans=" & m_intRefNScans)
      uniFile.WriteUnicodeLine ("RefPPT=" & m_intRefPPT)
      uniFile.WriteUnicodeLine ("RefTimeout=" & m_intRefTimeout)
      uniFile.WriteUnicodeLine ("RefVerifyRemindTime=" & m_intRefVerifyRemindTime)
      uniFile.WriteUnicodeLine ("RefVerifyTimeout=" & m_intRefVerifyTimeout)
      uniFile.WriteUnicodeLine ("RunTimeBias=" & chk_allowbias.Value)
      uniFile.WriteUnicodeLine ("ScanDir=" & m_scanDir)
      uniFile.WriteUnicodeLine ("UseAutoSmplr=" & m_useAutoSmplr)
      'uniFile.WriteUnicodeLine ("RefVerifyAccumTime=" & m_intRefVerifyAccumTime)
#End If

      uniFile.Flush
    
#If ABBFT Then
      ' Copy global system config into system operational variables
      unity_main.m_allowBias = m_allowBias
      unity_main.m_dbgLevel = m_dbgLevel
      unity_main.m_enableGlobalName = m_enableGlobalName
      unity_main.m_enableRunMode = m_enableRunMode
      unity_main.m_globalDate = m_globalDate
      unity_main.m_globalDateCounter = m_globalDateCounter
      unity_main.m_globalNameBase = m_globalNameBase
      unity_main.m_globalBaseCounter = m_globalBaseCounter
      unity_main.m_globalNameMode = m_globalNameMode
      unity_main.m_mb3000.m_ipAddr = m_ipAddr
      unity_main.m_mb3000.m_portNum = m_portNum
      unity_main.m_mb3000.m_refTempDiff = m_refTempDiff
      unity_main.m_mb3000.m_rspTimeout = m_rspTimeout
      
      ' Check if any interferometer communication parameters changed
      If (m_commChanged = True) Then
        ' Flag interferometer comms needs to initialized
        unity_main.m_mb3000.m_connStatus = False
      End If
#Else
      ' Copy global system config into system operational variables
      unity_main.m_allowBias = m_allowBias
      unity_main.m_autoSmplrPort = m_autoSmplrPort
      unity_main.m_batchRptPath = m_batchRptPath
      unity_main.m_darkScan = m_darkScan
      unity_main.m_darkSub = m_darkSub
      unity_main.m_dbgLevel = m_dbgLevel
      unity_main.m_enableGlobalName = m_enableGlobalName
      unity_main.m_enableRunMode = m_enableRunMode
      unity_main.m_globalDate = m_globalDate
      unity_main.m_globalDateCounter = m_globalDateCounter
      unity_main.m_globalNameBase = m_globalNameBase
      unity_main.m_globalBaseCounter = m_globalBaseCounter
      unity_main.m_globalNameMode = m_globalNameMode
      unity_main.m_intRefNScans = m_intRefNScans
      unity_main.m_intRefPPT = m_intRefPPT
      unity_main.m_intRefTimeout = m_intRefTimeout
      unity_main.m_intRefVerifyRemindTime = m_intRefVerifyRemindTime
      unity_main.m_intRefVerifyTimeout = m_intRefVerifyTimeout
      unity_main.m_scanDir = m_scanDir
      unity_main.m_useAutoSmplr = m_useAutoSmplr
      unity_main.m_intRefVerifyAccumTime = m_intRefVerifyAccumTime
      NumAutoSmplrTowers = 0
      tempNumAutoSmplrTowers = 0
  
      ' Setup batch scan buttons if auto-sampler in use
      If (m_useAutoSmplr = True) Then
        unity_main.cmd_runBatch.Visible = True
        unity_main.cmd_sample.Width = unity_main.cmd_runBatch.Width
        frmUtils.cmd_cfgScanBatch.Visible = True

        ' Initialize serial port communication w/ auto-sampler
#If SSRCS Then
        Dim parity As String
        parity = AUTO_SMPLR_PARITY
        SSRCSClientError = unity_main.SSRCSClient.InitASComms(unity_main.m_autoSmplrPort, AUTO_SMPLR_BAUD, parity, AUTO_SMPLR_RSP_TIMEOUT, AUTO_SMPLR_NUM_RETRIES, AUTO_SMPLR_CRC_USAGE)

        If (SSRCSClientError = 0) Then
          SSRCSClientError = unity_main.SSRCSClient.GetAllASTubesState
        Else
          If (unity_main.m_ssrcsConnected = True) Then
            CWrap.ShowMessageBoxW MLSupport.GGS_Params("ASErrMsg1", "Problem with configuring Auto-Sampler comm port %1. Verify no other communication option is configured to use this port", CStr(unity_main.m_autoSmplrPort)), vbCritical
          Else
            CWrap.ShowMessageBoxW MLSupport.GSS("frm_ssrcsConnect", "errMsg3", "Not connected to a SpectraStar RCS server"), vbCritical
          End If
        End If
#Else
        If (frm_batchRun.autoSmplr.init_comms(unity_main.m_autoSmplrPort, AUTO_SMPLR_BAUD, AUTO_SMPLR_PARITY, AUTO_SMPLR_RSP_TIMEOUT, AUTO_SMPLR_NUM_RETRIES) = True) Then
          ' Get max number of tubes supported
          frm_batchRun.autoSmplr.get_all_tubes_state AUTO_SMPLR_CRC_USAGE
        Else
          CWrap.ShowMessageBoxW MLSupport.GGS_Params("ASErrMsg1", "Problem with configuring Auto-Sampler comm port %1. Verify no other communication option is configured to use this port", CStr(unity_main.m_autoSmplrPort)), vbCritical
        End If
#End If
      Else
        unity_main.cmd_runBatch.Visible = False
        unity_main.cmd_sample.Width = unity_main.cmd_select.Width
        frmUtils.cmd_cfgScanBatch.Visible = False
        
        ' Close serial communication w/ auto-sampler if top window system
        If (MS11CfgData.devID = DTID_TOPWIND0) Or (MS11CfgData.devID = DTID_TOPWIND1) Then
#If SSRCS Then
          SSRCSClientError = unity_main.SSRCSClient.CloseASComms
#Else
          frm_batchRun.autoSmplr.close_comms
#End If
        End If
      End If
      
      ' Check if to start internal reference verification timer
      If (unity_main.m_allowIntRefCalAccess = True) Then
        If (unity_main.m_intRefVerifyTimeout > 0) Then
          If (unity_main.m_intRefVerifyTimer = 0) Then
            unity_main.m_intRefVerifyTimer = Timer
          End If
        Else
          unity_main.m_intRefVerifyTimer = 0
        End If
      End If
#End If
    Else
FILE_ERROR:
      errMsg = ((CFG_DIR & INST_CFG_FILE) & " file write error. " & Error$)
      unity_main.errorstring = errMsg
      unity_main.write_error
      uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", (CFG_DIR & INST_CFG_FILE), Error$)
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    End If
  
    uniFile.CloseFile
  End If
End Sub

Public Sub Loadmyinstini()
  Dim fileName As String
  Dim inString As String
#If SSTAR Then
  Dim sysScanDir As SCAN_DIRECTIONS
  Dim sysDarkScan As Integer
  Dim sysDarkSub As Integer
#End If
  Dim lineCnt As Integer
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim rc As Boolean
  Dim ii As Integer
  Dim pos1, pos2 As Integer
  Dim strlen As Integer
  Dim filePathName As String
  Dim errMsg As String
  Dim uniMsg As String
  
  m_badIniVal = False

#If ABBFT Then
  ' Initialize ABB Interferometer values before loading file
  m_ipAddr = ""
  m_portNum = 0
  m_refTempDiff = 5
  m_rspTimeout = 12
#Else
  ' Initialize SpectraStar values before loading file
  sysScanDir = MS11CfgData.sysScanMode And &H3
  sysDarkScan = (MS11CfgData.sysScanMode And &H8) / &H8
  sysDarkSub = (MS11CfgData.sysScanMode And &H4) / &H4
  m_autoSmplrPort = DFLT_AUTO_SMPLR_PORT
  m_batchRptPath = SCAN_BATCHES_RPT_DIR
  m_darkScan = sysDarkScan
  m_darkSub = sysDarkSub
  m_scanDir = sysScanDir
  m_useAutoSmplr = False
    
  unity_main.m_smplTable = 0
  unity_main.m_sysScanMode = -1
  m_intRefNScans = MS11DfltScanCfgData.nScans4Ref
  m_intRefPPT = MS11DfltScanCfgData.ref4PPT
  m_intRefTimeout = MS11DfltScanCfgData.refTimeout
  m_intRefVerifyRemindTime = DFLT_INT_REF_VERIFY_REMIND_TIME
  m_intRefVerifyTimeout = DFLT_INT_REF_VERIFY_TIME
#End If
  
  ' Initialize InfoStar default setting values
  unity_main.m_fileDevID = 0
  unity_main.m_fileVersion = INFOSTAR_VER
  m_dbgLevel = LOG_DBG_LEVEL1
  m_enableGlobalName = False
  m_enableRunMode = True
  m_globalBaseCounter = 0
  m_globalDate = "01/01/2010"
  m_globalDateCounter = 1
  m_globalNameBase = "GlobalSample"
  m_globalNameMode = "Base"
  m_allowBias = False

  ' Load instrument configuration file
  fileName = (CFG_DIR & INST_CFG_FILE)
  
  If (uniFile.OpenFileRead(fileName) = False) Then GoTo FILE_ERROR
  
  On Error GoTo FILE_ERROR
  fEncoding = uniFile.ReadBOM
    
  ' read first line to determine file format
  lineCnt = lineCnt + 1
  
  If (fEncoding = fe_ANSI) Then
    rc = uniFile.ReadAnsiLine(inString)
  Else
    rc = uniFile.ReadUnicodeLine(inString)
  End If
      
  If (rc = False) Then GoTo FILE_ERROR
          
  ' Check if older format
  If (inString <> "[signature settings]") Then
    ' Reset file position to beginning
    uniFile.SetFilePos 0, eupp_FILE_BEGIN
    fEncoding = uniFile.ReadBOM
    lineCnt = 0
    GoTo PROCESS_GLOBAL_VARS
  End If
  
  ' Process each line in .ini file
  While Not (uniFile.EOF())
    Select Case (inString)
      Case "[signature settings]"
        Call unity_main.load_file_signature_vals(fileName, uniFile, fEncoding, lineCnt, m_badIniVal)
        inString = unity_main.m_iniString
        
      Case "[global settings]"
PROCESS_GLOBAL_VARS:
        If (load_inst_global_vals(fileName, uniFile, fEncoding, lineCnt) = False) Then
          GoTo FILE_ERROR
        End If
        
      Case Else
        GoTo FILE_ERROR
    End Select
  Wend

  ' Close .ini file
  uniFile.CloseFile

#If ABBFT Then
  ' Check for invalid signature device ID
  If (unity_main.m_fileDevID <> DTID_ABBFT) Then
    unity_main.errorstring = (fileName & " had incompatible value. DevID was " & unity_main.m_fileDevID & "; updated to " & DTID_ABBFT)
    unity_main.write_error
    unity_main.m_fileDevID = DTID_ABBFT
    m_badIniVal = True
  End If
  
  If (m_ipAddr <> "") Then
    pos1 = 1
  
    ' Break #.#.#.# formatted IP address into 4 numbers
    For ii = 1 To 4
      If (pos1 < Len(m_ipAddr)) Then
        pos2 = InStr(pos1, m_ipAddr, ".")
  
        If (pos2 = 0) Then
          pos2 = Len(m_ipAddr) + 1
        End If
    
        strlen = pos2 - pos1
        txt_ipAddr(ii).Text = Mid(m_ipAddr, pos1, strlen)
      Else
        txt_ipAddr(ii).Text = 0
      End If

      pos1 = pos2 + 1
    Next ii
  Else
    For ii = 1 To 4
      txt_ipAddr(ii).Text = 0
    Next ii
  End If

  ' Check for invalid port number
  If (m_portNum <= 0) Then
    unity_main.errorstring = (fileName & " had incompatible value. PortNum was " & m_portNum & "; updated to 20000")
    unity_main.write_error
    m_badIniVal = True
    m_portNum = 20000
  End If
 
  ' Setup response timeout
  txt_portNum.Text = m_portNum
  
  ' Check for invalid response timeout
  If (m_rspTimeout < numInc_rspTimeout.Min) Or (m_rspTimeout > numInc_rspTimeout.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. RspTimeout was " & m_rspTimeout & "; updated to 12")
    unity_main.write_error
    m_badIniVal = True
    m_rspTimeout = 12
  End If
 
  ' Setup response timeout
  numInc_rspTimeout.Text = m_rspTimeout
  
  ' Check for invalid reference difference temperature
  If (m_refTempDiff < numInc_refTempDiff.Min) Or (m_refTempDiff > numInc_refTempDiff.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. RefTempDiff was " & m_refTempDiff & "; updated to 5")
    unity_main.write_error
    m_badIniVal = True
    m_refTempDiff = 5
  End If
 
  ' Setup reference difference temperature
  numInc_refTempDiff.Text = m_refTempDiff
#Else
  ' Check for invalid signature device ID
  If (unity_main.m_fileDevID <> MS11CfgData.devID) Then
    unity_main.errorstring = (fileName & " had incompatible value. DevID was " & unity_main.m_fileDevID & "; updated to " & MS11CfgData.devID)
    unity_main.write_error
    unity_main.m_fileDevID = MS11CfgData.devID
    m_badIniVal = True
  End If
  
  ' Check for invalid signature sample table
  If (unity_main.m_smplTable <> MS11CfgData.smplTblIX) Then
    unity_main.errorstring = (fileName & " had incompatible value. SmplTable was " & unity_main.m_smplTable & "; updated to " & MS11CfgData.smplTblIX)
    unity_main.write_error
    unity_main.m_smplTable = MS11CfgData.smplTblIX
    m_badIniVal = True
  End If
  
  ' Check for invalid signature scan mode
  If (unity_main.m_sysScanMode <> MS11CfgData.sysScanMode) Then
    unity_main.errorstring = (fileName & " had incompatible value. ScanMode was " & unity_main.m_sysScanMode & "; updated to " & MS11CfgData.sysScanMode)
    unity_main.write_error
    unity_main.m_sysScanMode = MS11CfgData.sysScanMode
    m_badIniVal = True
  End If
  
  ' Check for invalid number of scans
  If (m_autoSmplrPort < numInc_commPort.Min) Or (m_autoSmplrPort > numInc_commPort.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. AutoSmplrPort was " & m_autoSmplrPort & "; updated to " & DFLT_AUTO_SMPLR_PORT)
    unity_main.write_error
    m_badIniVal = True
    m_autoSmplrPort = DFLT_AUTO_SMPLR_PORT
  End If
  
  ' Setup auto-sampler comm port
  numInc_commPort.Text = m_autoSmplrPort
  
  ' Confirm batch report filepath contains '\' instead of '/'
  filePathName = m_batchRptPath
  check_filepathname_delimiters filePathName
  m_batchRptPath = filePathName
    
  ' Append "\" to batch report file path if not present
  If (Right(m_batchRptPath, 1) <> "\") Then
    m_batchRptPath = m_batchRptPath & "\"
  End If
  
  ' Setup batch report file path
  txt_batchRptPath.Text = m_batchRptPath
  
  ' Check for invalid dark scan
  If (Abs(CInt(m_darkScan)) <> sysDarkScan) Then
    unity_main.errorstring = (fileName & " had incompatible value. DarkScan was " & m_darkScan & "; updated to " & sysDarkScan)
    unity_main.write_error
    m_darkScan = sysDarkScan
    m_badIniVal = True
  End If

  ' Setup dark scan selection
  If (m_darkScan = False) Then
    chk_darkScan.Value = 0
  Else
    chk_darkScan.Value = 1
  End If
  
  ' Check for invalid dark subtract
  If (Abs(CInt(m_darkSub)) <> sysDarkSub) Then
    unity_main.errorstring = (fileName & " had incompatible value. DarkSub was " & m_darkSub & "; updated to " & sysDarkSub)
    unity_main.write_error
    m_darkSub = sysDarkSub
    m_badIniVal = True
  End If
  
  ' Setup dark subtract selection
  If (m_darkSub = False) Then
    chk_darkSub.Value = 0
  Else
    chk_darkSub.Value = 1
  End If

  ' Check for invalid number of scans
  If (m_intRefNScans < numInc_refNScans.Min) Or (m_intRefNScans > numInc_refNScans.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. RefNScans was " & m_intRefNScans & "; updated to " & MS11DfltScanCfgData.nScans4Ref)
    unity_main.write_error
    m_badIniVal = True
    m_intRefNScans = MS11DfltScanCfgData.nScans4Ref
  End If
  
  ' Setup number of scans
  numInc_refNScans.Text = m_intRefNScans
  
  ' Check for invalid reference qualification limits
  If (m_intRefPPT < numInc_refPPT.Min) Or (m_intRefPPT > numInc_refPPT.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. RefPPT was " & m_intRefPPT & "; updated to " & MS11DfltScanCfgData.ref4PPT)
    unity_main.write_error
    m_badIniVal = True
    m_intRefPPT = MS11DfltScanCfgData.ref4PPT
  End If
  
  ' Setup reference qualification limits
  numInc_refPPT.Text = m_intRefPPT
  
  ' Check for invalid reference timeout
  If (m_intRefTimeout < numInc_refTimeout.Min) Or (m_intRefTimeout > numInc_refTimeout.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. RefTimeout was " & m_intRefTimeout & "; updated to " & MS11DfltScanCfgData.refTimeout)
    unity_main.write_error
    m_badIniVal = True
    m_intRefTimeout = MS11DfltScanCfgData.refTimeout
  End If
 
  ' Setup reference timeout
  numInc_refTimeout.Text = m_intRefTimeout

  ' Check for invalid reference verification remind time
  If (m_intRefVerifyRemindTime < numInc_refVerifyRemindTime.Min) Or (m_intRefVerifyRemindTime > numInc_refVerifyRemindTime.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. RefVerifyRemindTime was " & m_intRefVerifyRemindTime & "; updated to " & DFLT_INT_REF_VERIFY_REMIND_TIME)
    unity_main.write_error
    m_badIniVal = True
    m_intRefVerifyRemindTime = DFLT_INT_REF_VERIFY_REMIND_TIME
  End If
 
  ' Setup reference verification timeout
  numInc_refVerifyRemindTime.Text = m_intRefVerifyRemindTime

  ' Check for invalid reference verification timeout
  If (m_intRefVerifyTimeout < numInc_refVerifyTimeout.Min) Or (m_intRefVerifyTimeout > numInc_refVerifyTimeout.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. RefVerifyTimeout was " & m_intRefVerifyTimeout & "; updated to " & DFLT_INT_REF_VERIFY_TIME)
    unity_main.write_error
    m_badIniVal = True
    m_intRefVerifyTimeout = DFLT_INT_REF_VERIFY_TIME
  End If
 
  ' Setup reference verification timeout
  numInc_refVerifyTimeout.Text = m_intRefVerifyTimeout

  ' Check for invalid scan direction
  If (m_scanDir <> sysScanDir) Then
    unity_main.errorstring = (fileName & " had incompatible value. ScanDir was " & m_scanDir & "; updated to " & sysScanDir)
    unity_main.write_error
    m_scanDir = sysScanDir
    m_badIniVal = True
  End If

  ' Setup scan direction selection
  Select Case (m_scanDir)
    Case SD_NONE             ' no scan direction
      opt_scanNone.Value = True
    Case SD_DOWN             ' down scan direction
      opt_scanDown.Value = True
    Case SD_UP               ' up scan direction
      opt_scanUp.Value = True
    Case SD_BOTH             ' down and up scan direction
      opt_scanBoth.Value = True
  End Select
  
  ' Check if configure to use auto-sampler and system does not support auto-sampler
  If (m_useAutoSmplr = True) And (MS11CfgData.devID = DTID_DRAWER0) Or (MS11CfgData.devID = DTID_DRAWER1) Then
    unity_main.errorstring = (fileName & " had incompatible value. UseAutoSmplr was " & m_useAutoSmplr & "; updated to False since instrument type does not support auto-sampler")
    unity_main.write_error
    m_badIniVal = True
    m_useAutoSmplr = False
  End If
  
  ' Setup utilize auto-sampler
  If (m_useAutoSmplr = True) Then
    chk_useAutoSmplr.Value = 1
  Else
    chk_useAutoSmplr.Value = 0
  End If
#End If

  ' Check for invalid signature file version
  If (unity_main.m_fileVersion <> INFOSTAR_VER) Then
    unity_main.errorstring = (fileName & " had incompatible value. Version was " & unity_main.m_fileVersion & "; updated to " & INFOSTAR_VER)
    unity_main.write_error
    unity_main.m_fileVersion = INFOSTAR_VER
    m_badIniVal = True
  End If

  ' Check for invalid debug level
  If (m_dbgLevel < numInc_dbgLevel.Min) Or (m_dbgLevel > numInc_dbgLevel.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. DbgLevel was " & m_dbgLevel & "; updated to " & LOG_DBG_LEVEL1)
    unity_main.write_error
    m_badIniVal = True
    m_dbgLevel = LOG_DBG_LEVEL1
  End If

  ' Setup debug level
  numInc_dbgLevel.Text = m_dbgLevel
  
  ' Setup enable run mode selection
  If (m_enableRunMode = False) Then
    chk_enableRunMode.Value = 0
  Else
    chk_enableRunMode.Value = 1
  End If
  
  ' Setup enable global sample name selection
  If (m_enableGlobalName = False) Then
    chk_enableGlobalSampName.Value = 0
    frame_globalSampName.enabled = False
  Else
    chk_enableGlobalSampName.Value = 1
    frame_globalSampName.enabled = True
  End If
  
  ' Check for invalid global base counter
  If (m_globalBaseCounter < frm_Inst.numInc_globalBaseCtr.Min) Or (m_globalBaseCounter > numInc_globalBaseCtr.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. GlobalBaseCtr was " & m_globalBaseCounter & "; updated to 0")
    unity_main.write_error
    m_globalBaseCounter = 0
    m_badIniVal = True
  End If
  
  ' Setup global base counter
  numInc_globalBaseCtr.Text = m_globalBaseCounter

  ' Check for invalid global date counter
  If (m_globalDateCounter < numInc_globalDateCtr.Min) Or (m_globalDateCounter > numInc_globalDateCtr.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. GlobalDateCtr was " & m_globalDateCounter & "; updated to 1")
    unity_main.write_error
    m_globalDateCounter = 1
    m_badIniVal = True
  End If
  
  ' Setup global date counter
  numInc_globalDateCtr.Text = m_globalDateCounter
  
  ' Setup global sample file base name
  txt_globalSampNameBase.Text = m_globalNameBase

  ' Setup global sample file naming convention selection
PROCESS_GLOBALNAMEMODE:
  Select Case (m_globalNameMode)
    Case "Base"
      txt_GlobalScanName.Text = txt_globalSampNameBase.Text & numInc_globalBaseCtr.Text
      opt_globalNameBase.Value = True
    Case "Ctr"
      txt_GlobalScanName.Text = numInc_globalBaseCtr.Text
      opt_globalNameCtr.Value = True
    Case "Date"
      Dim dateStrg As String
      Call frm_collect.rebuild_date(Date, dateStrg)
      txt_GlobalScanName.Text = dateStrg & "_" & numInc_globalDateCtr.Text
      opt_globalNameDate.Value = True
    Case Else       ' invalid value
      unity_main.errorstring = (fileName & " had incompatible value. NameScanType was " & m_globalNameMode & "; updated to Date")
      unity_main.write_error
      m_globalNameMode = "Base"
      m_badIniVal = True
      GoTo PROCESS_GLOBALNAMEMODE
  End Select
  
  ' Setup enable bias entry selection
  If (m_allowBias = False) Then
    chk_allowbias.Value = 0
  Else
    chk_allowbias.Value = 1
  End If
  
  ' Copy loaded global system config into system operational variables
#If ABBFT Then
  unity_main.m_mb3000.m_ipAddr = m_ipAddr
  unity_main.m_mb3000.m_portNum = m_portNum
  unity_main.m_mb3000.m_refTempDiff = m_refTempDiff
  unity_main.m_mb3000.m_rspTimeout = m_rspTimeout
#Else
  unity_main.m_autoSmplrPort = m_autoSmplrPort
  unity_main.m_batchRptPath = m_batchRptPath
  unity_main.m_darkScan = m_darkScan
  unity_main.m_darkSub = m_darkSub
  unity_main.m_intRefNScans = m_intRefNScans
  unity_main.m_intRefPPT = m_intRefPPT
  unity_main.m_intRefTimeout = m_intRefTimeout
  unity_main.m_intRefVerifyRemindTime = m_intRefVerifyRemindTime
  unity_main.m_intRefVerifyTimeout = m_intRefVerifyTimeout
  unity_main.m_scanDir = m_scanDir
  unity_main.m_useAutoSmplr = m_useAutoSmplr
  
  NumAutoSmplrTowers = 0
  tempNumAutoSmplrTowers = 0
  
  ' Setup batch scan buttons if auto-sampler in use
  If (m_useAutoSmplr = True) Then
    unity_main.cmd_runBatch.Visible = True
    unity_main.cmd_sample.Width = unity_main.cmd_runBatch.Width
    frmUtils.cmd_cfgScanBatch.Visible = True
    
    ' Initialize serial port communication w/ auto-sampler
#If SSRCS Then
    Dim parity As String
    parity = AUTO_SMPLR_PARITY
    SSRCSClientError = unity_main.SSRCSClient.InitASComms(unity_main.m_autoSmplrPort, AUTO_SMPLR_BAUD, parity, AUTO_SMPLR_RSP_TIMEOUT, AUTO_SMPLR_NUM_RETRIES, AUTO_SMPLR_CRC_USAGE)
    
    If (SSRCSClientError = 0) Then
      SSRCSClientError = unity_main.SSRCSClient.GetAllASTubesState
    Else
      If (unity_main.m_ssrcsConnected = True) Then
        CWrap.ShowMessageBoxW MLSupport.GGS_Params("ASErrMsg1", "Problem with configuring Auto-Sampler comm port %1. Verify no other communication option is configured to use this port", CStr(unity_main.m_autoSmplrPort)), vbCritical
      Else
        CWrap.ShowMessageBoxW MLSupport.GSS("frm_ssrcsConnect", "errMsg3", "Not connected to a SpectraStar RCS server"), vbCritical
      End If
    End If
#Else
    If (frm_batchRun.autoSmplr.init_comms(unity_main.m_autoSmplrPort, AUTO_SMPLR_BAUD, AUTO_SMPLR_PARITY, AUTO_SMPLR_RSP_TIMEOUT, AUTO_SMPLR_NUM_RETRIES) = True) Then
      ' Get max number of tubes supported
      frm_batchRun.autoSmplr.get_all_tubes_state AUTO_SMPLR_CRC_USAGE
    Else
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("ASErrMsg1", "Problem with configuring Auto-Sampler comm port %1. Verify no other communication option is configured to use this port", CStr(unity_main.m_autoSmplrPort)), vbCritical
    End If
#End If
  Else
    unity_main.cmd_runBatch.Visible = False
    unity_main.cmd_sample.Width = unity_main.cmd_select.Width
    frmUtils.cmd_cfgScanBatch.Visible = False
        
   ' Close serial communication w/ auto-sampler if top window system
    If (frame_autoSmplr.Visible = True) Then
#If SSRCS Then
      SSRCSClientError = unity_main.SSRCSClient.CloseASComms
#Else
      frm_batchRun.autoSmplr.close_comms
#End If
    End If
  End If
  
  ' Check if to start internal reference verification timer
  If ((unity_main.m_allowIntRefCalAccess = True) And (unity_main.m_intRefVerifyTimeout > 0) And (unity_main.m_intRefVerifyTimer = 0)) Then
    unity_main.m_intRefVerifyTimer = Timer
  End If
#End If

  unity_main.m_allowBias = m_allowBias
  unity_main.m_darkScan = m_darkScan
  unity_main.m_darkSub = m_darkSub
  unity_main.m_dbgLevel = m_dbgLevel
  unity_main.m_enableGlobalName = m_enableGlobalName
  unity_main.m_enableRunMode = m_enableRunMode
  unity_main.m_globalDate = m_globalDate
  unity_main.m_globalDateCounter = m_globalDateCounter
  unity_main.m_globalNameBase = m_globalNameBase
  unity_main.m_globalBaseCounter = m_globalBaseCounter
  unity_main.m_globalNameMode = m_globalNameMode

  ' Check if ini file had bad value
  If (m_badIniVal = True) Then
    unity_main.errorstring = (fileName & " had incompatible value(s). Updated file with proper/default value(s).")
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg7", "%1 had incompatible value(s). Updated file with proper/default value(s)", fileName)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg3", "%1. Refer to system log for more details", uniMsg), vbOKOnly
    Call savemyinsts(True, True)
  End If
  
  Exit Sub
  
FILE_ERROR:
  uniFile.CloseFile

  If (lineCnt = 0) Then
    errMsg = (fileName & " file open error." & Error$)
    uniMsg = MLSupport.GGS_Params("fileErrMsg2", "%1 file open error. %2", fileName, Error$)
  Else
    errMsg = (fileName & " file has error on line " & CStr(lineCnt) & ". " & Error$)
    uniMsg = MLSupport.GGS_Params("fileErrMsg3", "%1 file has error on line %2. %3", fileName, CStr(lineCnt), Error$)
  End If
  
  unity_main.errorstring = errMsg
  unity_main.write_error
  uniMsg = MLSupport.GGS_Params("frm_Inst.errMsg1", "%1. Using default SpectraStar values", uniMsg)
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub

Private Function load_inst_global_vals(ByVal fileName As String, ByVal uniFile As clsUniFile, ByVal fEncoding As eFileEncoding, ByRef lineCnt As Integer) As Boolean
  Dim inString As String
  Dim xx As String
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
      rc = uniFile.ReadAnsiLine(inString)
    Else
      rc = uniFile.ReadUnicodeLine(inString)
    End If
      
    If (rc = False) Then GoTo FILE_ERROR
          
    ' Get variable name and its value
    xx = InStr(1, inString, "=")
    
    ' If no "=" char, then older file format used "," filed delimiter
    If ((xx) = 0) Then
      xx = InStr(1, inString, ",")
    End If
    
    strlen = Len(inString)
    tmpStrg = Trim(Mid(inString, 1, xx - 1))
    cfgVar = LCase(tmpStrg)
    varVal = Trim(Mid(inString, xx + 1))
    
    ' Process value by variable name
    On Error GoTo BAD_INI_VALUE
    Select Case (cfgVar)
#If SSTAR Then
      Case "autosmplrport"
        m_autoSmplrPort = CInt(varVal)
      Case "batchrptpath"
        m_batchRptPath = varVal
#End If
      Case "dbglevel"
        m_dbgLevel = CInt(varVal)
      Case "darkscan"
        m_darkScan = CBool(varVal)
      Case "darksub"
        m_darkSub = CBool(varVal)
      Case "enableglobalname"
        m_enableGlobalName = CBool(varVal)
      Case "ipaddr"
        m_ipAddr = varVal
      Case "globalbasectr"
        m_globalBaseCounter = CLng(varVal)
      Case "globaldate"
        m_globalDate = varVal
      Case "globaldatectr"
        m_globalDateCounter = CLng(varVal)
      Case "globalnamebase"
        m_globalNameBase = varVal
      Case "globalnamemode"
        m_globalNameMode = varVal
      Case "mingui"
        m_enableRunMode = CBool(varVal)
      Case "portnum"
        m_portNum = CLng(varVal)
      Case "rsptimeout"
        m_rspTimeout = CLng(varVal)
      Case "refnscans"
        m_intRefNScans = CInt(varVal)
      Case "refppt"
        m_intRefPPT = CInt(varVal)
      Case "reftempdiff"
        m_refTempDiff = CInt(varVal)
      Case "reftimeout"
        m_intRefTimeout = CInt(varVal)
      Case "refverifyremindtime"
        m_intRefVerifyRemindTime = CInt(varVal)
      Case "refverifytimeout"
        m_intRefVerifyTimeout = CInt(varVal)
      Case "runtimebias"
        m_allowBias = CBool(varVal)
     
#If SSTAR Then
      Case "scandir"
        m_scanDir = CInt(varVal)
      Case "useautosmplr"
        m_useAutoSmplr = CBool(varVal)
#End If
    End Select
  Wend
  
  load_inst_global_vals = True
  Exit Function
  
BAD_INI_VALUE:
    unity_main.errorstring = (fileName & " had incompatible value. " & cfgVar & " = " & varVal & "; will use default value")
    unity_main.write_error
    m_badIniVal = True
    Resume Next
  
FILE_ERROR:
  load_inst_global_vals = False
End Function

Private Function chk_sys_cfg() As Boolean
  Dim ii As Integer
  
#If ABBFT Then
  For ii = 1 To 4
    If ((IsNumeric(txt_ipAddr(ii).Text) = False) Or (txt_ipAddr(ii).Text < 0) Or (txt_ipAddr(ii).Text > 255)) Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_inst", "errMsg1", "You must enter a valid IP Address"), vbCritical
      txt_ipAddr(ii).SetFocus
      chk_sys_cfg = False
      Exit Function
    End If
  Next ii
  
  If ((IsNumeric(txt_portNum.Text) = False) Or (txt_portNum.Text <= 0)) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_inst", "errMsg2", "You must enter a valid Port Number"), vbCritical
    txt_portNum.SetFocus
    chk_sys_cfg = False
    Exit Function
  End If
#End If

  chk_sys_cfg = True
End Function

Private Sub chk_enableGlobalSampName_Click()

  If (chk_enableGlobalSampName.Value = 1) Then
    frame_globalSampName.enabled = True
  Else
    frame_globalSampName.enabled = False
  End If
End Sub



#If SSTAR Then
Private Sub cmd_cfgSpectTreat_Click()

  unity_main.errorstring = "Spectrum Treatment Configuration screen Spectrum Treatment button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  frm_spectTreatCfg.load_cfg
  frm_spectTreatCfg.Show 1
End Sub
#End If

Private Sub cmd_cfgCSV_Click()

  unity_main.errorstring = "Global System Configuration screen CSV Report button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Call frm_csvCfg.load_cfg(False)
  frm_csvCfg.Show 1
End Sub

Private Sub cmd_cfgLIMS_Click()
  
  unity_main.errorstring = "Global System Configuration screen LIMS button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  frm_POG.setup_lims_lists
  Call frm_POG.load_lims(False)
  frm_POG.Show 1
End Sub

Private Sub cmd_cfgDynRpt_Click()

  unity_main.errorstring = "Global System Configuration screen Dynamic Report button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  Call frm_dynRptCfg.save_cfg(False, False)
  Call frm_dynRptCfg.load_cfg(False)
  frm_dynRptCfg.Show 1
End Sub

Private Sub cmd_cfgTkt_Click()
  
  unity_main.errorstring = "Global System Configuration screen Ticket Printer button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Call frm_ticket.loadticketfile(False)
  frm_ticket.Show 1
End Sub

#If SSTAR Then
Private Sub cmd_extRef_Click()

  unity_main.errorstring = "Global System Configuration screen External References button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  frm_collect.build_ref_name_list "external"
  frm_extRefMgmt.Show 1
End Sub
#End If

Private Sub cmd_superPW_Click()
  
  unity_main.errorstring = "Global System Configuration screen Supervisor Password button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  frm_guipw.loadguipw
  frm_guipw.Show 1
End Sub

Private Sub cmd_save_Click()
  
  unity_main.errorstring = "Global System Configuration screen Save Changes button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  If (chk_sys_cfg = True) Then
    Call savemyinsts(True, True)
    If m_saveInstsFail = False Then
     frm_Inst.Visible = False
     unity_main.errorstring = ("User saved new settings for configuration file: " & CFG_DIR & INST_CFG_FILE)
     unity_main.write_error (LOG_DBG_LEVEL1)
    End If
  End If
End Sub

Private Sub cmd_cancel_Click()
  
  unity_main.errorstring = "Global System Configuration screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Loadmyinstini
  frm_Inst.Visible = False
End Sub

Private Sub cmd_cfgInputs_Click()
  
  unity_main.errorstring = "Global System Configuration screen User Inputs button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  frm_buttoncfg.loadit False
  frm_buttoncfg.loadbuttonform
  frm_buttoncfg.loadbuttonconfig False
  frm_buttoncfg.Show 1
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
#If ABBFT Then
  chk_darkSub.Visible = False
  chk_darkScan.Visible = False
  cmd_extRef.Visible = False
  cmd_cfgSpectTreat.Visible = False
  Frame1.Visible = False
  frame_autoSmplr.Visible = False
  lbl_refNScans.Visible = False
  lbl_refPPT.Visible = False
  lbl_refTimeout.Visible = False
  lbl_refVerifyTimeout.Visible = False
  numInc_refNScans.Visible = False
  numInc_refPPT.Visible = False
  numInc_refTimeout.Visible = False
  numInc_refVerifyTimeout.Visible = False
#Else
  ' Check if top window
  If (MS11CfgData.devID = DTID_TOPWIND0) Or (MS11CfgData.devID = DTID_TOPWIND1) Then
    ' Display auto-sampler configuration parameters
    frame_autoSmplr.Visible = True
  Else
    ' Auto-sampler not supported for drawer systems
    frame_autoSmplr.Visible = False
  End If

  ' Check if calibrated TW system
  If (unity_main.m_allowIntRefCalAccess = True) Then
    ' Display internal reference verification time
    lbl_refVerifyTimeout.Visible = True
    numInc_refVerifyTimeout.Visible = True
  Else
    ' Internal reference verification not supported
    lbl_refVerifyTimeout.Visible = False
    numInc_refVerifyTimeout.Visible = False
  End If

  frame_network.Visible = False
  frame_reference.Visible = False
#End If
End Sub



#If SSTAR Then
Private Sub numInc_commPort_DblClick()
  
  unity_main.formfrom = 4
  unity_main.varfrom = 15
  frm_numpad.lbl_num.Caption = lbl_commPort.Caption
  frm_numpad.txt_num.Text = numInc_commPort.Text
  frm_numpad.Show 1
End Sub
#End If

Private Sub numInc_dbgLevel_DblClick()
  
  unity_main.formfrom = 4
  unity_main.varfrom = 4
  frm_numpad.lbl_num.Caption = lbl_dbgLevel.Caption
  frm_numpad.txt_num.Text = numInc_dbgLevel.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_globalBaseCtr_Change()
  
  Select Case (m_globalNameMode)
    Case "Base"
      Call opt_globalNameBase_Click
    Case "Ctr"
      Call opt_globalNameCtr_Click
  End Select
End Sub

Private Sub numInc_globalBaseCtr_DblClick()

  unity_main.formfrom = 4
  unity_main.varfrom = 6
  frm_numpad.lbl_num.Caption = MLSupport.GSS("frm_Inst", "lbl_num2", "Base Name Counter Value")
  frm_numpad.txt_num.Text = numInc_globalBaseCtr.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_globalDateCtr_Change()
  
  If (m_globalNameMode = "Date") Then
    Call opt_globalNameDate_Click
  End If
End Sub

Private Sub numInc_globalDateCtr_DblClick()

  unity_main.formfrom = 4
  unity_main.varfrom = 5
  frm_numpad.lbl_num.Caption = MLSupport.GSS("frm_Inst", "lbl_num1", "Date Counter Value")
  frm_numpad.txt_num.Text = numInc_globalDateCtr.Text
  frm_numpad.Show 1
End Sub

#If SSTAR Then
Private Sub numInc_refNScans_DblClick()
  
  unity_main.formfrom = 4
  unity_main.varfrom = 3
  frm_numpad.lbl_num.Caption = lbl_refNScans.Caption
  frm_numpad.txt_num.Text = numInc_refNScans.Text
  frm_numpad.Show 1
End Sub
#End If

#If SSTAR Then
Private Sub numInc_refPPT_DblClick()
  
  unity_main.formfrom = 4
  unity_main.varfrom = 2
  frm_numpad.lbl_num.Caption = lbl_refPPT.Caption
  frm_numpad.txt_num.Text = numInc_refPPT.Text
  frm_numpad.Show 1
End Sub
#End If

#If ABBFT Then
Private Sub numInc_refTempDiff_DblClick()

  unity_main.formfrom = 4
  unity_main.varfrom = 7
  frm_numpad.lbl_num.Caption = lbl_refTempDiff.Caption
  frm_numpad.txt_num.Text = numInc_refTempDiff.Text
  frm_numpad.Show 1
End Sub
#End If

#If SSTAR Then
Private Sub numInc_refTimeout_DblClick()
  
  unity_main.formfrom = 4
  unity_main.varfrom = 1
  frm_numpad.lbl_num.Caption = lbl_refTimeout.Caption
  frm_numpad.txt_num.Text = numInc_refTimeout.Text
  frm_numpad.Show 1
End Sub
#End If

#If SSTAR Then
Private Sub numInc_refVerifyRemindTime_DblClick()

  unity_main.formfrom = 4
  unity_main.varfrom = 17
  frm_numpad.lbl_num.Caption = lbl_refVerifyRemindTime.Caption
  frm_numpad.txt_num.Text = numInc_refVerifyRemindTime.Text
  frm_numpad.Show 1
End Sub
#End If

#If SSTAR Then
Private Sub numInc_refVerifyTimeout_DblClick()
  
  unity_main.formfrom = 4
  unity_main.varfrom = 16
  frm_numpad.lbl_num.Caption = lbl_refVerifyTimeout.Caption
  frm_numpad.txt_num.Text = numInc_refVerifyTimeout.Text
  frm_numpad.Show 1
End Sub
#End If

#If ABBFT Then
Private Sub numInc_rspTimeout_Change()

  m_commChanged = True
End Sub
#End If

#If ABBFT Then
Private Sub numInc_rspTimeout_DblClick()

  unity_main.formfrom = 4
  unity_main.varfrom = 9
  frm_numpad.lbl_num.Caption = lbl_rspTimeout.Caption
  frm_numpad.txt_num.Text = numInc_rspTimeout.Text
  frm_numpad.Show 1
End Sub
#End If

Private Sub opt_globalNameBase_Click()

  txt_GlobalScanName.Text = Trim(txt_globalSampNameBase.Text) & numInc_globalBaseCtr.Text
  m_globalNameMode = "Base"
End Sub

Private Sub opt_globalNameCtr_Click()

  txt_GlobalScanName.Text = numInc_globalBaseCtr.Text
  m_globalNameMode = "Ctr"
End Sub

Private Sub opt_globalNameDate_Click()
  Dim dateStrg As String
  
  Call frm_collect.rebuild_date(Date, dateStrg)
  txt_GlobalScanName.Text = dateStrg & "_" & numInc_globalDateCtr.Text
  m_globalNameMode = "Date"
End Sub

#If SSTAR Then
Private Sub txt_batchRptPath_DblClick(Button As Integer)

  unity_main.formfrom = 4
  unity_main.varfrom = 2
  frm_kybd.lbl_kybd.Caption = lbl_batchRptPath.Caption
  frm_kybd.txt_kybd.Text = txt_batchRptPath.Text
  frm_kybd.Show 1
End Sub
#End If

Private Sub txt_globalSampNameBase_Change()
  
  If (m_globalNameMode = "Base") Then
    txt_GlobalScanName.Text = Trim(txt_globalSampNameBase.Text) & numInc_globalBaseCtr.Text
  End If
End Sub

Private Sub txt_globalSampNameBase_DblClick(Button As Integer)
  
  unity_main.formfrom = 4
  unity_main.varfrom = 1
  frm_kybd.lbl_kybd.Caption = Label4.Caption
  frm_kybd.txt_kybd.Text = txt_globalSampNameBase.Text
  frm_kybd.Show 1
End Sub

#If ABBFT Then
Private Sub txt_ipAddr_Change(Index As Integer)

  m_commChanged = True
End Sub
#End If

#If ABBFT Then
Private Sub txt_ipAddr_DblClick(Index As Integer, Button As Integer)

  unity_main.formfrom = 4
  unity_main.varfrom = 10 + Index
  frm_numpad.lbl_num.Caption = lbl_ipAddr.Caption
  frm_numpad.txt_num.Text = txt_ipAddr(Index).Text
  frm_numpad.Show 1
End Sub
#End If

#If ABBFT Then
Private Sub txt_portNum_Change()

  m_commChanged = True
End Sub
#End If

#If ABBFT Then
Private Sub txt_portNum_DblClick(Button As Integer)
  
  unity_main.formfrom = 4
  unity_main.varfrom = 8
  frm_numpad.lbl_num.Caption = lbl_portNum.Caption
  frm_numpad.txt_num.Text = txt_portNum.Text
  frm_numpad.Show 1
End Sub
#End If








