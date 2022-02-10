VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_dynRptCfg 
   Caption         =   "Dynamic Report Configuration"
   ClientHeight    =   9000
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   11880
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
   ScaleHeight     =   9000
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniFrameXP frame_main 
      Height          =   7200
      Index           =   0
      Left            =   480
      Top             =   600
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   12700
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_dynRptCfg.frx":0000
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483642
      Tip             =   "frm_dynRptCfg.frx":0020
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_dynRptCfg.frx":0040
      Begin HexUniControls.ctlUniFrameXP frame2 
         Height          =   1695
         Left            =   7440
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":005C
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":0092
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":00B2
         Begin HexUniControls.ctlUniRadioXP opt_uniFormat 
            Height          =   450
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   2895
            _ExtentX        =   5106
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
            Caption         =   "frm_dynRptCfg.frx":00CE
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_dynRptCfg.frx":00FC
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":011C
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP opt_asciiFormat 
            Height          =   450
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   2895
            _ExtentX        =   5106
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
            Caption         =   "frm_dynRptCfg.frx":0138
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_dynRptCfg.frx":0162
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":0182
            ShowFocus       =   -1  'True
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame3 
         Height          =   3255
         Left            =   7440
         Top             =   2160
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   5741
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":019E
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":01D8
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":01F8
         Begin HexUniControls.ctlUniCheckXP chk_addTrailer 
            Height          =   450
            Left            =   120
            TabIndex        =   19
            Top             =   2280
            Width           =   2895
            _ExtentX        =   5106
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
            Caption         =   "frm_dynRptCfg.frx":0214
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_dynRptCfg.frx":025E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":027E
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniCheckXP chk_addHeader 
            Height          =   450
            Left            =   120
            TabIndex        =   18
            Top             =   1800
            Width           =   2895
            _ExtentX        =   5106
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
            Caption         =   "frm_dynRptCfg.frx":029A
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_dynRptCfg.frx":02E2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":0302
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniLabel lbl_fieldDelim 
            Height          =   300
            Left            =   120
            Top             =   360
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
            Caption         =   "frm_dynRptCfg.frx":031E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":0366
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":0386
         End
Begin HexUniControls.ctlUniComboBoxXP combo_fieldDelim
            Height          =   450
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
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
            Tip             =   "frm_dynRptCfg.frx":03A2
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   -1
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":03C2
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniCheckXP chk_addQuotes 
            Height          =   600
            Left            =   120
            TabIndex        =   134
            Top             =   1200
            Width           =   2895
            _ExtentX        =   5106
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
            Caption         =   "frm_dynRptCfg.frx":03DE
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   0
            Pressed         =   0   'False
            Tip             =   "frm_dynRptCfg.frx":0440
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":0460
            ShowFocus       =   -1  'True
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame1 
         Height          =   6615
         Left            =   240
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   11668
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":047C
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":04C4
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":04E4
         Begin HexUniControls.ctlUniTextBoxXP txt_rptName 
            Height          =   450
            Left            =   120
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   5760
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   794
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":0500
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
            Tip             =   "frm_dynRptCfg.frx":0520
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":0540
         End
         Begin HexUniControls.ctlUniLabel lbl_rptName 
            Height          =   300
            Left            =   120
            Top             =   5400
            Width           =   5175
            _ExtentX        =   9128
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
            Caption         =   "frm_dynRptCfg.frx":055C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":05A6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":05C6
         End
         Begin HexUniControls.ctlNumIncXP numInc_dateCounter 
            Height          =   600
            Left            =   4320
            TabIndex        =   9
            Top             =   1800
            Width           =   1815
            _ExtentX        =   3201
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
            MouseIcon       =   "frm_dynRptCfg.frx":05E2
         End
         Begin HexUniControls.ctlUniRadioXP opt_dateName 
            Height          =   600
            Left            =   120
            TabIndex        =   8
            Top             =   1800
            Width           =   4005
            _ExtentX        =   7064
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
            Caption         =   "frm_dynRptCfg.frx":05FE
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_dynRptCfg.frx":0646
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":0666
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_manualSuffix 
            Height          =   450
            Left            =   4320
            TabIndex        =   14
            Top             =   4800
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":0682
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
            Tip             =   "frm_dynRptCfg.frx":06A2
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":06C2
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_manualPrefix 
            Height          =   450
            Left            =   4320
            TabIndex        =   12
            Top             =   4320
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":06DE
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
            Tip             =   "frm_dynRptCfg.frx":06FE
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":071E
         End
         Begin HexUniControls.ctlUniCheckXP chk_addManualSuffix 
            Height          =   450
            Left            =   600
            TabIndex        =   13
            Top             =   4800
            Width           =   3495
            _ExtentX        =   6165
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
            Caption         =   "frm_dynRptCfg.frx":073A
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_dynRptCfg.frx":0780
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":07A0
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniCheckXP chk_addManualPrefix 
            Height          =   450
            Left            =   600
            TabIndex        =   11
            Top             =   4320
            Width           =   3495
            _ExtentX        =   6165
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
            Caption         =   "frm_dynRptCfg.frx":07BC
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_dynRptCfg.frx":0802
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":0822
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlNumIncXP numInc_baseCounter 
            Height          =   600
            Left            =   4320
            TabIndex        =   6
            Top             =   2400
            Width           =   1815
            _ExtentX        =   3201
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
            Min             =   0
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
            MouseIcon       =   "frm_dynRptCfg.frx":083E
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_baseName 
            Height          =   450
            Left            =   600
            TabIndex        =   7
            Top             =   3360
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":085A
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
            Tip             =   "frm_dynRptCfg.frx":0886
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":08A6
         End
         Begin HexUniControls.ctlUniLabel lbl_baseName 
            Height          =   300
            Left            =   600
            Top             =   3000
            Width           =   3495
            _ExtentX        =   6165
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
            Caption         =   "frm_dynRptCfg.frx":08C2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":08FE
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":091E
         End
         Begin HexUniControls.ctlUniRadioXP opt_manualName 
            Height          =   450
            Left            =   120
            TabIndex        =   10
            Top             =   3840
            Width           =   4005
            _ExtentX        =   7064
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
            Caption         =   "frm_dynRptCfg.frx":093A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_dynRptCfg.frx":098E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":09AE
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_fileExt 
            Height          =   450
            Left            =   4320
            TabIndex        =   3
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":09CA
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
            Tip             =   "frm_dynRptCfg.frx":09F0
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":0A10
         End
         Begin HexUniControls.ctlUniLabel lbl_fileExt 
            Height          =   300
            Left            =   4320
            Top             =   360
            Width           =   2625
            _ExtentX        =   4630
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
            Caption         =   "frm_dynRptCfg.frx":0A2C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":0A72
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":0A92
         End
         Begin HexUniControls.ctlUniRadioXP opt_baseName 
            Height          =   600
            Left            =   120
            TabIndex        =   5
            Top             =   2400
            Width           =   4005
            _ExtentX        =   7064
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
            Caption         =   "frm_dynRptCfg.frx":0AAE
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_dynRptCfg.frx":0B00
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":0B20
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP opt_sampleName 
            Height          =   450
            Left            =   120
            TabIndex        =   4
            Top             =   1320
            Width           =   4005
            _ExtentX        =   7064
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
            Caption         =   "frm_dynRptCfg.frx":0B3C
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_dynRptCfg.frx":0B80
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":0BA0
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_filePath 
            Height          =   450
            Left            =   120
            TabIndex        =   2
            Top             =   720
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":0BBC
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
            Tip             =   "frm_dynRptCfg.frx":0BFE
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":0C1E
         End
         Begin HexUniControls.ctlUniLabel lbl_filePath 
            Height          =   300
            Left            =   120
            Top             =   360
            Width           =   4005
            _ExtentX        =   7064
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
            Caption         =   "frm_dynRptCfg.frx":0C3A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":0C6C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":0C8C
         End
      End
   End
   Begin HexUniControls.ctlUniFrameXP frame_main 
      Height          =   7200
      Index           =   2
      Left            =   480
      Top             =   600
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   12700
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_dynRptCfg.frx":0CA8
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483642
      Tip             =   "frm_dynRptCfg.frx":0CC8
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_dynRptCfg.frx":0CE8
      Begin HexUniControls.ctlUniLabel lbl_usrNumFields 
         Height          =   615
         Left            =   1440
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
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
         Caption         =   "frm_dynRptCfg.frx":0D04
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_dynRptCfg.frx":0D4E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":0D6E
      End
      Begin HexUniControls.ctlNumIncXP numInc_usrNumFields 
         Height          =   615
         Left            =   240
         TabIndex        =   50
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
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
         Text            =   "9"
         Min             =   0
         Max             =   9
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
         MouseIcon       =   "frm_dynRptCfg.frx":0D8A
      End
      Begin HexUniControls.ctlUniFrameXP frame_usr 
         Height          =   1995
         Index           =   9
         Left            =   7200
         Top             =   4920
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":0DA6
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":0DDE
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":0DFE
         Begin HexUniControls.ctlUniTextBoxXP txt_usrFieldTxt 
            Height          =   450
            Index           =   9
            Left            =   120
            TabIndex        =   77
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":0E1A
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
            Tip             =   "frm_dynRptCfg.frx":0E3A
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":0E5A
         End
Begin HexUniControls.ctlUniComboBoxXP combo_usrFieldType
            Height          =   420
            Index           =   9
            Left            =   120
            TabIndex        =   76
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":0E76
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":0E96
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_usrMaxChars 
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   75
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":0EB2
         End
         Begin HexUniControls.ctlUniLabel lbl_usrFieldType 
            Height          =   450
            Index           =   9
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":0ECE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":0EF6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":0F16
         End
         Begin HexUniControls.ctlUniLabel lbl_usrMaxChrs 
            Height          =   450
            Index           =   9
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":0F32
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":0F6E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":0F8E
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_usr 
         Height          =   1995
         Index           =   8
         Left            =   3720
         Top             =   4920
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":0FAA
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":0FE2
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":1002
         Begin HexUniControls.ctlUniTextBoxXP txt_usrFieldTxt 
            Height          =   450
            Index           =   8
            Left            =   120
            TabIndex        =   74
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":101E
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
            Tip             =   "frm_dynRptCfg.frx":103E
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":105E
         End
Begin HexUniControls.ctlUniComboBoxXP combo_usrFieldType
            Height          =   420
            Index           =   8
            Left            =   120
            TabIndex        =   73
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":107A
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":109A
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_usrMaxChars 
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   72
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":10B6
         End
         Begin HexUniControls.ctlUniLabel lbl_usrFieldType 
            Height          =   450
            Index           =   8
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":10D2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":10FA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":111A
         End
         Begin HexUniControls.ctlUniLabel lbl_usrMaxChrs 
            Height          =   450
            Index           =   8
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":1136
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":1172
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1192
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_usr 
         Height          =   1995
         Index           =   7
         Left            =   240
         Top             =   4920
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":11AE
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":11E6
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":1206
         Begin HexUniControls.ctlUniTextBoxXP txt_usrFieldTxt 
            Height          =   450
            Index           =   7
            Left            =   120
            TabIndex        =   71
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":1222
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
            Tip             =   "frm_dynRptCfg.frx":1242
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1262
         End
Begin HexUniControls.ctlUniComboBoxXP combo_usrFieldType
            Height          =   420
            Index           =   7
            Left            =   120
            TabIndex        =   70
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":127E
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":129E
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_usrMaxChars 
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":12BA
         End
         Begin HexUniControls.ctlUniLabel lbl_usrFieldType 
            Height          =   450
            Index           =   7
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":12D6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":12FE
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":131E
         End
         Begin HexUniControls.ctlUniLabel lbl_usrMaxChrs 
            Height          =   450
            Index           =   7
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":133A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":1376
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1396
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_usr 
         Height          =   1995
         Index           =   6
         Left            =   7200
         Top             =   2880
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":13B2
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":13EA
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":140A
         Begin HexUniControls.ctlUniTextBoxXP txt_usrFieldTxt 
            Height          =   450
            Index           =   6
            Left            =   120
            TabIndex        =   68
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":1426
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
            Tip             =   "frm_dynRptCfg.frx":1446
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1466
         End
Begin HexUniControls.ctlUniComboBoxXP combo_usrFieldType
            Height          =   420
            Index           =   6
            Left            =   120
            TabIndex        =   67
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":1482
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":14A2
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_usrMaxChars 
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":14BE
         End
         Begin HexUniControls.ctlUniLabel lbl_usrFieldType 
            Height          =   450
            Index           =   6
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":14DA
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":1502
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1522
         End
         Begin HexUniControls.ctlUniLabel lbl_usrMaxChrs 
            Height          =   450
            Index           =   6
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":153E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":157A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":159A
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_usr 
         Height          =   1995
         Index           =   5
         Left            =   3720
         Top             =   2880
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":15B6
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":15EE
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":160E
         Begin HexUniControls.ctlUniTextBoxXP txt_usrFieldTxt 
            Height          =   450
            Index           =   5
            Left            =   120
            TabIndex        =   65
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":162A
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
            Tip             =   "frm_dynRptCfg.frx":164A
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":166A
         End
Begin HexUniControls.ctlUniComboBoxXP combo_usrFieldType
            Height          =   420
            Index           =   5
            Left            =   120
            TabIndex        =   64
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":1686
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":16A6
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_usrMaxChars 
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   63
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":16C2
         End
         Begin HexUniControls.ctlUniLabel lbl_usrFieldType 
            Height          =   450
            Index           =   5
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":16DE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":1706
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1726
         End
         Begin HexUniControls.ctlUniLabel lbl_usrMaxChrs 
            Height          =   450
            Index           =   5
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":1742
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":177E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":179E
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_usr 
         Height          =   1995
         Index           =   4
         Left            =   240
         Top             =   2880
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":17BA
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":17F2
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":1812
         Begin HexUniControls.ctlUniTextBoxXP txt_usrFieldTxt 
            Height          =   450
            Index           =   4
            Left            =   120
            TabIndex        =   62
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":182E
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
            Tip             =   "frm_dynRptCfg.frx":184E
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":186E
         End
Begin HexUniControls.ctlUniComboBoxXP combo_usrFieldType
            Height          =   420
            Index           =   4
            Left            =   120
            TabIndex        =   61
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":188A
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":18AA
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_usrMaxChars 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":18C6
         End
         Begin HexUniControls.ctlUniLabel lbl_usrFieldType 
            Height          =   450
            Index           =   4
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":18E2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":190A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":192A
         End
         Begin HexUniControls.ctlUniLabel lbl_usrMaxChrs 
            Height          =   450
            Index           =   4
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":1946
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":1982
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":19A2
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_usr 
         Height          =   1995
         Index           =   3
         Left            =   7200
         Top             =   840
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":19BE
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":19F6
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":1A16
         Begin HexUniControls.ctlUniTextBoxXP txt_usrFieldTxt 
            Height          =   450
            Index           =   3
            Left            =   120
            TabIndex        =   59
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":1A32
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
            Tip             =   "frm_dynRptCfg.frx":1A52
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1A72
         End
Begin HexUniControls.ctlUniComboBoxXP combo_usrFieldType
            Height          =   420
            Index           =   3
            Left            =   120
            TabIndex        =   58
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":1A8E
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1AAE
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_usrMaxChars 
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":1ACA
         End
         Begin HexUniControls.ctlUniLabel lbl_usrFieldType 
            Height          =   450
            Index           =   3
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":1AE6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":1B0E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1B2E
         End
         Begin HexUniControls.ctlUniLabel lbl_usrMaxChrs 
            Height          =   450
            Index           =   3
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":1B4A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":1B86
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1BA6
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_usr 
         Height          =   1995
         Index           =   2
         Left            =   3720
         Top             =   840
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":1BC2
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":1BFA
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":1C1A
         Begin HexUniControls.ctlUniTextBoxXP txt_usrFieldTxt 
            Height          =   450
            Index           =   2
            Left            =   120
            TabIndex        =   56
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":1C36
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
            Tip             =   "frm_dynRptCfg.frx":1C56
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1C76
         End
Begin HexUniControls.ctlUniComboBoxXP combo_usrFieldType
            Height          =   420
            Index           =   2
            Left            =   120
            TabIndex        =   55
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":1C92
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1CB2
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_usrMaxChars 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":1CCE
         End
         Begin HexUniControls.ctlUniLabel lbl_usrFieldType 
            Height          =   450
            Index           =   2
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":1CEA
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":1D12
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1D32
         End
         Begin HexUniControls.ctlUniLabel lbl_usrMaxChrs 
            Height          =   450
            Index           =   2
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":1D4E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":1D8A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1DAA
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_usr 
         Height          =   1995
         Index           =   1
         Left            =   240
         Top             =   840
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":1DC6
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":1DFE
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":1E1E
         Begin HexUniControls.ctlUniTextBoxXP txt_usrFieldTxt 
            Height          =   450
            Index           =   1
            Left            =   120
            TabIndex        =   53
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":1E3A
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
            Tip             =   "frm_dynRptCfg.frx":1E5A
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1E7A
         End
Begin HexUniControls.ctlUniComboBoxXP combo_usrFieldType
            Height          =   420
            Index           =   1
            Left            =   120
            TabIndex        =   52
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":1E96
            Sorted          =   0   'False
            HScroll         =   -1  'True
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1EB6
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_usrMaxChars 
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":1ED2
         End
         Begin HexUniControls.ctlUniLabel lbl_usrMaxChrs 
            Height          =   450
            Index           =   1
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":1EEE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":1F2A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1F4A
         End
         Begin HexUniControls.ctlUniLabel lbl_usrFieldType 
            Height          =   450
            Index           =   1
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":1F66
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":1F8E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":1FAE
         End
      End
      Begin HexUniControls.ctlUniRadioXP opt_usrPosPre 
         Height          =   615
         Left            =   3960
         TabIndex        =   135
         Top             =   120
         Width           =   3075
         _ExtentX        =   5424
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
         Caption         =   "frm_dynRptCfg.frx":1FCA
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_dynRptCfg.frx":2022
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":2042
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniRadioXP opt_usrPosPost 
         Height          =   615
         Left            =   7320
         TabIndex        =   136
         Top             =   120
         Width           =   3075
         _ExtentX        =   5424
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
         Caption         =   "frm_dynRptCfg.frx":205E
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_dynRptCfg.frx":20B4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":20D4
         ShowFocus       =   -1  'True
      End
   End
   Begin HexUniControls.ctlUniFrameXP frame_main 
      Height          =   7200
      Index           =   3
      Left            =   480
      Top             =   600
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   12700
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_dynRptCfg.frx":20F0
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483642
      Tip             =   "frm_dynRptCfg.frx":2110
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_dynRptCfg.frx":2130
      Begin HexUniControls.ctlUniLabel lbl_recNumFields 
         Height          =   615
         Left            =   1440
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
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
         Caption         =   "frm_dynRptCfg.frx":214C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_dynRptCfg.frx":2196
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":21B6
      End
      Begin HexUniControls.ctlNumIncXP numInc_recNumFields 
         Height          =   615
         Left            =   240
         TabIndex        =   78
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
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
         Text            =   "9"
         Min             =   0
         Max             =   9
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
         MouseIcon       =   "frm_dynRptCfg.frx":21D2
      End
      Begin HexUniControls.ctlUniFrameXP frame_rec 
         Height          =   1995
         Index           =   9
         Left            =   7200
         Top             =   4920
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":21EE
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":2226
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":2246
         Begin HexUniControls.ctlUniTextBoxXP txt_recFieldTxt 
            Height          =   450
            Index           =   9
            Left            =   120
            TabIndex        =   79
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":2262
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
            Tip             =   "frm_dynRptCfg.frx":2282
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":22A2
         End
Begin HexUniControls.ctlUniComboBoxXP combo_recFieldType
            Height          =   420
            Index           =   9
            Left            =   120
            TabIndex        =   80
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":22BE
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":22DE
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_recMaxChars 
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   81
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":22FA
         End
         Begin HexUniControls.ctlUniLabel lbl_recFieldType 
            Height          =   450
            Index           =   9
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":2316
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":233E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":235E
         End
         Begin HexUniControls.ctlUniLabel lbl_recMaxChrs 
            Height          =   450
            Index           =   9
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":237A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":23B6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":23D6
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_rec 
         Height          =   1995
         Index           =   8
         Left            =   3720
         Top             =   4920
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":23F2
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":242A
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":244A
         Begin HexUniControls.ctlUniTextBoxXP txt_recFieldTxt 
            Height          =   450
            Index           =   8
            Left            =   120
            TabIndex        =   82
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":2466
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
            Tip             =   "frm_dynRptCfg.frx":2486
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":24A6
         End
Begin HexUniControls.ctlUniComboBoxXP combo_recFieldType
            Height          =   420
            Index           =   8
            Left            =   120
            TabIndex        =   83
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":24C2
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":24E2
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_recMaxChars 
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   84
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":24FE
         End
         Begin HexUniControls.ctlUniLabel lbl_recFieldType 
            Height          =   450
            Index           =   8
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":251A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":2542
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":2562
         End
         Begin HexUniControls.ctlUniLabel lbl_recMaxChrs 
            Height          =   450
            Index           =   8
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":257E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":25BA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":25DA
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_rec 
         Height          =   1995
         Index           =   7
         Left            =   240
         Top             =   4920
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":25F6
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":262E
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":264E
         Begin HexUniControls.ctlUniTextBoxXP txt_recFieldTxt 
            Height          =   450
            Index           =   7
            Left            =   120
            TabIndex        =   85
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":266A
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
            Tip             =   "frm_dynRptCfg.frx":268A
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":26AA
         End
Begin HexUniControls.ctlUniComboBoxXP combo_recFieldType
            Height          =   420
            Index           =   7
            Left            =   120
            TabIndex        =   86
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":26C6
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":26E6
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_recMaxChars 
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   87
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":2702
         End
         Begin HexUniControls.ctlUniLabel lbl_recFieldType 
            Height          =   450
            Index           =   7
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":271E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":2746
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":2766
         End
         Begin HexUniControls.ctlUniLabel lbl_recMaxChrs 
            Height          =   450
            Index           =   7
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":2782
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":27BE
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":27DE
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_rec 
         Height          =   1995
         Index           =   6
         Left            =   7200
         Top             =   2880
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":27FA
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":2832
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":2852
         Begin HexUniControls.ctlUniTextBoxXP txt_recFieldTxt 
            Height          =   450
            Index           =   6
            Left            =   120
            TabIndex        =   88
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":286E
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
            Tip             =   "frm_dynRptCfg.frx":288E
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":28AE
         End
Begin HexUniControls.ctlUniComboBoxXP combo_recFieldType
            Height          =   420
            Index           =   6
            Left            =   120
            TabIndex        =   89
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":28CA
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":28EA
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_recMaxChars 
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   90
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":2906
         End
         Begin HexUniControls.ctlUniLabel lbl_recFieldType 
            Height          =   450
            Index           =   6
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":2922
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":294A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":296A
         End
         Begin HexUniControls.ctlUniLabel lbl_recMaxChrs 
            Height          =   450
            Index           =   6
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":2986
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":29C2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":29E2
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_rec 
         Height          =   1995
         Index           =   5
         Left            =   3720
         Top             =   2880
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":29FE
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":2A36
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":2A56
         Begin HexUniControls.ctlUniTextBoxXP txt_recFieldTxt 
            Height          =   450
            Index           =   5
            Left            =   120
            TabIndex        =   91
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":2A72
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
            Tip             =   "frm_dynRptCfg.frx":2A92
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":2AB2
         End
Begin HexUniControls.ctlUniComboBoxXP combo_recFieldType
            Height          =   420
            Index           =   5
            Left            =   120
            TabIndex        =   92
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":2ACE
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":2AEE
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_recMaxChars 
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   93
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":2B0A
         End
         Begin HexUniControls.ctlUniLabel lbl_recFieldType 
            Height          =   450
            Index           =   5
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":2B26
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":2B4E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":2B6E
         End
         Begin HexUniControls.ctlUniLabel lbl_recMaxChrs 
            Height          =   450
            Index           =   5
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":2B8A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":2BC6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":2BE6
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_rec 
         Height          =   1995
         Index           =   4
         Left            =   240
         Top             =   2880
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":2C02
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":2C3A
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":2C5A
         Begin HexUniControls.ctlUniTextBoxXP txt_recFieldTxt 
            Height          =   450
            Index           =   4
            Left            =   120
            TabIndex        =   94
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":2C76
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
            Tip             =   "frm_dynRptCfg.frx":2C96
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":2CB6
         End
Begin HexUniControls.ctlUniComboBoxXP combo_recFieldType
            Height          =   420
            Index           =   4
            Left            =   120
            TabIndex        =   95
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":2CD2
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":2CF2
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_recMaxChars 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   96
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":2D0E
         End
         Begin HexUniControls.ctlUniLabel lbl_recFieldType 
            Height          =   450
            Index           =   4
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":2D2A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":2D52
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":2D72
         End
         Begin HexUniControls.ctlUniLabel lbl_recMaxChrs 
            Height          =   450
            Index           =   4
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":2D8E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":2DCA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":2DEA
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_rec 
         Height          =   1995
         Index           =   3
         Left            =   7200
         Top             =   840
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":2E06
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":2E3E
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":2E5E
         Begin HexUniControls.ctlUniTextBoxXP txt_recFieldTxt 
            Height          =   450
            Index           =   3
            Left            =   120
            TabIndex        =   97
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":2E7A
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
            Tip             =   "frm_dynRptCfg.frx":2E9A
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":2EBA
         End
Begin HexUniControls.ctlUniComboBoxXP combo_recFieldType
            Height          =   420
            Index           =   3
            Left            =   120
            TabIndex        =   98
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":2ED6
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":2EF6
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_recMaxChars 
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   99
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":2F12
         End
         Begin HexUniControls.ctlUniLabel lbl_recFieldType 
            Height          =   450
            Index           =   3
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":2F2E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":2F56
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":2F76
         End
         Begin HexUniControls.ctlUniLabel lbl_recMaxChrs 
            Height          =   450
            Index           =   3
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":2F92
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":2FCE
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":2FEE
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_rec 
         Height          =   1995
         Index           =   2
         Left            =   3720
         Top             =   840
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":300A
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":3042
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":3062
         Begin HexUniControls.ctlUniTextBoxXP txt_recFieldTxt 
            Height          =   450
            Index           =   2
            Left            =   120
            TabIndex        =   100
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":307E
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
            Tip             =   "frm_dynRptCfg.frx":309E
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":30BE
         End
Begin HexUniControls.ctlUniComboBoxXP combo_recFieldType
            Height          =   420
            Index           =   2
            Left            =   120
            TabIndex        =   101
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":30DA
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":30FA
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_recMaxChars 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   102
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":3116
         End
         Begin HexUniControls.ctlUniLabel lbl_recFieldType 
            Height          =   450
            Index           =   2
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":3132
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":315A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":317A
         End
         Begin HexUniControls.ctlUniLabel lbl_recMaxChrs 
            Height          =   450
            Index           =   2
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":3196
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":31D2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":31F2
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_rec 
         Height          =   1995
         Index           =   1
         Left            =   240
         Top             =   840
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":320E
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":3246
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":3266
         Begin HexUniControls.ctlUniTextBoxXP txt_recFieldTxt 
            Height          =   450
            Index           =   1
            Left            =   120
            TabIndex        =   103
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":3282
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
            Tip             =   "frm_dynRptCfg.frx":32A2
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":32C2
         End
Begin HexUniControls.ctlUniComboBoxXP combo_recFieldType
            Height          =   420
            Index           =   1
            Left            =   120
            TabIndex        =   104
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":32DE
            Sorted          =   0   'False
            HScroll         =   -1  'True
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":32FE
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_recMaxChars 
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   105
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":331A
         End
         Begin HexUniControls.ctlUniLabel lbl_recMaxChrs 
            Height          =   450
            Index           =   1
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":3336
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":3372
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":3392
         End
         Begin HexUniControls.ctlUniLabel lbl_recFieldType 
            Height          =   450
            Index           =   1
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":33AE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":33D6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":33F6
         End
      End
   End
   Begin HexUniControls.ctlUniFrameXP frame_main 
      Height          =   7200
      Index           =   1
      Left            =   480
      Top             =   600
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   12700
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_dynRptCfg.frx":3412
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483642
      Tip             =   "frm_dynRptCfg.frx":3432
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_dynRptCfg.frx":3452
      Begin HexUniControls.ctlNumIncXP numInc_hdrNumFields 
         Height          =   615
         Left            =   240
         TabIndex        =   22
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
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
         Text            =   "9"
         Min             =   0
         Max             =   9
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
         MouseIcon       =   "frm_dynRptCfg.frx":346E
      End
      Begin HexUniControls.ctlUniFrameXP frame_hdr 
         Height          =   1995
         Index           =   1
         Left            =   240
         Top             =   840
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":348A
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":34C2
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":34E2
         Begin HexUniControls.ctlUniLabel lbl_hdrMaxChrs 
            Height          =   450
            Index           =   1
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":34FE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":353A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":355A
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_hdrFieldTxt 
            Height          =   450
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":3576
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
            Tip             =   "frm_dynRptCfg.frx":3596
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":35B6
         End
         Begin HexUniControls.ctlNumIncXP numInc_hdrMaxChars 
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":35D2
         End
         Begin HexUniControls.ctlUniLabel lbl_hdrFieldType 
            Height          =   450
            Index           =   1
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":35EE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":3616
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":3636
         End
Begin HexUniControls.ctlUniComboBoxXP combo_hdrFieldType
            Height          =   420
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":3652
            Sorted          =   0   'False
            HScroll         =   -1  'True
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":3672
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_hdr 
         Height          =   1995
         Index           =   9
         Left            =   7200
         Top             =   4920
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":368E
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":36C6
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":36E6
         Begin HexUniControls.ctlUniTextBoxXP txt_hdrFieldTxt 
            Height          =   450
            Index           =   9
            Left            =   120
            TabIndex        =   49
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":3702
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
            Tip             =   "frm_dynRptCfg.frx":3722
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":3742
         End
         Begin HexUniControls.ctlUniLabel lbl_hdrFieldType 
            Height          =   450
            Index           =   9
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":375E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":3786
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":37A6
         End
Begin HexUniControls.ctlUniComboBoxXP combo_hdrFieldType
            Height          =   420
            Index           =   9
            Left            =   120
            TabIndex        =   48
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":37C2
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":37E2
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_hdrMaxChrs 
            Height          =   450
            Index           =   9
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":37FE
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":383A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":385A
         End
         Begin HexUniControls.ctlNumIncXP numInc_hdrMaxChars 
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":3876
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_hdr 
         Height          =   1995
         Index           =   8
         Left            =   3720
         Top             =   4920
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":3892
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":38CA
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":38EA
         Begin HexUniControls.ctlUniTextBoxXP txt_hdrFieldTxt 
            Height          =   450
            Index           =   8
            Left            =   120
            TabIndex        =   46
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":3906
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
            Tip             =   "frm_dynRptCfg.frx":3926
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":3946
         End
         Begin HexUniControls.ctlUniLabel lbl_hdrFieldType 
            Height          =   450
            Index           =   8
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":3962
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":398A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":39AA
         End
Begin HexUniControls.ctlUniComboBoxXP combo_hdrFieldType
            Height          =   420
            Index           =   8
            Left            =   120
            TabIndex        =   45
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":39C6
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":39E6
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_hdrMaxChrs 
            Height          =   450
            Index           =   8
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":3A02
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":3A3E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":3A5E
         End
         Begin HexUniControls.ctlNumIncXP numInc_hdrMaxChars 
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":3A7A
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_hdr 
         Height          =   1995
         Index           =   6
         Left            =   7200
         Top             =   2880
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":3A96
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":3ACE
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":3AEE
         Begin HexUniControls.ctlUniTextBoxXP txt_hdrFieldTxt 
            Height          =   450
            Index           =   6
            Left            =   120
            TabIndex        =   40
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":3B0A
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
            Tip             =   "frm_dynRptCfg.frx":3B2A
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":3B4A
         End
         Begin HexUniControls.ctlUniLabel lbl_hdrFieldType 
            Height          =   450
            Index           =   6
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":3B66
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":3B8E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":3BAE
         End
Begin HexUniControls.ctlUniComboBoxXP combo_hdrFieldType
            Height          =   420
            Index           =   6
            Left            =   120
            TabIndex        =   39
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":3BCA
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":3BEA
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_hdrMaxChrs 
            Height          =   450
            Index           =   6
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":3C06
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":3C42
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":3C62
         End
         Begin HexUniControls.ctlNumIncXP numInc_hdrMaxChars 
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":3C7E
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_hdr 
         Height          =   1995
         Index           =   5
         Left            =   3720
         Top             =   2880
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":3C9A
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":3CD2
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":3CF2
         Begin HexUniControls.ctlUniTextBoxXP txt_hdrFieldTxt 
            Height          =   450
            Index           =   5
            Left            =   120
            TabIndex        =   37
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":3D0E
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
            Tip             =   "frm_dynRptCfg.frx":3D2E
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":3D4E
         End
         Begin HexUniControls.ctlUniLabel lbl_hdrFieldType 
            Height          =   450
            Index           =   5
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":3D6A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":3D92
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":3DB2
         End
Begin HexUniControls.ctlUniComboBoxXP combo_hdrFieldType
            Height          =   420
            Index           =   5
            Left            =   120
            TabIndex        =   36
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":3DCE
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":3DEE
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_hdrMaxChrs 
            Height          =   450
            Index           =   5
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":3E0A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":3E46
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":3E66
         End
         Begin HexUniControls.ctlNumIncXP numInc_hdrMaxChars 
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":3E82
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_hdr 
         Height          =   1995
         Index           =   7
         Left            =   240
         Top             =   4920
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":3E9E
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":3ED6
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":3EF6
         Begin HexUniControls.ctlUniTextBoxXP txt_hdrFieldTxt 
            Height          =   450
            Index           =   7
            Left            =   120
            TabIndex        =   43
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":3F12
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
            Tip             =   "frm_dynRptCfg.frx":3F32
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":3F52
         End
         Begin HexUniControls.ctlUniLabel lbl_hdrFieldType 
            Height          =   450
            Index           =   7
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":3F6E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":3F96
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":3FB6
         End
Begin HexUniControls.ctlUniComboBoxXP combo_hdrFieldType
            Height          =   420
            Index           =   7
            Left            =   120
            TabIndex        =   42
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":3FD2
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":3FF2
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_hdrMaxChrs 
            Height          =   450
            Index           =   7
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":400E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":404A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":406A
         End
         Begin HexUniControls.ctlNumIncXP numInc_hdrMaxChars 
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":4086
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_hdr 
         Height          =   1995
         Index           =   4
         Left            =   240
         Top             =   2880
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":40A2
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":40DA
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":40FA
         Begin HexUniControls.ctlUniTextBoxXP txt_hdrFieldTxt 
            Height          =   450
            Index           =   4
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":4116
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
            Tip             =   "frm_dynRptCfg.frx":4136
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":4156
         End
         Begin HexUniControls.ctlUniLabel lbl_hdrFieldType 
            Height          =   450
            Index           =   4
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":4172
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":419A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":41BA
         End
Begin HexUniControls.ctlUniComboBoxXP combo_hdrFieldType
            Height          =   420
            Index           =   4
            Left            =   120
            TabIndex        =   33
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":41D6
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":41F6
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_hdrMaxChrs 
            Height          =   450
            Index           =   4
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":4212
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":424E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":426E
         End
         Begin HexUniControls.ctlNumIncXP numInc_hdrMaxChars 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":428A
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_hdr 
         Height          =   1995
         Index           =   3
         Left            =   7200
         Top             =   840
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":42A6
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":42DE
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":42FE
         Begin HexUniControls.ctlUniTextBoxXP txt_hdrFieldTxt 
            Height          =   450
            Index           =   3
            Left            =   120
            TabIndex        =   31
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":431A
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
            Tip             =   "frm_dynRptCfg.frx":433A
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":435A
         End
         Begin HexUniControls.ctlUniLabel lbl_hdrFieldType 
            Height          =   450
            Index           =   3
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":4376
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":439E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":43BE
         End
Begin HexUniControls.ctlUniComboBoxXP combo_hdrFieldType
            Height          =   420
            Index           =   3
            Left            =   120
            TabIndex        =   30
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":43DA
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":43FA
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_hdrMaxChrs 
            Height          =   450
            Index           =   3
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":4416
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":4452
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":4472
         End
         Begin HexUniControls.ctlNumIncXP numInc_hdrMaxChars 
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":448E
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_hdr 
         Height          =   1995
         Index           =   2
         Left            =   3720
         Top             =   840
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":44AA
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":44E2
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":4502
         Begin HexUniControls.ctlUniTextBoxXP txt_hdrFieldTxt 
            Height          =   450
            Index           =   2
            Left            =   120
            TabIndex        =   28
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":451E
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
            Tip             =   "frm_dynRptCfg.frx":453E
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":455E
         End
         Begin HexUniControls.ctlUniLabel lbl_hdrFieldType 
            Height          =   450
            Index           =   2
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":457A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":45A2
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":45C2
         End
Begin HexUniControls.ctlUniComboBoxXP combo_hdrFieldType
            Height          =   420
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":45DE
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":45FE
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_hdrMaxChrs 
            Height          =   450
            Index           =   2
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":461A
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":4656
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":4676
         End
         Begin HexUniControls.ctlNumIncXP numInc_hdrMaxChars 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":4692
         End
      End
      Begin HexUniControls.ctlUniLabel lbl_hdrNumFields 
         Height          =   615
         Left            =   1440
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
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
         Caption         =   "frm_dynRptCfg.frx":46AE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_dynRptCfg.frx":46F8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":4718
      End
   End
   Begin HexUniControls.ctlUniFrameXP frame_main 
      Height          =   7200
      Index           =   4
      Left            =   480
      Top             =   600
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   12700
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_dynRptCfg.frx":4734
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483642
      Tip             =   "frm_dynRptCfg.frx":4754
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_dynRptCfg.frx":4774
      Begin HexUniControls.ctlNumIncXP numInc_trlNumFields 
         Height          =   615
         Left            =   240
         TabIndex        =   106
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
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
         Text            =   "9"
         Min             =   0
         Max             =   9
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
         MouseIcon       =   "frm_dynRptCfg.frx":4790
      End
      Begin HexUniControls.ctlUniFrameXP frame_trl 
         Height          =   1995
         Index           =   1
         Left            =   240
         Top             =   840
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":47AC
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":47E4
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":4804
         Begin HexUniControls.ctlUniLabel lbl_trlMaxChrs 
            Height          =   450
            Index           =   1
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":4820
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":485C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":487C
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_trlFieldTxt 
            Height          =   450
            Index           =   1
            Left            =   120
            TabIndex        =   107
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":4898
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
            Tip             =   "frm_dynRptCfg.frx":48B8
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":48D8
         End
         Begin HexUniControls.ctlNumIncXP numInc_trlMaxChars 
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   108
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":48F4
         End
         Begin HexUniControls.ctlUniLabel lbl_trlFieldType 
            Height          =   450
            Index           =   1
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":4910
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":4938
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":4958
         End
Begin HexUniControls.ctlUniComboBoxXP combo_trlFieldType
            Height          =   420
            Index           =   1
            Left            =   120
            TabIndex        =   109
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":4974
            Sorted          =   0   'False
            HScroll         =   -1  'True
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":4994
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_trl 
         Height          =   1995
         Index           =   9
         Left            =   7200
         Top             =   4920
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":49B0
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":49E8
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":4A08
         Begin HexUniControls.ctlUniTextBoxXP txt_trlFieldTxt 
            Height          =   450
            Index           =   9
            Left            =   120
            TabIndex        =   110
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":4A24
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
            Tip             =   "frm_dynRptCfg.frx":4A44
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":4A64
         End
         Begin HexUniControls.ctlUniLabel lbl_trlFieldType 
            Height          =   450
            Index           =   9
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":4A80
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":4AA8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":4AC8
         End
Begin HexUniControls.ctlUniComboBoxXP combo_trlFieldType
            Height          =   420
            Index           =   9
            Left            =   120
            TabIndex        =   111
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":4AE4
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":4B04
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_trlMaxChrs 
            Height          =   450
            Index           =   9
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":4B20
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":4B5C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":4B7C
         End
         Begin HexUniControls.ctlNumIncXP numInc_trlMaxChars 
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   112
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":4B98
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_trl 
         Height          =   1995
         Index           =   8
         Left            =   3720
         Top             =   4920
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":4BB4
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":4BEC
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":4C0C
         Begin HexUniControls.ctlUniTextBoxXP txt_trlFieldTxt 
            Height          =   450
            Index           =   8
            Left            =   120
            TabIndex        =   113
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":4C28
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
            Tip             =   "frm_dynRptCfg.frx":4C48
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":4C68
         End
         Begin HexUniControls.ctlUniLabel lbl_trlFieldType 
            Height          =   450
            Index           =   8
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":4C84
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":4CAC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":4CCC
         End
Begin HexUniControls.ctlUniComboBoxXP combo_trlFieldType
            Height          =   420
            Index           =   8
            Left            =   120
            TabIndex        =   114
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":4CE8
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":4D08
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_trlMaxChrs 
            Height          =   450
            Index           =   8
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":4D24
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":4D60
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":4D80
         End
         Begin HexUniControls.ctlNumIncXP numInc_trlMaxChars 
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   115
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":4D9C
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_trl 
         Height          =   1995
         Index           =   6
         Left            =   7200
         Top             =   2880
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":4DB8
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":4DF0
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":4E10
         Begin HexUniControls.ctlUniTextBoxXP txt_trlFieldTxt 
            Height          =   450
            Index           =   6
            Left            =   120
            TabIndex        =   116
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":4E2C
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
            Tip             =   "frm_dynRptCfg.frx":4E4C
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":4E6C
         End
         Begin HexUniControls.ctlUniLabel lbl_trlFieldType 
            Height          =   450
            Index           =   6
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":4E88
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":4EB0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":4ED0
         End
Begin HexUniControls.ctlUniComboBoxXP combo_trlFieldType
            Height          =   420
            Index           =   6
            Left            =   120
            TabIndex        =   117
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":4EEC
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":4F0C
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_trlMaxChrs 
            Height          =   450
            Index           =   6
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":4F28
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":4F64
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":4F84
         End
         Begin HexUniControls.ctlNumIncXP numInc_trlMaxChars 
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   118
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":4FA0
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_trl 
         Height          =   1995
         Index           =   5
         Left            =   3720
         Top             =   2880
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":4FBC
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":4FF4
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":5014
         Begin HexUniControls.ctlUniTextBoxXP txt_trlFieldTxt 
            Height          =   450
            Index           =   5
            Left            =   120
            TabIndex        =   119
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":5030
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
            Tip             =   "frm_dynRptCfg.frx":5050
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":5070
         End
         Begin HexUniControls.ctlUniLabel lbl_trlFieldType 
            Height          =   450
            Index           =   5
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":508C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":50B4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":50D4
         End
Begin HexUniControls.ctlUniComboBoxXP combo_trlFieldType
            Height          =   420
            Index           =   5
            Left            =   120
            TabIndex        =   120
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":50F0
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":5110
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_trlMaxChrs 
            Height          =   450
            Index           =   5
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":512C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":5168
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":5188
         End
         Begin HexUniControls.ctlNumIncXP numInc_trlMaxChars 
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   121
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":51A4
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_trl 
         Height          =   1995
         Index           =   7
         Left            =   240
         Top             =   4920
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":51C0
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":51F8
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":5218
         Begin HexUniControls.ctlUniTextBoxXP txt_trlFieldTxt 
            Height          =   450
            Index           =   7
            Left            =   120
            TabIndex        =   122
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":5234
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
            Tip             =   "frm_dynRptCfg.frx":5254
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":5274
         End
         Begin HexUniControls.ctlUniLabel lbl_trlFieldType 
            Height          =   450
            Index           =   7
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":5290
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":52B8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":52D8
         End
Begin HexUniControls.ctlUniComboBoxXP combo_trlFieldType
            Height          =   420
            Index           =   7
            Left            =   120
            TabIndex        =   123
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":52F4
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":5314
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_trlMaxChrs 
            Height          =   450
            Index           =   7
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":5330
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":536C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":538C
         End
         Begin HexUniControls.ctlNumIncXP numInc_trlMaxChars 
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   124
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":53A8
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_trl 
         Height          =   1995
         Index           =   4
         Left            =   240
         Top             =   2880
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":53C4
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":53FC
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":541C
         Begin HexUniControls.ctlUniTextBoxXP txt_trlFieldTxt 
            Height          =   450
            Index           =   4
            Left            =   120
            TabIndex        =   125
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":5438
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
            Tip             =   "frm_dynRptCfg.frx":5458
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":5478
         End
         Begin HexUniControls.ctlUniLabel lbl_trlFieldType 
            Height          =   450
            Index           =   4
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":5494
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":54BC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":54DC
         End
Begin HexUniControls.ctlUniComboBoxXP combo_trlFieldType
            Height          =   420
            Index           =   4
            Left            =   120
            TabIndex        =   126
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":54F8
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":5518
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_trlMaxChrs 
            Height          =   450
            Index           =   4
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":5534
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":5570
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":5590
         End
         Begin HexUniControls.ctlNumIncXP numInc_trlMaxChars 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   127
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":55AC
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_trl 
         Height          =   1995
         Index           =   3
         Left            =   7200
         Top             =   840
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":55C8
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":5600
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":5620
         Begin HexUniControls.ctlUniTextBoxXP txt_trlFieldTxt 
            Height          =   450
            Index           =   3
            Left            =   120
            TabIndex        =   128
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":563C
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
            Tip             =   "frm_dynRptCfg.frx":565C
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":567C
         End
         Begin HexUniControls.ctlUniLabel lbl_trlFieldType 
            Height          =   450
            Index           =   3
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":5698
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":56C0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":56E0
         End
Begin HexUniControls.ctlUniComboBoxXP combo_trlFieldType
            Height          =   420
            Index           =   3
            Left            =   120
            TabIndex        =   129
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":56FC
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":571C
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_trlMaxChrs 
            Height          =   450
            Index           =   3
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":5738
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":5774
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":5794
         End
         Begin HexUniControls.ctlNumIncXP numInc_trlMaxChars 
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   130
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":57B0
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_trl 
         Height          =   1995
         Index           =   2
         Left            =   3720
         Top             =   840
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_dynRptCfg.frx":57CC
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_dynRptCfg.frx":5804
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":5824
         Begin HexUniControls.ctlUniTextBoxXP txt_trlFieldTxt 
            Height          =   450
            Index           =   2
            Left            =   120
            TabIndex        =   131
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_dynRptCfg.frx":5840
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
            Tip             =   "frm_dynRptCfg.frx":5860
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":5880
         End
         Begin HexUniControls.ctlUniLabel lbl_trlFieldType 
            Height          =   450
            Index           =   2
            Left            =   2040
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
            Caption         =   "frm_dynRptCfg.frx":589C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":58C4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":58E4
         End
Begin HexUniControls.ctlUniComboBoxXP combo_trlFieldType
            Height          =   420
            Index           =   2
            Left            =   120
            TabIndex        =   132
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
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
            Tip             =   "frm_dynRptCfg.frx":5900
            Sorted          =   0   'False
            HScroll         =   0   'False
Style=3
            ButtonBackColor =   -2147483633
            ButtonForeColor =   0
            ButtonWidth     =   17
            Locked          =   0   'False
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            TrapTab         =   0   'False
            ButtonStyle     =   4
            SelectorStyle   =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":5920
            DropDownOnTextClick=   -1  'True
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_trlMaxChrs 
            Height          =   450
            Index           =   2
            Left            =   1320
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "frm_dynRptCfg.frx":593C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_dynRptCfg.frx":5978
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_dynRptCfg.frx":5998
         End
         Begin HexUniControls.ctlNumIncXP numInc_trlMaxChars 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   133
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            Text            =   "99"
            Min             =   1
            Max             =   99
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
            MouseIcon       =   "frm_dynRptCfg.frx":59B4
         End
      End
      Begin HexUniControls.ctlUniLabel lbl_trlNumFields 
         Height          =   615
         Left            =   1440
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
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
         Caption         =   "frm_dynRptCfg.frx":59D0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_dynRptCfg.frx":5A1A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_dynRptCfg.frx":5A3A
      End
   End
   Begin HexUniControls.ctlUniTabbedXP tab_frame 
      Height          =   320
      Left            =   600
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   300
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   16711680
      BackColorTab    =   -2147483633
      ForeColorTab    =   8947848
      ShowFocus       =   0   'False
      Tip             =   "frm_dynRptCfg.frx":5A56
      ButtonStyle     =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_dynRptCfg.frx":5A76
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   11280
      Top             =   120
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9000
      FormDesignWidth =   11880
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   645
      Left            =   8520
      TabIndex        =   0
      Top             =   8040
      Width           =   1995
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_dynRptCfg.frx":5A92
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_dynRptCfg.frx":5ABE
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_dynRptCfg.frx":5ADE
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   645
      Left            =   5880
      TabIndex        =   1
      Top             =   8040
      Width           =   1995
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_dynRptCfg.frx":5AFA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_dynRptCfg.frx":5B32
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_dynRptCfg.frx":5B52
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   11400
      Top             =   720
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
      Left            =   10680
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_dynRptCfg.frx":5B6E
   End
End
Attribute VB_Name = "frm_dynRptCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_fileVersion As String

Private m_rptAddHeader As Boolean
Private m_rptAddManPrefix As Boolean
Private m_rptAddManSuffix As Boolean
Private m_rptAddQuotes As Boolean
Private m_rptAddTrailer As Boolean
Private m_rptBaseCounter As Long
Private m_rptBaseName As String
Private m_rptDate As String
Private m_rptDateCounter As Long
Private m_rptFieldDelim As String
Private m_rptFileExt As String
Private m_rptFileFormat As String
Private m_rptFilePath As String
Private m_rptHdrFieldTxt(1 To MAX_RPT_FIELDS) As String
Private m_rptHdrFieldType(1 To MAX_RPT_FIELDS) As String
Private m_rptHdrMaxChars(1 To MAX_RPT_FIELDS) As Integer
Private m_rptHdrNumFields As Integer
Private m_rptManPrefix As String
Private m_rptManSuffix As String
Private m_rptNameMode As String
Private m_rptRecFieldTxt(1 To MAX_RPT_FIELDS) As String
Private m_rptRecFieldType(1 To MAX_RPT_FIELDS) As String
Private m_rptRecMaxChars(1 To MAX_RPT_FIELDS) As Integer
Private m_rptRecNumFields As Integer
Private m_rptTrlFieldTxt(1 To MAX_RPT_FIELDS) As String
Private m_rptTrlFieldType(1 To MAX_RPT_FIELDS) As String
Private m_rptTrlMaxChars(1 To MAX_RPT_FIELDS) As Integer
Private m_rptTrlNumFields As Integer
Private m_rptUsrFieldTxt(1 To MAX_RPT_FIELDS) As String
Private m_rptUsrFieldType(1 To MAX_RPT_FIELDS) As String
Private m_rptUsrMaxChars(1 To MAX_RPT_FIELDS) As Integer
Private m_rptUsrNumFields As Integer
Private m_rptUsrPos As String

Private m_saveRptIni As Boolean
Private m_dateTime As Variant
Private m_date As Variant
Private m_time As Variant
Private m_badIniVal As Boolean

Const DFLT_FILE_EXT = "txt"
Const DFLT_FILE_FORMAT = "ASCII"
Const DFLT_FILE_NAME_MODE = "Sample"
Const DFLT_FLD_DELIM = ","
Const DFLT_MAX_CHARS = 20
Const DFLT_FLD_TYPE = "Text"
Const DFLT_USR_POS = "PreRecords"

Public Sub load_cfg(mustBeCfg As Boolean)
  Dim ii As Integer
  Dim fileName As String
  Dim filePathName As String
  Dim exist As Boolean
  Dim errMsg As String
  Dim uniMsg As String

  m_badIniVal = False
  fileName = (CFG_DIR & DYN_REPORT_CFG_FILE)
  exist = CFile.st_FileExist(fileName)
  
  If (exist = False) Then
    ' Check if file should be configured
    If (mustBeCfg = True) Then
      uniMsg = MLSupport.GGS_Params("fileErrMsg6", "%1 file not found", fileName)
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_dynRptCfg.errMsg1", "%1. Loaded product configured for dynamic report output; please configure the Dynamic Report settings and save them to file.", uniMsg), vbExclamation
      Exit Sub
    End If
  End If
  
  ' Setup default values
  m_fileVersion = INFOSTAR_VER
  m_rptAddHeader = False
  m_rptAddManPrefix = False
  m_rptAddManSuffix = False
  m_rptAddQuotes = False
  m_rptAddTrailer = False
  m_rptBaseCounter = 0
  m_rptBaseName = "Report"
  m_rptDate = "1/1/2010"
  m_rptDateCounter = 1
  m_rptFieldDelim = DFLT_FLD_DELIM
  m_rptFileExt = DFLT_FILE_EXT
  m_rptFileFormat = DFLT_FILE_FORMAT
  m_rptFilePath = REPORTS_DIR
  
  For ii = 1 To MAX_RPT_FIELDS
    m_rptHdrFieldTxt(ii) = ""
    m_rptHdrFieldType(ii) = DFLT_FLD_TYPE
    m_rptHdrMaxChars(ii) = DFLT_MAX_CHARS
    m_rptRecFieldTxt(ii) = ""
    m_rptRecFieldType(ii) = DFLT_FLD_TYPE
    m_rptRecMaxChars(ii) = DFLT_MAX_CHARS
    m_rptTrlFieldTxt(ii) = ""
    m_rptTrlFieldType(ii) = DFLT_FLD_TYPE
    m_rptTrlMaxChars(ii) = DFLT_MAX_CHARS
    m_rptUsrFieldTxt(ii) = ""
    m_rptUsrFieldType(ii) = DFLT_FLD_TYPE
    m_rptUsrMaxChars(ii) = DFLT_MAX_CHARS
  Next ii
  
  m_rptHdrNumFields = 0
  m_rptManPrefix = ""
  m_rptManSuffix = ""
  m_rptNameMode = DFLT_FILE_NAME_MODE
  m_rptRecNumFields = 0
  m_rptTrlNumFields = 0
  m_rptUsrNumFields = 0
  m_rptUsrPos = DFLT_USR_POS
  
  If (exist = True) Then
    If (load_dyn_rpt_file_vals(fileName) = True) Then
      ' Check for invalid file version
      If (m_fileVersion <> INFOSTAR_VER) Then
        unity_main.errorstring = (fileName & " had incompatible value. Version was " & m_fileVersion & "; updated to " & INFOSTAR_VER)
        unity_main.write_error
        m_fileVersion = INFOSTAR_VER
        m_badIniVal = True
      End If
      
      ' Check for invalid base name counter
      If (m_rptBaseCounter < numInc_baseCounter.Min) Or (m_rptBaseCounter > numInc_baseCounter.Max) Then
        unity_main.errorstring = (fileName & " had incompatible value. BaseCounter was " & m_rptBaseCounter & "; updated to 0")
        unity_main.write_error
        m_rptBaseCounter = 0
        m_badIniVal = True
      End If
  
      ' Check for invalid date name counter
      If (m_rptDateCounter < numInc_dateCounter.Min) Or (m_rptDateCounter > numInc_dateCounter.Max) Then
        unity_main.errorstring = (fileName & " had incompatible value. DateCounter was " & m_rptDateCounter & "; updated to 1")
        unity_main.write_error
        m_rptDateCounter = 1
        m_badIniVal = True
      End If
  
      ' Check for invalid data field delimiter
      Select Case (m_rptFieldDelim)
        Case "SP"
        Case ","
        Case ";"
        Case ":"
        Case "\"
        Case "\\"
        Case "/"
        Case "//"
        Case Else
          unity_main.errorstring = (fileName & " had incompatible value. FieldDelim was " & m_rptFieldDelim & "; updated to " & DFLT_FLD_DELIM)
          unity_main.write_error
          m_rptFieldDelim = DFLT_FLD_DELIM
          m_badIniVal = True
      End Select
      
      ' Check for invalid report file extension
      If (m_rptFileExt = "") Then
        unity_main.errorstring = (fileName & " had incompatible value. FileExt was " & m_rptFileExt & "; updated to " & DFLT_FILE_EXT)
        unity_main.write_error
        m_rptFileExt = DFLT_FILE_EXT
        m_badIniVal = True
      End If
      
      ' Check for invalid report file format
      Select Case (m_rptFileFormat)
        Case "ASCII"
        Case "Unicode"
        Case Else
          unity_main.errorstring = (fileName & " had incompatible value. FileFormat was " & m_rptFileFormat & "; updated to " & DFLT_FILE_FORMAT)
          unity_main.write_error
          m_rptFileFormat = DFLT_FILE_FORMAT
          m_badIniVal = True
      End Select
      
      ' Check for invalid header max chars
      For ii = 1 To MAX_RPT_FIELDS
        If (m_rptHdrMaxChars(ii) < numInc_hdrMaxChars(ii).Min) Or (m_rptHdrMaxChars(ii) > numInc_hdrMaxChars(ii).Max) Then
          unity_main.errorstring = (fileName & " had incompatible value. HdrMaxChars" & ii & " was " & m_rptHdrMaxChars(ii) & "; updated to " & DFLT_MAX_CHARS)
          unity_main.write_error
          m_rptHdrMaxChars(ii) = DFLT_MAX_CHARS
          m_badIniVal = True
        End If
      Next ii
      
      ' Check for invalid header number of fields
      If (m_rptHdrNumFields < numInc_hdrNumFields.Min) Or (m_rptHdrNumFields > numInc_hdrNumFields.Max) Then
        unity_main.errorstring = (fileName & " had incompatible value. HdrNumFields was " & m_rptHdrNumFields & "; updated to " & numInc_hdrNumFields.Min)
        unity_main.write_error
        m_rptHdrNumFields = numInc_hdrNumFields.Min
        m_badIniVal = True
      End If
  
      ' Check for invalid user inputs max chars
      For ii = 1 To MAX_RPT_FIELDS
        If (m_rptUsrMaxChars(ii) < numInc_usrMaxChars(ii).Min) Or (m_rptUsrMaxChars(ii) > numInc_usrMaxChars(ii).Max) Then
          unity_main.errorstring = (fileName & " had incompatible value. UsrMaxChars" & ii & " was " & m_rptUsrMaxChars(ii) & "; updated to " & DFLT_MAX_CHARS)
          unity_main.write_error
          m_rptUsrMaxChars(ii) = DFLT_MAX_CHARS
          m_badIniVal = True
        End If
      Next ii
    
      ' Check for invalid user inputs number of fields
      If (m_rptUsrNumFields < numInc_usrNumFields.Min) Or (m_rptUsrNumFields > numInc_usrNumFields.Max) Then
        unity_main.errorstring = (fileName & " had incompatible value. UsrNumFields was " & m_rptUsrNumFields & "; updated to " & numInc_usrNumFields.Min)
        unity_main.write_error
        m_rptUsrNumFields = numInc_usrNumFields.Min
        m_badIniVal = True
      End If
  
      ' Check for invalid record max chars
      For ii = 1 To MAX_RPT_FIELDS
        If (m_rptRecMaxChars(ii) < numInc_recMaxChars(ii).Min) Or (m_rptRecMaxChars(ii) > numInc_recMaxChars(ii).Max) Then
          unity_main.errorstring = (fileName & " had incompatible value. RecMaxChars" & ii & " was " & m_rptRecMaxChars(ii) & "; updated to " & DFLT_MAX_CHARS)
          unity_main.write_error
          m_rptRecMaxChars(ii) = DFLT_MAX_CHARS
          m_badIniVal = True
        End If
      Next ii
      
      ' Check for invalid record number of fields
      If (m_rptRecNumFields < numInc_recNumFields.Min) Or (m_rptRecNumFields > numInc_recNumFields.Max) Then
        unity_main.errorstring = (fileName & " had incompatible value. RecNumFields was " & m_rptRecNumFields & "; updated to " & numInc_recNumFields.Min)
        unity_main.write_error
        m_rptRecNumFields = numInc_recNumFields.Min
        m_badIniVal = True
      End If
  
      ' Check for invalid trailer max chars
      For ii = 1 To MAX_RPT_FIELDS
        If (m_rptTrlMaxChars(ii) < numInc_trlMaxChars(ii).Min) Or (m_rptTrlMaxChars(ii) > numInc_trlMaxChars(ii).Max) Then
          unity_main.errorstring = (fileName & " had incompatible value. TrlMaxChars" & ii & " was " & m_rptTrlMaxChars(ii) & "; updated to " & DFLT_MAX_CHARS)
          unity_main.write_error
          m_rptTrlMaxChars(ii) = DFLT_MAX_CHARS
          m_badIniVal = True
        End If
      Next ii
    
      ' Check for invalid trailer number of fields
      If (m_rptTrlNumFields < numInc_trlNumFields.Min) Or (m_rptTrlNumFields > numInc_trlNumFields.Max) Then
        unity_main.errorstring = (fileName & " had incompatible value. TrlNumFields was " & m_rptTrlNumFields & "; updated to " & numInc_trlNumFields.Min)
        unity_main.write_error
        m_rptTrlNumFields = numInc_trlNumFields.Min
        m_badIniVal = True
      End If
      
      ' Check for invalid user inputs output position
      Select Case (m_rptUsrPos)
        Case "PreRecords"
        Case "PostRecords"
        Case Else
          unity_main.errorstring = (fileName & " had incompatible value. UsrPos was " & m_rptUsrPos & "; updated to " & DFLT_USR_POS)
          unity_main.write_error
          m_rptUsrPos = DFLT_USR_POS
          m_badIniVal = True
      End Select
    End If
  End If

  ' Setup write report header selection
  If (m_rptAddHeader = True) Then
    frm_dynRptCfg.chk_addHeader.Value = 1
    frm_dynRptCfg.frame_main(1).Enabled = True
  Else
    frm_dynRptCfg.chk_addHeader.Value = 0
    frm_dynRptCfg.frame_main(1).Enabled = False
  End If

  ' Setup add report name prefix selection
  If (m_rptAddManPrefix = True) Then
    frm_dynRptCfg.chk_addManualPrefix.Value = 1
  Else
    frm_dynRptCfg.chk_addManualPrefix.Value = 0
  End If
  
  ' Setup add report name suffix selection
  If (m_rptAddManSuffix = True) Then
    frm_dynRptCfg.chk_addManualSuffix.Value = 1
  Else
    frm_dynRptCfg.chk_addManualSuffix.Value = 0
  End If
  
  ' Setup add report data field quotes selection
  If (m_rptAddQuotes = True) Then
    frm_dynRptCfg.chk_addQuotes.Value = 1
  Else
    frm_dynRptCfg.chk_addQuotes.Value = 0
  End If
  
  ' Setup write report trailer selection
  If (m_rptAddTrailer = True) Then
    frm_dynRptCfg.chk_addTrailer.Value = 1
    frm_dynRptCfg.frame_main(3).Enabled = True
  Else
    frm_dynRptCfg.chk_addTrailer.Value = 0
    frm_dynRptCfg.frame_main(3).Enabled = False
  End If
  
  ' Setup base name counter
  frm_dynRptCfg.numInc_baseCounter.Text = m_rptBaseCounter
  
  ' Setup base name
  frm_dynRptCfg.txt_baseName.Text = m_rptBaseName
  
  ' Setup report date & date name counter
  frm_dynRptCfg.numInc_dateCounter.Text = m_rptDateCounter
  
  ' Setup data field delimiter
  frm_dynRptCfg.combo_fieldDelim.Text = m_rptFieldDelim
  
  ' Remove any leading "." from report file extension
  If (Left(m_rptFileExt, 1) = ".") Then
    m_rptFileExt = Mid(m_rptFileExt, 2)
  End If
  
  ' Setup report file extension
  frm_dynRptCfg.txt_fileExt.Text = m_rptFileExt
  
  ' Confirm report filepath contains '\' instead of '/'
  filePathName = m_rptFilePath
  check_filepathname_delimiters filePathName
  m_rptFilePath = filePathName
    
  ' Append "\" to report file path if not present
  If (Right(m_rptFilePath, 1) <> "\") Then
    m_rptFilePath = m_rptFilePath & "\"
  End If
  
  ' Setup report file path
  frm_dynRptCfg.txt_filePath.Text = m_rptFilePath

  ' Setup report file format
  If (m_rptFileFormat = "ASCII") Then
    frm_dynRptCfg.opt_asciiFormat.Value = True
  Else
    frm_dynRptCfg.opt_uniFormat.Value = True
  End If
  
  ' Setup header, user inputs, record and trailer fields text
  For ii = 1 To MAX_RPT_FIELDS
    frm_dynRptCfg.txt_hdrFieldTxt(ii).Text = m_rptHdrFieldTxt(ii)
    frm_dynRptCfg.txt_usrFieldTxt(ii).Text = m_rptUsrFieldTxt(ii)
    frm_dynRptCfg.txt_recFieldTxt(ii).Text = m_rptRecFieldTxt(ii)
    frm_dynRptCfg.txt_trlFieldTxt(ii).Text = m_rptTrlFieldTxt(ii)
  Next ii

  ' Check and setup header fields type
  For ii = 1 To MAX_RPT_FIELDS
    Select Case (m_rptHdrFieldType(ii))
      Case "Text"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = True
      Case "DateTime"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "Date"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "Time"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "YYMMDD"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "YYYYMMDD"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "HHMMSS"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "ProdName"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "SampID"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "Comment"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "ManEntry"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "SerNum"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "Input1"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "Input2"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "Input3"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "Input4"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "Input5"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "Input6"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "Input7"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "Input8"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "NumRecs"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case "NumLines"
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = False
      Case Else
        unity_main.errorstring = (fileName & " had incompatible value. HdrFieldType" & ii & " was " & m_rptHdrFieldType(ii) & "; updated to " & DFLT_FLD_TYPE)
        unity_main.write_error
        m_rptHdrFieldType(ii) = DFLT_FLD_TYPE
        m_badIniVal = True
        frm_dynRptCfg.txt_hdrFieldTxt(ii).Visible = True
    End Select
    
    frm_dynRptCfg.combo_hdrFieldType(ii).Text = m_rptHdrFieldType(ii)
  Next ii

  ' Check and setup user inputs fields type
  For ii = 1 To MAX_RPT_FIELDS
    Select Case (m_rptUsrFieldType(ii))
      Case "Text"
        frm_dynRptCfg.txt_usrFieldTxt(ii).Visible = True
      Case "Input1"
        frm_dynRptCfg.txt_usrFieldTxt(ii).Visible = False
      Case "Input2"
        frm_dynRptCfg.txt_usrFieldTxt(ii).Visible = False
      Case "Input3"
        frm_dynRptCfg.txt_usrFieldTxt(ii).Visible = False
      Case "Input4"
        frm_dynRptCfg.txt_usrFieldTxt(ii).Visible = False
      Case "Input5"
        frm_dynRptCfg.txt_usrFieldTxt(ii).Visible = False
      Case "Input6"
        frm_dynRptCfg.txt_usrFieldTxt(ii).Visible = False
      Case "Input7"
        frm_dynRptCfg.txt_usrFieldTxt(ii).Visible = False
      Case "Input8"
        frm_dynRptCfg.txt_usrFieldTxt(ii).Visible = False
      Case Else
        unity_main.errorstring = (fileName & " had incompatible value. UsrFieldType" & ii & " was " & m_rptHdrFieldType(ii) & "; updated to " & DFLT_FLD_TYPE)
        unity_main.write_error
        m_rptUsrFieldType(ii) = DFLT_FLD_TYPE
        m_badIniVal = True
        frm_dynRptCfg.txt_usrFieldTxt(ii).Visible = True
    End Select
    
    frm_dynRptCfg.combo_usrFieldType(ii).Text = m_rptUsrFieldType(ii)
  Next ii

  ' Check and record header fields type
  For ii = 1 To MAX_RPT_FIELDS
    Select Case (m_rptRecFieldType(ii))
      Case "Text"
        frm_dynRptCfg.txt_recFieldTxt(ii).Visible = True
      Case "PropName"
        frm_dynRptCfg.txt_recFieldTxt(ii).Visible = False
      Case "PropVal"
        frm_dynRptCfg.txt_recFieldTxt(ii).Visible = False
      Case "PropMDst"
        frm_dynRptCfg.txt_recFieldTxt(ii).Visible = False
      Case "PropSRes"
        frm_dynRptCfg.txt_recFieldTxt(ii).Visible = False
      Case "PropOutl"
        frm_dynRptCfg.txt_recFieldTxt(ii).Visible = False
      Case "PropND"
        frm_dynRptCfg.txt_recFieldTxt(ii).Visible = False
      Case "PropIntc"
        frm_dynRptCfg.txt_recFieldTxt(ii).Visible = False
      Case "PropSlop"
        frm_dynRptCfg.txt_recFieldTxt(ii).Visible = False
      Case Else
        unity_main.errorstring = (fileName & " had incompatible value. RecFieldType" & ii & " was " & m_rptRecFieldType(ii) & "; updated to " & DFLT_FLD_TYPE)
        unity_main.write_error
        m_rptRecFieldType(ii) = DFLT_FLD_TYPE
        m_badIniVal = True
        frm_dynRptCfg.txt_recFieldTxt(ii).Visible = True
    End Select
  
    frm_dynRptCfg.combo_recFieldType(ii).Text = m_rptRecFieldType(ii)
  Next ii

  ' Check and setup trailer fields type
  For ii = 1 To MAX_RPT_FIELDS
    Select Case (m_rptTrlFieldType(ii))
      Case "Text"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = True
      Case "DateTime"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "Date"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "Time"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "YYMMDD"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "YYYYMMDD"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "HHMMSS"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "ProdName"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "SampID"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "Comment"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "ManEntry"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "SerNum"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "Input1"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "Input2"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "Input3"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "Input4"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "Input5"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "Input6"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "Input7"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "Input8"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "NumRecs"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case "NumLines"
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = False
      Case Else
        unity_main.errorstring = (fileName & " had incompatible value. TrlFieldType" & ii & " was " & m_rptTrlFieldType(ii) & "; updated to " & DFLT_FLD_TYPE)
        unity_main.write_error
        m_rptTrlFieldType(ii) = DFLT_FLD_TYPE
        m_badIniVal = True
        frm_dynRptCfg.txt_trlFieldTxt(ii).Visible = True
    End Select
    
    frm_dynRptCfg.combo_trlFieldType(ii).Text = m_rptTrlFieldType(ii)
  Next ii

  ' Setup header, user inputs, record and trailer max chars
  For ii = 1 To MAX_RPT_FIELDS
    frm_dynRptCfg.numInc_hdrMaxChars(ii).Text = m_rptHdrMaxChars(ii)
    frm_dynRptCfg.numInc_usrMaxChars(ii).Text = m_rptUsrMaxChars(ii)
    frm_dynRptCfg.numInc_recMaxChars(ii).Text = m_rptRecMaxChars(ii)
    frm_dynRptCfg.numInc_trlMaxChars(ii).Text = m_rptTrlMaxChars(ii)
  Next ii
    
  ' Setup header number of fields
  frm_dynRptCfg.numInc_hdrNumFields.Text = m_rptHdrNumFields
  
  ' Setup user inputs number of fields
  frm_dynRptCfg.numInc_usrNumFields.Text = m_rptUsrNumFields
  
  If (m_rptUsrNumFields = 0) Then
    frm_dynRptCfg.opt_usrPosPre.Visible = False
    frm_dynRptCfg.opt_usrPosPost.Visible = False
  Else
    frm_dynRptCfg.opt_usrPosPre.Visible = True
    frm_dynRptCfg.opt_usrPosPost.Visible = True
  End If
  
  ' Setup record number of fields
  frm_dynRptCfg.numInc_recNumFields.Text = m_rptRecNumFields
  
  ' Setup trailer number of fields
  frm_dynRptCfg.numInc_trlNumFields.Text = m_rptTrlNumFields
  
  ' Setup report name prefix and suffix
  frm_dynRptCfg.txt_manualPrefix.Text = m_rptManPrefix
  frm_dynRptCfg.txt_manualSuffix.Text = m_rptManSuffix

  ' Check and setup report name mode
  If (m_rptNameMode = "Sample") Then
    frm_dynRptCfg.opt_sampleName.Value = True
  Else
    If (m_rptNameMode = "Base") Then
      frm_dynRptCfg.opt_baseName.Value = True
    Else
      If (m_rptNameMode = "Date") Then
        frm_dynRptCfg.opt_dateName.Value = True
      Else
        If (m_rptNameMode = "Manual") Then
          frm_dynRptCfg.opt_manualName.Value = True
        Else
          unity_main.errorstring = (fileName & " had incompatible value. NameMode was " & m_rptNameMode & "; updated to " & DFLT_FILE_NAME_MODE)
          unity_main.write_error
          m_rptNameMode = DFLT_FILE_NAME_MODE
          frm_dynRptCfg.opt_sampleName.Value = True
          m_badIniVal = True
        End If
      End If
    End If
  End If
  
  ' Setup user inputs output position
  If (m_rptUsrPos = DFLT_USR_POS) Then
    frm_dynRptCfg.opt_usrPosPre.Value = True
  Else
    frm_dynRptCfg.opt_usrPosPost.Value = True
  End If
  
  ' Check if ini file had bad value
  If (m_badIniVal = True) Then
    unity_main.errorstring = (fileName & " had incompatible value(s). Updated file with proper/default value(s).")
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg7", "%1 had incompatible value(s). Updated file with proper/default value(s)", fileName)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg3", "%1. Refer to system log for more details", uniMsg), vbOKOnly
    Call save_cfg(True, False)
  End If
  
  If (mustBeCfg = True) Then
    ' Copy dynamic report config into system operational variables
    unity_main.m_rptAddHeader = m_rptAddHeader
    unity_main.m_rptAddManPrefix = m_rptAddManPrefix
    unity_main.m_rptAddManSuffix = m_rptAddManSuffix
    unity_main.m_rptAddQuotes = m_rptAddQuotes
    unity_main.m_rptAddTrailer = m_rptAddTrailer
    unity_main.m_rptBaseCounter = m_rptBaseCounter
    unity_main.m_rptBaseName = m_rptBaseName
    unity_main.m_rptDate = m_rptDate
    unity_main.m_rptDateCounter = m_rptDateCounter
    
    If (m_rptFieldDelim = "SP") Then
      unity_main.m_rptFieldDelim = " "
    Else
      unity_main.m_rptFieldDelim = m_rptFieldDelim
    End If
    
    unity_main.m_rptFileExt = m_rptFileExt
    unity_main.m_rptFileFormat = m_rptFileFormat
    unity_main.m_rptFilePath = m_rptFilePath
   
    For ii = 1 To MAX_RPT_FIELDS
      RptHdrFieldTxt(ii) = m_rptHdrFieldTxt(ii)
      RptHdrFieldType(ii) = m_rptHdrFieldType(ii)
      RptHdrMaxChars(ii) = m_rptHdrMaxChars(ii)
    Next ii
    
    unity_main.m_rptHdrNumFields = m_rptHdrNumFields
    unity_main.m_rptManPrefix = m_rptManPrefix
    unity_main.m_rptManSuffix = m_rptManSuffix
    unity_main.m_rptNameMode = m_rptNameMode
 
    For ii = 1 To MAX_RPT_FIELDS
      RptRecFieldTxt(ii) = m_rptRecFieldTxt(ii)
      RptRecFieldType(ii) = m_rptRecFieldType(ii)
      RptRecMaxChars(ii) = m_rptRecMaxChars(ii)
    Next ii
    
    unity_main.m_rptRecNumFields = m_rptRecNumFields

    For ii = 1 To MAX_RPT_FIELDS
      RptTrlFieldTxt(ii) = m_rptTrlFieldTxt(ii)
      RptTrlFieldType(ii) = m_rptTrlFieldType(ii)
      RptTrlMaxChars(ii) = m_rptTrlMaxChars(ii)
    Next ii
    
    unity_main.m_rptTrlNumFields = m_rptTrlNumFields

    For ii = 1 To MAX_RPT_FIELDS
      RptUsrFieldTxt(ii) = m_rptUsrFieldTxt(ii)
      RptUsrFieldType(ii) = m_rptUsrFieldType(ii)
      RptUsrMaxChars(ii) = m_rptUsrMaxChars(ii)
    Next ii
    
    unity_main.m_rptUsrNumFields = m_rptUsrNumFields
    unity_main.m_rptUsrPos = m_rptUsrPos
  End If
End Sub

Private Function load_dyn_rpt_file_vals(srcFile As String) As Boolean
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
            m_fileVersion = Trim(varVal)
          Case "addheader"
            m_rptAddHeader = CBool(varVal)
          Case "addmanprefix"
            m_rptAddManPrefix = CBool(varVal)
          Case "addmansuffix"
            m_rptAddManSuffix = CBool(varVal)
          Case "addquotes"
            m_rptAddQuotes = CBool(varVal)
          Case "addtrailer"
            m_rptAddTrailer = CBool(varVal)
          Case "basecounter"
            m_rptBaseCounter = CLng(varVal)
          Case "basename"
            m_rptBaseName = varVal
          Case "datecounter"
            m_rptDateCounter = CLng(varVal)
          Case "fielddelim"
            m_rptFieldDelim = varVal
          Case "fileext"
            m_rptFileExt = varVal
          Case "fileformat"
            m_rptFileFormat = varVal
          Case "filepath"
            m_rptFilePath = varVal
          Case "hdrfieldtxt1"
            m_rptHdrFieldTxt(1) = varVal
          Case "hdrfieldtxt2"
            m_rptHdrFieldTxt(2) = varVal
          Case "hdrfieldtxt3"
            m_rptHdrFieldTxt(3) = varVal
          Case "hdrfieldtxt4"
            m_rptHdrFieldTxt(4) = varVal
          Case "hdrfieldtxt5"
            m_rptHdrFieldTxt(5) = varVal
          Case "hdrfieldtxt6"
            m_rptHdrFieldTxt(6) = varVal
          Case "hdrfieldtxt7"
            m_rptHdrFieldTxt(7) = varVal
          Case "hdrfieldtxt8"
            m_rptHdrFieldTxt(8) = varVal
          Case "hdrfieldtxt9"
            m_rptHdrFieldTxt(9) = varVal
          Case "hdrfieldtype1"
            m_rptHdrFieldType(1) = varVal
          Case "hdrfieldtype2"
            m_rptHdrFieldType(2) = varVal
          Case "hdrfieldtype3"
            m_rptHdrFieldType(3) = varVal
          Case "hdrfieldtype4"
            m_rptHdrFieldType(4) = varVal
          Case "hdrfieldtype5"
            m_rptHdrFieldType(5) = varVal
          Case "hdrfieldtype6"
            m_rptHdrFieldType(6) = varVal
          Case "hdrfieldtype7"
            m_rptHdrFieldType(7) = varVal
          Case "hdrfieldtype8"
            m_rptHdrFieldType(8) = varVal
          Case "hdrfieldtype9"
            m_rptHdrFieldType(9) = varVal
          Case "hdrmaxchars1"
            m_rptHdrMaxChars(1) = CInt(varVal)
          Case "hdrmaxchars2"
            m_rptHdrMaxChars(2) = CInt(varVal)
          Case "hdrmaxchars3"
            m_rptHdrMaxChars(3) = CInt(varVal)
          Case "hdrmaxchars4"
            m_rptHdrMaxChars(4) = CInt(varVal)
          Case "hdrmaxchars5"
            m_rptHdrMaxChars(5) = CInt(varVal)
          Case "hdrmaxchars6"
            m_rptHdrMaxChars(6) = CInt(varVal)
          Case "hdrmaxchars7"
            m_rptHdrMaxChars(7) = CInt(varVal)
          Case "hdrmaxchars8"
            m_rptHdrMaxChars(8) = CInt(varVal)
          Case "hdrmaxchars9"
            m_rptHdrMaxChars(9) = CInt(varVal)
          Case "hdrnumfields"
            m_rptHdrNumFields = CInt(varVal)
          Case "manualprefix"
            m_rptManPrefix = varVal
          Case "manualsuffix"
            m_rptManSuffix = varVal
          Case "namemode"
            m_rptNameMode = varVal
          Case "recfieldtxt1"
            m_rptRecFieldTxt(1) = varVal
          Case "recfieldtxt2"
            m_rptRecFieldTxt(2) = varVal
          Case "recfieldtxt3"
            m_rptRecFieldTxt(3) = varVal
          Case "recfieldtxt4"
            m_rptRecFieldTxt(4) = varVal
          Case "recfieldtxt5"
            m_rptRecFieldTxt(5) = varVal
          Case "recfieldtxt6"
            m_rptRecFieldTxt(6) = varVal
          Case "recfieldtxt7"
            m_rptRecFieldTxt(7) = varVal
          Case "recfieldtxt8"
            m_rptRecFieldTxt(8) = varVal
          Case "recfieldtxt9"
            m_rptRecFieldTxt(9) = varVal
          Case "recfieldtype1"
            m_rptRecFieldType(1) = varVal
          Case "recfieldtype2"
            m_rptRecFieldType(2) = varVal
          Case "recfieldtype3"
            m_rptRecFieldType(3) = varVal
          Case "recfieldtype4"
            m_rptRecFieldType(4) = varVal
          Case "recfieldtype5"
            m_rptRecFieldType(5) = varVal
          Case "recfieldtype6"
            m_rptRecFieldType(6) = varVal
          Case "recfieldtype7"
            m_rptRecFieldType(7) = varVal
          Case "recfieldtype8"
            m_rptRecFieldType(8) = varVal
          Case "recfieldtype9"
            m_rptRecFieldType(9) = varVal
          Case "recmaxchars1"
            m_rptRecMaxChars(1) = CInt(varVal)
          Case "recmaxchars2"
            m_rptRecMaxChars(2) = CInt(varVal)
          Case "recmaxchars3"
            m_rptRecMaxChars(3) = CInt(varVal)
          Case "recmaxchars4"
            m_rptRecMaxChars(4) = CInt(varVal)
          Case "recmaxchars5"
            m_rptRecMaxChars(5) = CInt(varVal)
          Case "recmaxchars6"
            m_rptRecMaxChars(6) = CInt(varVal)
          Case "recmaxchars7"
            m_rptRecMaxChars(7) = CInt(varVal)
          Case "recmaxchars8"
            m_rptRecMaxChars(8) = CInt(varVal)
          Case "recmaxchars9"
            m_rptRecMaxChars(9) = CInt(varVal)
          Case "recnumfields"
            m_rptRecNumFields = CInt(varVal)
          Case "reportdate"
            m_rptDate = varVal
          Case "trlfieldtxt1"
            m_rptTrlFieldTxt(1) = varVal
          Case "trlfieldtxt2"
            m_rptTrlFieldTxt(2) = varVal
          Case "trlfieldtxt3"
            m_rptTrlFieldTxt(3) = varVal
          Case "trlfieldtxt4"
            m_rptTrlFieldTxt(4) = varVal
          Case "trlfieldtxt5"
            m_rptTrlFieldTxt(5) = varVal
          Case "trlfieldtxt6"
            m_rptTrlFieldTxt(6) = varVal
          Case "trlfieldtxt7"
            m_rptTrlFieldTxt(7) = varVal
          Case "trlfieldtxt8"
            m_rptTrlFieldTxt(8) = varVal
          Case "trlfieldtxt9"
            m_rptTrlFieldTxt(9) = varVal
          Case "trlfieldtype1"
            m_rptTrlFieldType(1) = varVal
          Case "trlfieldtype2"
            m_rptTrlFieldType(2) = varVal
          Case "trlfieldtype3"
            m_rptTrlFieldType(3) = varVal
          Case "trlfieldtype4"
            m_rptTrlFieldType(4) = varVal
          Case "trlfieldtype5"
            m_rptTrlFieldType(5) = varVal
          Case "trlfieldtype6"
            m_rptTrlFieldType(6) = varVal
          Case "trlfieldtype7"
            m_rptTrlFieldType(7) = varVal
          Case "trlfieldtype8"
            m_rptTrlFieldType(8) = varVal
          Case "trlfieldtype9"
            m_rptTrlFieldType(9) = varVal
          Case "trlmaxchars1"
            m_rptTrlMaxChars(1) = CInt(varVal)
          Case "trlmaxchars2"
            m_rptTrlMaxChars(2) = CInt(varVal)
          Case "trlmaxchars3"
            m_rptTrlMaxChars(3) = CInt(varVal)
          Case "trlmaxchars4"
            m_rptTrlMaxChars(4) = CInt(varVal)
          Case "trlmaxchars5"
            m_rptTrlMaxChars(5) = CInt(varVal)
          Case "trlmaxchars6"
            m_rptTrlMaxChars(6) = CInt(varVal)
          Case "trlmaxchars7"
            m_rptTrlMaxChars(7) = CInt(varVal)
          Case "trlmaxchars8"
            m_rptTrlMaxChars(8) = CInt(varVal)
          Case "trlmaxchars9"
            m_rptTrlMaxChars(9) = CInt(varVal)
          Case "trlnumfields"
            m_rptTrlNumFields = CInt(varVal)
          Case "usrfieldtxt1"
            m_rptUsrFieldTxt(1) = varVal
          Case "usrfieldtxt2"
            m_rptUsrFieldTxt(2) = varVal
          Case "usrfieldtxt3"
            m_rptUsrFieldTxt(3) = varVal
          Case "usrfieldtxt4"
            m_rptUsrFieldTxt(4) = varVal
          Case "usrfieldtxt5"
            m_rptUsrFieldTxt(5) = varVal
          Case "usrfieldtxt6"
            m_rptUsrFieldTxt(6) = varVal
          Case "usrfieldtxt7"
            m_rptUsrFieldTxt(7) = varVal
          Case "usrfieldtxt8"
            m_rptUsrFieldTxt(8) = varVal
          Case "usrfieldtxt9"
            m_rptUsrFieldTxt(9) = varVal
          Case "usrfieldtype1"
            m_rptUsrFieldType(1) = varVal
          Case "usrfieldtype2"
            m_rptUsrFieldType(2) = varVal
          Case "usrfieldtype3"
            m_rptUsrFieldType(3) = varVal
          Case "usrfieldtype4"
            m_rptUsrFieldType(4) = varVal
          Case "usrfieldtype5"
            m_rptUsrFieldType(5) = varVal
          Case "usrfieldtype6"
            m_rptUsrFieldType(6) = varVal
          Case "usrfieldtype7"
            m_rptUsrFieldType(7) = varVal
          Case "usrfieldtype8"
            m_rptUsrFieldType(8) = varVal
          Case "usrfieldtype9"
            m_rptUsrFieldType(9) = varVal
          Case "usrmaxchars1"
            m_rptUsrMaxChars(1) = CInt(varVal)
          Case "usrmaxchars2"
            m_rptUsrMaxChars(2) = CInt(varVal)
          Case "usrmaxchars3"
            m_rptUsrMaxChars(3) = CInt(varVal)
          Case "usrmaxchars4"
            m_rptUsrMaxChars(4) = CInt(varVal)
          Case "usrmaxchars5"
            m_rptUsrMaxChars(5) = CInt(varVal)
          Case "usrmaxchars6"
            m_rptUsrMaxChars(6) = CInt(varVal)
          Case "usrmaxchars7"
            m_rptUsrMaxChars(7) = CInt(varVal)
          Case "usrmaxchars8"
            m_rptUsrMaxChars(8) = CInt(varVal)
          Case "usrmaxchars9"
            m_rptUsrMaxChars(9) = CInt(varVal)
          Case "usrnumfields"
            m_rptUsrNumFields = CInt(varVal)
          Case "usrpos"
            m_rptUsrPos = varVal
        End Select
      End If
    Wend
  
    load_dyn_rpt_file_vals = True
  Else
FILE_ERROR:
    errMsg = (srcFile & " file read error. " & Error$)
    unity_main.errorstring = errMsg
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg4", "%1 file read error. %2", srcFile, Error$)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    load_dyn_rpt_file_vals = False
  End If
  
  uniFile.CloseFile
  Exit Function
  
BAD_INI_VALUE:
  unity_main.errorstring = (srcFile & " line " & CStr(lineCnt) & " had incompatible value. " & Trim(tmpStrg) & " - " & varVal & "; will use default value")
  unity_main.write_error
  m_badIniVal = True
  Resume Next
End Function

Public Sub save_cfg(saveBtn As Boolean, setupCfg As Boolean)
  Dim ii As Integer
  Dim filePathName As String
  Dim errMsg As String
  Dim uniFile As New clsUniFile
  Dim uniMsg As String
  
  ' Check if config files needs to be saved due to user request or name counter update
  If (saveBtn = True) Or (m_saveRptIni = True) Then
    m_saveRptIni = False
  
    ' Save dynamic report variables
    If (chk_addHeader.Value <> 0) Then
      m_rptAddHeader = True
    Else
      m_rptAddHeader = False
    End If
  
    If (chk_addManualPrefix.Value <> 0) Then
      m_rptAddManPrefix = True
    Else
      m_rptAddManPrefix = False
    End If
  
    If (chk_addManualSuffix.Value <> 0) Then
      m_rptAddManSuffix = True
    Else
      m_rptAddManSuffix = False
    End If
  
    If (chk_addQuotes.Value <> 0) Then
      m_rptAddQuotes = True
    Else
      m_rptAddQuotes = False
    End If
  
    If (chk_addTrailer.Value <> 0) Then
      m_rptAddTrailer = True
    Else
      m_rptAddTrailer = False
    End If
  
    m_rptDate = unity_main.m_rptDate
    m_rptBaseCounter = numInc_baseCounter.Text
  
    ' Save report base file name
    m_rptBaseName = Trim(txt_baseName.Text)
  
    m_rptDateCounter = numInc_dateCounter.Text
    m_rptFieldDelim = combo_fieldDelim.List(frm_dynRptCfg.combo_fieldDelim.ListIndex)
    m_rptFileExt = Trim(txt_fileExt.Text)
  
    ' Save file format
    If (opt_asciiFormat = True) Then
      m_rptFileFormat = "ASCII"
    Else
      m_rptFileFormat = "Unicode"
    End If

    ' Confirm report filepath contains '\' instead of '/'
    filePathName = Trim(frm_dynRptCfg.txt_filePath.Text)
    check_filepathname_delimiters filePathName
    frm_dynRptCfg.txt_filePath.Text = filePathName
    
    ' Append "\" to report file path if not present
    If (Right(frm_dynRptCfg.txt_filePath.Text, 1) <> "\") Then
      frm_dynRptCfg.txt_filePath.Text = frm_dynRptCfg.txt_filePath.Text & "\"
    End If
    
    m_rptFilePath = frm_dynRptCfg.txt_filePath.Text
  
    For ii = 1 To MAX_RPT_FIELDS
      m_rptHdrFieldTxt(ii) = Trim(frm_dynRptCfg.txt_hdrFieldTxt(ii).Text)
      m_rptHdrFieldType(ii) = frm_dynRptCfg.combo_hdrFieldType(ii).List(frm_dynRptCfg.combo_hdrFieldType(ii).ListIndex)
      m_rptHdrMaxChars(ii) = frm_dynRptCfg.numInc_hdrMaxChars(ii).Text
    Next ii
  
    m_rptHdrNumFields = frm_dynRptCfg.numInc_hdrNumFields.Text
    m_rptManPrefix = Trim(frm_dynRptCfg.txt_manualPrefix.Text)
    m_rptManSuffix = Trim(frm_dynRptCfg.txt_manualSuffix.Text)
  
    ' Save report name type
    If (frm_dynRptCfg.opt_sampleName.Value = True) Then
      m_rptNameMode = "Sample"
    Else
      If (frm_dynRptCfg.opt_baseName.Value = True) Then
        m_rptNameMode = "Base"
      Else
        If (frm_dynRptCfg.opt_dateName.Value = True) Then
          m_rptNameMode = "Date"
        Else
          If (frm_dynRptCfg.opt_manualName.Value = True) Then
            m_rptNameMode = "Manual"
          End If
        End If
      End If
    End If
  
    For ii = 1 To MAX_RPT_FIELDS
      m_rptRecFieldTxt(ii) = Trim(frm_dynRptCfg.txt_recFieldTxt(ii).Text)
      m_rptRecFieldType(ii) = frm_dynRptCfg.combo_recFieldType(ii).List(frm_dynRptCfg.combo_recFieldType(ii).ListIndex)
      m_rptRecMaxChars(ii) = frm_dynRptCfg.numInc_recMaxChars(ii).Text
    Next ii
  
    m_rptRecNumFields = frm_dynRptCfg.numInc_recNumFields.Text

    For ii = 1 To MAX_RPT_FIELDS
      m_rptTrlFieldTxt(ii) = Trim(frm_dynRptCfg.txt_trlFieldTxt(ii).Text)
      m_rptTrlFieldType(ii) = frm_dynRptCfg.combo_trlFieldType(ii).List(frm_dynRptCfg.combo_trlFieldType(ii).ListIndex)
      m_rptTrlMaxChars(ii) = frm_dynRptCfg.numInc_trlMaxChars(ii).Text
    Next ii
  
    m_rptTrlNumFields = frm_dynRptCfg.numInc_trlNumFields.Text
    
    For ii = 1 To MAX_RPT_FIELDS
      m_rptUsrFieldTxt(ii) = Trim(frm_dynRptCfg.txt_usrFieldTxt(ii).Text)
      m_rptUsrFieldType(ii) = frm_dynRptCfg.combo_usrFieldType(ii).List(frm_dynRptCfg.combo_usrFieldType(ii).ListIndex)
      m_rptUsrMaxChars(ii) = frm_dynRptCfg.numInc_usrMaxChars(ii).Text
    Next ii
  
    m_rptUsrNumFields = numInc_usrNumFields.Text
  
    If (opt_usrPosPre.Value = True) Then
      m_rptUsrPos = "PreRecords"
    Else
      m_rptUsrPos = "PostRecords"
    End If
  
    If (uniFile.OpenFileWrite(CFG_DIR & DYN_REPORT_CFG_FILE) = True) Then
      On Error GoTo FILE_ERROR
      uniFile.WriteBOM fe_UTF16LE
      uniFile.WriteUnicodeLine ("Version=" & INFOSTAR_VER)
      uniFile.WriteUnicodeLine ("AddHeader=" & m_rptAddHeader)
      uniFile.WriteUnicodeLine ("AddManPrefix=" & m_rptAddManPrefix)
      uniFile.WriteUnicodeLine ("AddManSuffix=" & m_rptAddManSuffix)
      uniFile.WriteUnicodeLine ("AddQuotes=" & m_rptAddQuotes)
      uniFile.WriteUnicodeLine ("AddTrailer=" & m_rptAddTrailer)
      uniFile.WriteUnicodeLine ("BaseCounter=" & m_rptBaseCounter)
      uniFile.WriteUnicodeLine ("BaseName=" & m_rptBaseName)
      uniFile.WriteUnicodeLine ("DateCounter=" & m_rptDateCounter)
      uniFile.WriteUnicodeLine ("FieldDelim=" & m_rptFieldDelim)
      uniFile.WriteUnicodeLine ("FileExt=" & m_rptFileExt)
      uniFile.WriteUnicodeLine ("FileFormat=" & m_rptFileFormat)
      uniFile.WriteUnicodeLine ("FilePath=" & m_rptFilePath)
     
     For ii = 1 To MAX_RPT_FIELDS
        uniFile.WriteUnicodeLine ("HdrFieldTxt" & ii & "=" & m_rptHdrFieldTxt(ii))
        uniFile.WriteUnicodeLine ("HdrFieldType" & ii & "=" & m_rptHdrFieldType(ii))
        uniFile.WriteUnicodeLine ("HdrMaxChars" & ii & "=" & m_rptHdrMaxChars(ii))
      Next ii
      
      uniFile.WriteUnicodeLine ("HdrNumFields=" & m_rptHdrNumFields)
      uniFile.WriteUnicodeLine ("ManualPrefix=" & m_rptManPrefix)
      uniFile.WriteUnicodeLine ("ManualSuffix=" & m_rptManSuffix)
      uniFile.WriteUnicodeLine ("NameMode=" & m_rptNameMode)
    
      For ii = 1 To MAX_RPT_FIELDS
        uniFile.WriteUnicodeLine ("RecFieldTxt" & ii & "=" & m_rptRecFieldTxt(ii))
        uniFile.WriteUnicodeLine ("RecFieldType" & ii & "=" & m_rptRecFieldType(ii))
        uniFile.WriteUnicodeLine ("RecMaxChars" & ii & "=" & m_rptRecMaxChars(ii))
      Next ii
      
      uniFile.WriteUnicodeLine ("RecNumFields=" & m_rptRecNumFields)
      uniFile.WriteUnicodeLine ("ReportDate=" & m_rptDate)
    
      For ii = 1 To MAX_RPT_FIELDS
        uniFile.WriteUnicodeLine ("TrlFieldTxt" & ii & "=" & m_rptTrlFieldTxt(ii))
        uniFile.WriteUnicodeLine ("TrlFieldType" & ii & "=" & m_rptTrlFieldType(ii))
        uniFile.WriteUnicodeLine ("TrlMaxChars" & ii & "=" & m_rptTrlMaxChars(ii))
      Next ii
      
      uniFile.WriteUnicodeLine ("TrlNumFields=" & m_rptTrlNumFields)
    
      For ii = 1 To MAX_RPT_FIELDS
        uniFile.WriteUnicodeLine ("UsrFieldTxt" & ii & "=" & m_rptUsrFieldTxt(ii))
        uniFile.WriteUnicodeLine ("UsrFieldType" & ii & "=" & m_rptUsrFieldType(ii))
        uniFile.WriteUnicodeLine ("UsrMaxChars" & ii & "=" & m_rptUsrMaxChars(ii))
      Next ii
      
      uniFile.WriteUnicodeLine ("UsrNumFields=" & m_rptUsrNumFields)
      uniFile.WriteUnicodeLine ("UsrPos=" & m_rptUsrPos)
      uniFile.Flush
    
      If (setupCfg = True) Then
        ' Copy dynamic report config into system operational variables
        unity_main.m_rptAddHeader = m_rptAddHeader
        unity_main.m_rptAddManPrefix = m_rptAddManPrefix
        unity_main.m_rptAddManSuffix = m_rptAddManSuffix
        unity_main.m_rptAddQuotes = m_rptAddQuotes
        unity_main.m_rptAddTrailer = m_rptAddTrailer
        unity_main.m_rptBaseCounter = m_rptBaseCounter
        unity_main.m_rptBaseName = m_rptBaseName
        unity_main.m_rptDate = m_rptDate
        unity_main.m_rptDateCounter = m_rptDateCounter
    
        If (m_rptFieldDelim = "SP") Then
          unity_main.m_rptFieldDelim = " "
        Else
          unity_main.m_rptFieldDelim = m_rptFieldDelim
        End If
    
        unity_main.m_rptFileExt = m_rptFileExt
        unity_main.m_rptFileFormat = m_rptFileFormat
        unity_main.m_rptFilePath = m_rptFilePath
    
        For ii = 1 To MAX_RPT_FIELDS
          RptHdrFieldTxt(ii) = m_rptHdrFieldTxt(ii)
          RptHdrFieldType(ii) = m_rptHdrFieldType(ii)
          RptHdrMaxChars(ii) = m_rptHdrMaxChars(ii)
        Next ii
    
        unity_main.m_rptHdrNumFields = m_rptHdrNumFields
        unity_main.m_rptManPrefix = m_rptManPrefix
        unity_main.m_rptManSuffix = m_rptManSuffix
        unity_main.m_rptNameMode = m_rptNameMode
  
        For ii = 1 To MAX_RPT_FIELDS
          RptRecFieldTxt(ii) = m_rptRecFieldTxt(ii)
          RptRecFieldType(ii) = m_rptRecFieldType(ii)
          RptRecMaxChars(ii) = m_rptRecMaxChars(ii)
        Next ii
    
        unity_main.m_rptRecNumFields = m_rptRecNumFields

        For ii = 1 To MAX_RPT_FIELDS
          RptTrlFieldTxt(ii) = m_rptTrlFieldTxt(ii)
          RptTrlFieldType(ii) = m_rptTrlFieldType(ii)
          RptTrlMaxChars(ii) = m_rptTrlMaxChars(ii)
        Next ii
    
        unity_main.m_rptTrlNumFields = m_rptTrlNumFields

        For ii = 1 To MAX_RPT_FIELDS
          RptUsrFieldTxt(ii) = m_rptUsrFieldTxt(ii)
          RptUsrFieldType(ii) = m_rptUsrFieldType(ii)
          RptUsrMaxChars(ii) = m_rptUsrMaxChars(ii)
        Next ii
    
        unity_main.m_rptUsrNumFields = m_rptUsrNumFields
        unity_main.m_rptUsrPos = m_rptUsrPos
      End If
    Else
FILE_ERROR:
      errMsg = ((CFG_DIR & DYN_REPORT_CFG_FILE) & " file write error. " & Error$)
      unity_main.errorstring = errMsg
      unity_main.write_error
      uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", (CFG_DIR & DYN_REPORT_CFG_FILE), Error$)
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    End If
  
    uniFile.CloseFile
  End If
End Sub

Public Sub show_frame(n As Integer)
  Dim ii As Integer
  
  For ii = 0 To frame_main.Count - 1
    If ii = n Then
      frame_main(ii).Visible = True
    Else
      frame_main(ii).Visible = False
    End If
  Next
End Sub

Public Sub write_dynamic_report()
  Dim fileName As String
  Dim uniFile As New clsUniFile
  Dim hdrStrg As String
  Dim recStrgs() As String
  Dim trlStrg As String
  Dim usrStrg As String
  Dim numprops As Integer
  Dim ii As Integer
  Dim errMsg As String
  Dim uniMsg As String

  ' Save system date & time
  m_dateTime = Now
  m_date = Date
  m_time = Time

  ' Build report file name
  If (build_rpt_file_name(fileName) = False) Then
    Exit Sub
  End If

  ' Check if to build header
  If (unity_main.m_rptAddHeader = True) Then
    Call build_rpt_header(hdrStrg)
  End If
  
  ' Check if to build user inputs
  If (unity_main.m_rptUsrNumFields <> 0) Then
    Call build_rpt_usr_inputs(usrStrg)
  End If
  
  ' Build property records
  numprops = frmedmod.numprops.Text
  Call build_rpt_records(recStrgs)
  
  ' Check if to build trailer
  If (unity_main.m_rptAddTrailer = True) Then
    Call build_rpt_trailer(trlStrg)
  End If
  
  CreatePath CFile.st_FilePath(unity_main.m_rptFilePath)
  
  ' Open report file for writing
  If (uniFile.OpenFileWrite(unity_main.m_rptFilePath & fileName) = True) Then
    On Error GoTo FILE_ERROR
    
    ' Check if to write file in ASCII format
    If (unity_main.m_rptFileFormat = "ASCII") Then
      ' Check if to add header to report
      If (unity_main.m_rptAddHeader = True) Then
        uniFile.WriteAnsiLine hdrStrg
      End If

      ' Check if to add user inputs to report before record section
      If (unity_main.m_rptUsrNumFields <> 0) And (unity_main.m_rptUsrPos = DFLT_USR_POS) Then
        uniFile.WriteAnsiLine usrStrg
      End If

      ' Write property records to report
      For ii = 1 To numprops
        uniFile.WriteAnsiLine recStrgs(ii)
      Next ii
      
      ' Check if to add user inputs to report after record section
      If (unity_main.m_rptUsrNumFields <> 0) And (unity_main.m_rptUsrPos <> DFLT_USR_POS) Then
        uniFile.WriteAnsiLine usrStrg
      End If

      ' Check if to add trailer to report
      If (unity_main.m_rptAddTrailer = True) Then
        uniFile.WriteAnsiLine trlStrg
      End If
    Else    ' Unicode format
      uniFile.WriteBOM fe_UTF16LE
      
      ' Check if to add header to report
      If (unity_main.m_rptAddHeader = True) Then
        uniFile.WriteUnicodeLine hdrStrg
      End If

      ' Check if to add user inputs to report before record section
      If (unity_main.m_rptUsrNumFields <> 0) And (unity_main.m_rptUsrPos = DFLT_USR_POS) Then
        uniFile.WriteUnicodeLine usrStrg
      End If

      ' Write property records to report
      For ii = 1 To numprops
        uniFile.WriteUnicodeLine recStrgs(ii)
      Next ii
      
      ' Check if to add user inputs to report after record section
      If (unity_main.m_rptUsrNumFields <> 0) And (unity_main.m_rptUsrPos <> DFLT_USR_POS) Then
        uniFile.WriteUnicodeLine usrStrg
      End If

      ' Check if to add trailer to report
      If (unity_main.m_rptAddTrailer = True) Then
        uniFile.WriteUnicodeLine trlStrg
      End If
    End If
      
    uniFile.Flush
  Else
FILE_ERROR:
    errMsg = ((unity_main.m_rptFilePath & fileName) & " file write error. " & Error$)
    unity_main.errorstring = errMsg
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", (unity_main.m_rptFilePath & fileName), Error$)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  End If
  
  uniFile.CloseFile
End Sub

Private Function build_rpt_file_name(fileName As String) As Boolean
  Dim ctr As Long
  Dim datex As String
  Dim buildName As String
  Dim optVal As Integer
  Dim uniMsg As String

  ' Process file name based on mode
  Select Case (unity_main.m_rptNameMode)
    Case "Sample"
      fileName = unity_main.txtsamplename.Text
      
    Case "Base"
      m_saveRptIni = True     ' flag to save ini file
      
INC_BASE_CTR:
      ctr = frm_dynRptCfg.numInc_baseCounter.Text
      fileName = unity_main.m_rptBaseName & ctr & "." & unity_main.m_rptFileExt
        
      ' Check if file exists
      If (CFile.st_FileExist(unity_main.m_rptFilePath & fileName & "." & unity_main.m_rptFileExt) = True) Then
        ctr = ctr + 1
        
        If (ctr > frm_dynRptCfg.numInc_baseCounter.Max) Then
          ctr = frm_dynRptCfg.numInc_baseCounter.Min
        End If
        
        frm_dynRptCfg.numInc_baseCounter.Text = ctr
        GoTo INC_BASE_CTR
      End If
      
      fileName = unity_main.m_rptBaseName & frm_dynRptCfg.numInc_baseCounter.Text
          
      ' Setup for next report
      ctr = ctr + 1
      
      If (ctr > frm_dynRptCfg.numInc_baseCounter.Max) Then
        frm_dynRptCfg.numInc_baseCounter.Text = frm_dynRptCfg.numInc_baseCounter.Min
      End If
          
      frm_dynRptCfg.numInc_baseCounter.Text = ctr
    
    Case "Date"
      m_saveRptIni = True     ' flag to save ini file
      datex = Date
  
      ' Reset date counter if new day
      If (unity_main.m_rptDate <> datex) Then
        unity_main.m_rptDate = datex
        frm_dynRptCfg.numInc_dateCounter.Text = frm_dynRptCfg.numInc_dateCounter.Min
      End If
  
INC_DATE_CTR:
      Call frm_collect.rebuild_date(unity_main.m_rptDate, buildName)
      buildName = buildName & "_"
      ctr = frm_dynRptCfg.numInc_dateCounter.Text
      fileName = buildName & ctr
        
      ' Check if file exists
      If (CFile.st_FileExist(unity_main.m_rptFilePath & fileName & "." & unity_main.m_rptFileExt) = True) Then
        ctr = ctr + 1
        
        If (ctr > frm_dynRptCfg.numInc_dateCounter.Max) Then
          ctr = frm_dynRptCfg.numInc_dateCounter.Min
        End If
        
        frm_dynRptCfg.numInc_dateCounter.Text = ctr
        GoTo INC_DATE_CTR
      End If
        
      fileName = buildName & frm_dynRptCfg.numInc_dateCounter.Text
        
      ' Setup for next report
      ctr = ctr + 1
        
      If (ctr > frm_dynRptCfg.numInc_dateCounter.Max) Then
        ctr = frm_dynRptCfg.numInc_dateCounter.Min
      End If
        
      frm_dynRptCfg.numInc_dateCounter.Text = ctr
    
    Case "Manual"
ENTER_NAME:
      ' Prompt for report file name
      unity_main.m_rptNameEntry = ""
      frm_dynRptName.Show 1
  
      ' Check if user cancel report
      If (unity_main.m_rptNameEntry = "") Then
        build_rpt_file_name = False
        Exit Function
      End If
    
      ' Check if to add prefix to name
      If (unity_main.m_rptAddManPrefix = True) Then
        fileName = unity_main.m_rptManPrefix & unity_main.m_rptNameEntry
      Else
        fileName = unity_main.m_rptNameEntry
      End If
    
      ' Check if to add suffix to name
      If (unity_main.m_rptAddManSuffix = True) Then
        fileName = fileName & unity_main.m_rptManSuffix
      End If
      
      If (CFile.st_FileExist(unity_main.m_rptFilePath & fileName & "." & unity_main.m_rptFileExt) = True) Then
        uniMsg = MLSupport.GSS("frm_dynRptCfg", "statMsg1", "A dynamic report file with this name already exists, do you want to overwrite it?")
        optVal = CWrap.ShowMessageBoxW(uniMsg, vbYesNo)
    
        If (optVal = vbNo) Then
          GoTo ENTER_NAME
        End If
      End If
  End Select
  
  ' Add file name extension
  fileName = fileName & "." & unity_main.m_rptFileExt
  build_rpt_file_name = True
End Function

Private Sub build_rpt_header(hdrStrg As String)
  Dim ii As Integer
  Dim txt As String
  Dim numLines As Integer

  hdrStrg = ""
    
  For ii = 1 To unity_main.m_rptHdrNumFields
    Select Case (RptHdrFieldType(ii))
      Case "Text"
        txt = RptHdrFieldTxt(ii)
      Case "DateTime"
        txt = m_dateTime
      Case "Date"
        txt = m_date
      Case "Time"
        txt = m_time
      Case "YYMMDD"
        txt = Format(m_date, "yymmdd")
      Case "YYYYMMDD"
        txt = Format(m_date, "yyyymmdd")
      Case "HHMMSS"
        txt = Format(m_time, "hhmmss")
      Case "ProdName"
        txt = Trim(unity_main.lblProd1.Caption)
      Case "SampID"
        txt = Trim(unity_main.txtsamplename.Text)
      Case "Comment"
        txt = Trim(unity_main.txtsampcomment.Text)
      Case "ManEntry"
        txt = unity_main.m_rptNameEntry
      Case "SerNum"
        txt = unity_main.m_sysSerialNum
      Case "Input1"
        Call get_input_data(txt, 1)
      Case "Input2"
        Call get_input_data(txt, 2)
      Case "Input3"
        Call get_input_data(txt, 3)
      Case "Input4"
        Call get_input_data(txt, 4)
      Case "Input5"
        Call get_input_data(txt, 5)
      Case "Input6"
        Call get_input_data(txt, 6)
      Case "Input7"
        Call get_input_data(txt, 7)
      Case "Input8"
        Call get_input_data(txt, 8)
      Case "NumRecs"
        txt = frmedmod.numprops.Text
      Case "NumLines"
        numLines = frmedmod.numprops.Text
        
        If (unity_main.m_rptAddHeader = True) Then
          numLines = numLines + 1
        End If
        
        If (unity_main.m_rptUsrNumFields <> 0) Then
          numLines = numLines + 1
        End If
        
        If (unity_main.m_rptAddTrailer = True) Then
          numLines = numLines + 1
        End If
        
        txt = CStr(numLines)
    End Select
  
    Call ck_max_chars(txt, RptHdrMaxChars(ii))
    
    ' Check if to enclose data in quotes
    If (unity_main.m_rptAddQuotes = True) Then
      hdrStrg = hdrStrg & (Chr(34) & txt & Chr(34))
    Else
      hdrStrg = hdrStrg & txt
    End If
    
    ' Ck if to add delimiter char
    If (ii <> unity_main.m_rptHdrNumFields) Then
      hdrStrg = hdrStrg & unity_main.m_rptFieldDelim
    End If
  Next ii
End Sub
    
Private Sub build_rpt_usr_inputs(usrStrg As String)
  Dim ii As Integer
  Dim txt As String

  usrStrg = ""
    
  For ii = 1 To unity_main.m_rptUsrNumFields
    Select Case (RptUsrFieldType(ii))
      Case "Text"
        txt = RptUsrFieldTxt(ii)
      Case "Input1"
        Call get_input_data(txt, 1)
      Case "Input2"
        Call get_input_data(txt, 2)
      Case "Input3"
        Call get_input_data(txt, 3)
      Case "Input4"
        Call get_input_data(txt, 4)
      Case "Input5"
        Call get_input_data(txt, 5)
      Case "Input6"
        Call get_input_data(txt, 6)
      Case "Input7"
        Call get_input_data(txt, 7)
      Case "Input8"
        Call get_input_data(txt, 8)
    End Select
  
    Call ck_max_chars(txt, RptUsrMaxChars(ii))
    
    ' Check if to enclose data in quotes
    If (unity_main.m_rptAddQuotes = True) Then
      usrStrg = usrStrg & (Chr(34) & txt & Chr(34))
    Else
      usrStrg = usrStrg & txt
    End If
    
    ' Ck if to add delimiter char
    If (ii <> unity_main.m_rptUsrNumFields) Then
      usrStrg = usrStrg & unity_main.m_rptFieldDelim
    End If
  Next ii
End Sub

Private Sub build_rpt_records(recStrgs() As String)
  Dim numprops As Integer
  Dim ii As Integer
  Dim jj As Integer
  Dim txt As String

  numprops = frmedmod.numprops.Text
  ReDim recStrgs(1 To numprops)

  For jj = 1 To numprops
    recStrgs(jj) = ""
 
    For ii = 1 To unity_main.m_rptRecNumFields
      Select Case (RptRecFieldType(ii))
        Case "Text"
          txt = RptRecFieldTxt(ii)
        Case "PropName"
          unity_main.fpspread_pred.Row = jj
          unity_main.fpspread_pred.Col = 1
          txt = unity_main.fpspread_pred.Text
        Case "PropVal"
          unity_main.fpspread_pred.Row = jj
          unity_main.fpspread_pred.Col = 2
          txt = unity_main.fpspread_pred.Text
        Case "PropMDst"
          If (unity_main.lstmd.ListCount = 0) Then
            txt = unity_main.m_noOLVal
          Else
            txt = unity_main.lstmd.List(jj - 1)
          End If
        Case "PropSRes"
          If (unity_main.lstresrat.ListCount = 0) Then
            txt = unity_main.m_noOLVal
          Else
            txt = unity_main.lstresrat.List(jj - 1)
          End If
        Case "PropOutl"
          If (unity_main.lst_qual.ListCount = 0) Then
            txt = unity_main.m_noOLVal
          Else
            txt = unity_main.lst_qual.List(jj - 1)
          End If
        Case "PropND"
          If (unity_main.lst_nd.ListCount = 0) Then
            txt = unity_main.m_noOLVal
          Else
            txt = unity_main.lst_nd.List(jj - 1)
          End If
        Case "PropIntc"
          If (unity_main.lstint.ListCount = 0) Then
            txt = unity_main.m_noOLVal
          Else
            txt = unity_main.lstint.List(jj - 1)
          End If
        Case "PropSlop"
          If (unity_main.lstslope.ListCount = 0) Then
            txt = unity_main.m_noOLVal
          Else
            txt = unity_main.lstslope.List(jj - 1)
          End If
      End Select
    
      Call ck_max_chars(txt, RptRecMaxChars(ii))
      
      ' Check if to enclose data in quotes
      If (unity_main.m_rptAddQuotes = True) Then
        recStrgs(jj) = recStrgs(jj) & (Chr(34) & txt & Chr(34))
      Else
        recStrgs(jj) = recStrgs(jj) & txt
      End If
    
      ' Ck if to add delimiter char
      If (ii <> unity_main.m_rptRecNumFields) Then
        recStrgs(jj) = recStrgs(jj) & unity_main.m_rptFieldDelim
      End If
    Next ii
  Next jj
End Sub

Private Sub build_rpt_trailer(trlStrg As String)
  Dim ii As Integer
  Dim txt As String
  Dim numLines As Integer

  trlStrg = ""
    
  For ii = 1 To unity_main.m_rptTrlNumFields
    Select Case (RptTrlFieldType(ii))
      Case "Text"
        txt = RptTrlFieldTxt(ii)
      Case "DateTime"
        txt = m_dateTime
      Case "Date"
        txt = m_date
      Case "Time"
        txt = m_time
      Case "YYMMDD"
        txt = Format(m_date, "yymmdd")
      Case "YYYYMMDD"
        txt = Format(m_date, "yyyymmdd")
      Case "HHMMSS"
        txt = Format(m_time, "hhmmss")
      Case "ProdName"
        txt = Trim(unity_main.lblProd1.Caption)
      Case "SampID"
        txt = Trim(unity_main.txtsamplename.Text)
      Case "Comment"
        txt = Trim(unity_main.txtsampcomment.Text)
      Case "ManEntry"
        txt = unity_main.m_rptNameEntry
      Case "SerNum"
        txt = unity_main.m_sysSerialNum
      Case "Input1"
        Call get_input_data(txt, 1)
      Case "Input2"
        Call get_input_data(txt, 2)
      Case "Input3"
        Call get_input_data(txt, 3)
      Case "Input4"
        Call get_input_data(txt, 4)
      Case "Input5"
        Call get_input_data(txt, 5)
      Case "Input6"
        Call get_input_data(txt, 6)
      Case "Input7"
        Call get_input_data(txt, 7)
      Case "Input8"
        Call get_input_data(txt, 8)
      Case "NumRecs"
        txt = frmedmod.numprops.Text
      Case "NumLines"
        numLines = frmedmod.numprops.Text
        
        If (unity_main.m_rptAddHeader = True) Then
          numLines = numLines + 1
        End If
        
        If (unity_main.m_rptUsrNumFields <> 0) Then
          numLines = numLines + 1
        End If
        
        If (unity_main.m_rptAddTrailer = True) Then
          numLines = numLines + 1
        End If
        
        txt = CStr(numLines)
    End Select
  
    Call ck_max_chars(txt, RptTrlMaxChars(ii))
    
    ' Check if to enclose data in quotes
    If (unity_main.m_rptAddQuotes = True) Then
      trlStrg = trlStrg & (Chr(34) & txt & Chr(34))
    Else
      trlStrg = trlStrg & txt
    End If
    
    ' Ck if to add delimiter char
    If (ii <> unity_main.m_rptTrlNumFields) Then
      trlStrg = trlStrg & unity_main.m_rptFieldDelim
    End If
  Next ii
End Sub

Private Sub get_input_data(txt As String, inpNum As Integer)

  ' Setup for input enable field
  frm_buttoncfg.ss_buttonconfig.Col = inpNum
  frm_buttoncfg.ss_buttonconfig.Row = 1
  
  ' Check if input enabled
  If (unity_main.m_useMIV = True) And (frm_buttoncfg.ss_buttonconfig.Value = 1) Then
    ' Setup for text entry/list box selection field
    frm_buttoncfg.ss_buttonconfig.Col = inpNum
    frm_buttoncfg.ss_buttonconfig.Row = 2
  
    ' Check if using text entry
    If (frm_buttoncfg.ss_buttonconfig.Value = 0) Then
      txt = Trim(frm_scanname.txtbx(inpNum).Text)
    Else    ' Using list
      txt = Trim(frm_scanname.combo(inpNum).Text)
    End If
  Else
    txt = MLSupport.GSS("Headers", "na", "NA")
  End If
End Sub

Private Sub ck_max_chars(txt As String, maxChars As Integer)
  
  If (Len(txt) > maxChars) Then
    txt = Left(txt, maxChars)
  End If
End Sub

Private Function check_settings() As Boolean
  Dim numFields As Integer
  Dim ii As Integer

  check_settings = False

  ' Check Report Settings parameters
  If (frm_dynRptCfg.txt_filePath.Text = "") Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_dynRptCfg", "errMsg1", "Please enter a file directory in Report Settings Tab"), vbExclamation
    Exit Function
  End If
  
  If (frm_dynRptCfg.txt_fileExt.Text = "") Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_dynRptCfg", "errMsg2", "Please enter a file extension in Report Settings Tab"), vbExclamation
    Exit Function
  End If
  
  If (frm_dynRptCfg.opt_manualName.Value = True) And (frm_dynRptCfg.chk_addManualPrefix.Value <> 0) And (frm_dynRptCfg.txt_manualPrefix.Text = "") Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_dynRptCfg", "errMsg4", "Please enter a file name prefix in Report Settings Tab"), vbExclamation
    Exit Function
  End If
  
  If (frm_dynRptCfg.opt_manualName.Value = True) And (frm_dynRptCfg.chk_addManualSuffix.Value <> 0) And (frm_dynRptCfg.txt_manualSuffix.Text = "") Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_dynRptCfg", "errMsg5", "Please enter a file name suffix in Report Settings Tab"), vbExclamation
    Exit Function
  End If
  
  ' Check Header Settings parameters if configured
  If (frm_dynRptCfg.chk_addHeader.Value <> 0) Then
    numFields = frm_dynRptCfg.numInc_hdrNumFields.Text
    
    If (numFields = 0) Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_dynRptCfg", "errMsg6", "Please configure at least one data field in Header Settings Tab or uncheck 'Add header to report' in Report Settings Tab"), vbExclamation
      Exit Function
    End If
    
    For ii = 1 To numFields
      Select Case (frm_dynRptCfg.combo_hdrFieldType(ii).Text)
        Case "Text"
          If (frm_dynRptCfg.txt_hdrFieldTxt(ii).Text = "") Then
            CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_dynRptCfg.errMsg2", "Please enter text for data field %1 in Header Settings Tab", CStr(ii)), vbExclamation
            Exit Function
          End If
        
          If (frm_dynRptCfg.numInc_hdrMaxChars(ii).Text < Len(frm_dynRptCfg.txt_hdrFieldTxt(ii).Text)) Then
            CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_dynRptCfg.errMsg3", "Please increase max characters to %1 for data field %2 in Header Settings Tab", CStr(Len(txt_hdrFieldTxt(ii).Text)), CStr(ii)), vbExclamation
            Exit Function
          End If
        Case "DateTime"
        Case "Date"
        Case "Time"
        Case "YYMMDD"
          If (frm_dynRptCfg.numInc_hdrMaxChars(ii).Text < Len(frm_dynRptCfg.combo_hdrFieldType(ii).Text)) Then
            CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_dynRptCfg.errMsg3", "Please increase max characters to %1 for data field %2 in Header Settings Tab", CStr(Len(combo_hdrFieldType(ii).Text)), CStr(ii)), vbExclamation
            Exit Function
          End If
        Case "YYYYMMDD"
          If (frm_dynRptCfg.numInc_hdrMaxChars(ii).Text < Len(frm_dynRptCfg.combo_hdrFieldType(ii).Text)) Then
            CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_dynRptCfg.errMsg3", "Please increase max characters to %1 for data field %2 in Header Settings Tab", CStr(Len(combo_hdrFieldType(ii).Text)), CStr(ii)), vbExclamation
            Exit Function
          End If
        Case "HHMMSS"
          If (frm_dynRptCfg.numInc_hdrMaxChars(ii).Text < Len(frm_dynRptCfg.combo_hdrFieldType(ii).Text)) Then
            CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_dynRptCfg.errMsg3", "Please increase max characters to %1 for data field %2 in Header Settings Tab", CStr(Len(combo_hdrFieldType(ii).Text)), CStr(ii)), vbExclamation
            Exit Function
          End If
        Case "ProdName"
        Case "SampID"
        Case "Comment"
        Case "ManEntry"
        Case "SerNum"
          If (frm_dynRptCfg.numInc_hdrMaxChars(ii).Text < 4) Then
            CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_dynRptCfg.errMsg3", "Please increase max characters to %1 for data field %2 in Header Settings Tab", CStr(4), CStr(ii)), vbExclamation
            Exit Function
          End If
        Case "Input1"
        Case "Input2"
        Case "Input3"
        Case "Input4"
        Case "Input5"
        Case "Input6"
        Case "Input7"
        Case "Input8"
        Case "NumRecs"
        Case "NumLines"
      End Select
    Next ii
  End If
  
  ' Check User Inputs Settings parameters
  numFields = frm_dynRptCfg.numInc_usrNumFields.Text
    
  If (numFields <> 0) Then
    For ii = 1 To numFields
      Select Case (frm_dynRptCfg.combo_usrFieldType(ii).Text)
        Case "Text"
          If (frm_dynRptCfg.txt_usrFieldTxt(ii).Text = "") Then
            CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_dynRptCfg.errMsg4", "Please enter text for data field %1 in User Inputs Settings Tab", CStr(ii)), vbExclamation
            Exit Function
          End If
        
          If (frm_dynRptCfg.numInc_usrMaxChars(ii).Text < Len(frm_dynRptCfg.txt_usrFieldTxt(ii).Text)) Then
            CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_dynRptCfg.errMsg5", "Please increase max characters to %1 for data field %2 in User Inputs Settings Tab", CStr(Len(txt_recFieldTxt(ii).Text)), CStr(ii)), vbExclamation
            Exit Function
          End If
        Case "Input1"
        Case "Input2"
        Case "Input3"
        Case "Input4"
        Case "Input5"
        Case "Input6"
        Case "Input7"
        Case "Input8"
      End Select
    Next ii
  End If
  
  ' Check Record Settings parameters
  numFields = frm_dynRptCfg.numInc_recNumFields.Text
    
  If (numFields = 0) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_dynRptCfg", "errMsg7", "Please configure at least one data field in Record Settings Tab"), vbExclamation
    Exit Function
  End If
    
  For ii = 1 To numFields
    Select Case (frm_dynRptCfg.combo_recFieldType(ii).Text)
      Case "Text"
        If (frm_dynRptCfg.txt_recFieldTxt(ii).Text = "") Then
          CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_dynRptCfg.errMsg6", "Please enter text for data field %1 in Record Settings Tab", CStr(ii)), vbExclamation
          Exit Function
        End If
        
        If (frm_dynRptCfg.numInc_recMaxChars(ii).Text < Len(frm_dynRptCfg.txt_recFieldTxt(ii).Text)) Then
          CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_dynRptCfg.errMsg7", "Please increase max characters to %1 for data field %2 in Record Settings Tab", CStr(Len(txt_recFieldTxt(ii).Text)), CStr(ii)), vbExclamation
          Exit Function
        End If
      Case "PropName"
      Case "PropVal"
      Case "PropMDst"
      Case "PropSRes"
      Case "PropOutl"
      Case "PropND"
      Case "PropIntc"
      Case "PropSlop"
    End Select
  Next ii
  
  ' Check Trailer Settings parameters if configured
  If (frm_dynRptCfg.chk_addTrailer.Value <> 0) Then
    numFields = frm_dynRptCfg.numInc_trlNumFields.Text
    
    If (numFields = 0) Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_dynRptCfg", "errMsg8", "Please configure at least one data field in Trailer Settings Tab or uncheck 'Add trailer to report' in Report Settings Tab"), vbExclamation
      Exit Function
    End If
    
    For ii = 1 To numFields
      Select Case (frm_dynRptCfg.combo_trlFieldType(ii).Text)
        Case "Text"
          If (frm_dynRptCfg.txt_trlFieldTxt(ii).Text = "") Then
            CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_dynRptCfg.errMsg8", "Please enter text for data field %1 in Trailer Settings Tab", CStr(ii)), vbExclamation
            Exit Function
          End If
        
          If (frm_dynRptCfg.numInc_trlMaxChars(ii).Text < Len(frm_dynRptCfg.txt_trlFieldTxt(ii).Text)) Then
            CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_dynRptCfg.errMsg9", "Please increase max characters to %1 for data field %2 in Trailer Settings Tab", CStr(Len(txt_trlFieldTxt(ii).Text)), CStr(ii)), vbExclamation
            Exit Function
          End If
        Case "DateTime"
        Case "Date"
        Case "Time"
        Case "YYMMDD"
          If (frm_dynRptCfg.numInc_trlMaxChars(ii).Text < Len(frm_dynRptCfg.combo_trlFieldType(ii).Text)) Then
            CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_dynRptCfg.errMsg7", "Please increase max characters to %1 for data field %2 in Trailer Settings Tab", CStr(Len(combo_trlFieldType(ii).Text)), CStr(ii)), vbExclamation
            Exit Function
          End If
        Case "YYYYMMDD"
          If (frm_dynRptCfg.numInc_trlMaxChars(ii).Text < Len(frm_dynRptCfg.combo_trlFieldType(ii).Text)) Then
            CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_dynRptCfg.errMsg7", "Please increase max characters to %1 for data field %2 in Trailer Settings Tab", CStr(Len(combo_trlFieldType(ii).Text)), CStr(ii)), vbExclamation
            Exit Function
          End If
        Case "HHMMSS"
          If (frm_dynRptCfg.numInc_trlMaxChars(ii).Text < Len(frm_dynRptCfg.combo_trlFieldType(ii).Text)) Then
            CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_dynRptCfg.errMsg7", "Please increase max characters to %1 for data field %2 in Trailer Settings Tab", CStr(Len(combo_trlFieldType(ii).Text)), CStr(ii)), vbExclamation
            Exit Function
          End If
        Case "ProdName"
        Case "SampID"
        Case "Comment"
        Case "ManEntry"
        Case "SerNum"
          If (frm_dynRptCfg.numInc_trlMaxChars(ii).Text < 4) Then
            CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_dynRptCfg.errMsg7", "Please increase max characters to %1 for data field %2 in Trailer Settings Tab", CStr(4), CStr(ii)), vbExclamation
            Exit Function
          End If
        Case "Input1"
        Case "Input2"
        Case "Input3"
        Case "Input4"
        Case "Input5"
        Case "Input6"
        Case "Input7"
        Case "Input8"
        Case "NumRecs"
        Case "NumLines"
      End Select
    Next ii
  End If
  
  check_settings = True
End Function

Private Sub chk_addHeader_Click()

  If (chk_addHeader.Value <> 0) Then
    frm_dynRptCfg.frame_main(1).Enabled = True
  Else
    frm_dynRptCfg.frame_main(1).Enabled = False
  End If
End Sub

Private Sub chk_addManualPrefix_Click()

  If (m_rptNameMode = "Manual") Then
    opt_manualName_Click
  End If
End Sub

Private Sub chk_addManualSuffix_Click()

  If (m_rptNameMode = "Manual") Then
    opt_manualName_Click
  End If
End Sub

Private Sub chk_addTrailer_Click()

  If (chk_addTrailer.Value <> 0) Then
    frm_dynRptCfg.frame_main(3).Enabled = True
  Else
    frm_dynRptCfg.frame_main(3).Enabled = False
  End If
End Sub

Private Sub cmd_cancel_Click()

  unity_main.errorstring = "Dynamic Report Configuration screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  frm_dynRptCfg.Visible = False
End Sub

Private Sub cmd_save_Click()
  
  ' Check if valid configuration
  If (check_settings = True) Then
    unity_main.errorstring = "Dynamic Report Configuration screen Save Changes button selected"
    unity_main.write_error (LOG_DBG_LEVEL3)
    
    Call save_cfg(True, True)
    frm_dynRptCfg.Visible = False
  End If
End Sub

Private Sub combo_hdrFieldType_Click(Index As Integer)

  Select Case (combo_hdrFieldType(Index).Text)
    Case "Text"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = True
    Case "DateTime"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "Date"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "Time"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "YYMMDD"
      numInc_hdrMaxChars(Index).Text = 6
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "YYYYMMDD"
      numInc_hdrMaxChars(Index).Text = 8
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "HHMMSS"
      numInc_hdrMaxChars(Index).Text = 6
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "ProdName"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "SampID"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "Comment"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "ManEntry"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "SerNum"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "Input1"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "Input2"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "Input3"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "Input4"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "Input5"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "Input6"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "Input7"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "Input8"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "NumRecs"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
    Case "NumLines"
      frm_dynRptCfg.txt_hdrFieldTxt(Index).Visible = False
  End Select
End Sub

Private Sub combo_recFieldType_Click(Index As Integer)

  Select Case (combo_recFieldType(Index).Text)
    Case "Text"
      frm_dynRptCfg.txt_recFieldTxt(Index).Visible = True
    Case "PropName"
      frm_dynRptCfg.txt_recFieldTxt(Index).Visible = False
    Case "PropVal"
      frm_dynRptCfg.txt_recFieldTxt(Index).Visible = False
    Case "PropMDst"
      frm_dynRptCfg.txt_recFieldTxt(Index).Visible = False
    Case "PropSRes"
      frm_dynRptCfg.txt_recFieldTxt(Index).Visible = False
    Case "PropOutl"
      frm_dynRptCfg.txt_recFieldTxt(Index).Visible = False
    Case "PropND"
      frm_dynRptCfg.txt_recFieldTxt(Index).Visible = False
    Case "PropIntc"
      frm_dynRptCfg.txt_recFieldTxt(Index).Visible = False
    Case "PropSlop"
      frm_dynRptCfg.txt_recFieldTxt(Index).Visible = False
  End Select
End Sub

Private Sub combo_trlFieldType_Click(Index As Integer)

  Select Case (combo_trlFieldType(Index).Text)
    Case "Text"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = True
    Case "DateTime"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "Date"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "Time"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "YYMMDD"
      numInc_trlMaxChars(Index).Text = 6
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "YYYYMMDD"
      numInc_trlMaxChars(Index).Text = 8
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "HHMMSS"
      numInc_trlMaxChars(Index).Text = 6
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "ProdName"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "SampID"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "Comment"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "ManEntry"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "SerNum"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "Input1"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "Input2"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "Input3"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "Input4"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "Input5"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "Input6"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "Input7"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "Input8"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "NumRecs"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
    Case "NumLines"
      frm_dynRptCfg.txt_trlFieldTxt(Index).Visible = False
  End Select
End Sub

Private Sub combo_usrFieldType_Click(Index As Integer)

  Select Case (combo_usrFieldType(Index).Text)
    Case "Text"
      frm_dynRptCfg.txt_usrFieldTxt(Index).Visible = True
    Case "Input1"
      frm_dynRptCfg.txt_usrFieldTxt(Index).Visible = False
    Case "Input2"
      frm_dynRptCfg.txt_usrFieldTxt(Index).Visible = False
    Case "Input3"
      frm_dynRptCfg.txt_usrFieldTxt(Index).Visible = False
    Case "Input4"
      frm_dynRptCfg.txt_usrFieldTxt(Index).Visible = False
    Case "Input5"
      frm_dynRptCfg.txt_usrFieldTxt(Index).Visible = False
    Case "Input6"
      frm_dynRptCfg.txt_usrFieldTxt(Index).Visible = False
    Case "Input7"
      frm_dynRptCfg.txt_usrFieldTxt(Index).Visible = False
    Case "Input8"
      frm_dynRptCfg.txt_usrFieldTxt(Index).Visible = False
  End Select
End Sub

Private Sub Form_Load()
  Dim ii As Integer
  Dim jj As Integer
  Dim strg As String

  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
  ' Setup tab headers
  frm_dynRptCfg.tab_frame.AddTab MLSupport.GSS("frm_dynRptCfg", "TabCaption0", "Report Settings")
  frm_dynRptCfg.tab_frame.AddTab MLSupport.GSS("frm_dynRptCfg", "TabCaption1", "Header Settings")
  frm_dynRptCfg.tab_frame.AddTab MLSupport.GSS("frm_dynRptCfg", "TabCaption2", "User Inputs Settings")
  frm_dynRptCfg.tab_frame.AddTab MLSupport.GSS("frm_dynRptCfg", "TabCaption3", "Record Settings")
  frm_dynRptCfg.tab_frame.AddTab MLSupport.GSS("frm_dynRptCfg", "TabCaption4", "Trailer Settings")
  
  ' Setup data field frame labels
  For ii = 1 To MAX_RPT_FIELDS
    strg = MLSupport.GSS("frm_dynRptCfg", "frame_field", "Data Field") & " " & ii
    frm_dynRptCfg.frame_hdr(ii).Caption = strg
    frm_dynRptCfg.frame_rec(ii).Caption = strg
    frm_dynRptCfg.frame_trl(ii).Caption = strg
  Next ii
  
  ' Setup data field delimiter selections
  frm_dynRptCfg.combo_fieldDelim.AddItem ("SP")
  frm_dynRptCfg.combo_fieldDelim.AddItem (",")
  frm_dynRptCfg.combo_fieldDelim.AddItem (";")
  frm_dynRptCfg.combo_fieldDelim.AddItem (":")
  frm_dynRptCfg.combo_fieldDelim.AddItem ("\")
  frm_dynRptCfg.combo_fieldDelim.AddItem ("\\")
  frm_dynRptCfg.combo_fieldDelim.AddItem ("/")
  frm_dynRptCfg.combo_fieldDelim.AddItem ("//")
  
  ' Setup header data field type selections
  For ii = 1 To MAX_RPT_FIELDS
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("Text")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("DateTime")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("Date")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("Time")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("YYMMDD")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("YYYYMMDD")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("HHMMSS")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("ProdName")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("SampID")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("Comment")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("ManEntry")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("SerNum")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("Input1")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("Input2")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("Input3")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("Input4")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("Input5")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("Input6")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("Input7")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("Input8")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("NumRecs")
    frm_dynRptCfg.combo_hdrFieldType(ii).AddItem ("NumLines")
  Next ii
  
  ' Setup user input data field type selections
  For ii = 1 To MAX_RPT_FIELDS
    frm_dynRptCfg.combo_usrFieldType(ii).AddItem ("Text")
    frm_dynRptCfg.combo_usrFieldType(ii).AddItem ("Input1")
    frm_dynRptCfg.combo_usrFieldType(ii).AddItem ("Input2")
    frm_dynRptCfg.combo_usrFieldType(ii).AddItem ("Input3")
    frm_dynRptCfg.combo_usrFieldType(ii).AddItem ("Input4")
    frm_dynRptCfg.combo_usrFieldType(ii).AddItem ("Input5")
    frm_dynRptCfg.combo_usrFieldType(ii).AddItem ("Input6")
    frm_dynRptCfg.combo_usrFieldType(ii).AddItem ("Input7")
    frm_dynRptCfg.combo_usrFieldType(ii).AddItem ("Input8")
  Next ii
  
  ' Setup record data field type selections
  For ii = 1 To MAX_RPT_FIELDS
    frm_dynRptCfg.combo_recFieldType(ii).AddItem ("Text")
    frm_dynRptCfg.combo_recFieldType(ii).AddItem ("PropName")
    frm_dynRptCfg.combo_recFieldType(ii).AddItem ("PropVal")
    frm_dynRptCfg.combo_recFieldType(ii).AddItem ("PropMDst")
    frm_dynRptCfg.combo_recFieldType(ii).AddItem ("PropSRes")
    frm_dynRptCfg.combo_recFieldType(ii).AddItem ("PropOutl")
    frm_dynRptCfg.combo_recFieldType(ii).AddItem ("PropND")
    frm_dynRptCfg.combo_recFieldType(ii).AddItem ("PropIntc")
    frm_dynRptCfg.combo_recFieldType(ii).AddItem ("PropSlop")
  Next ii
  
  ' Setup trailer data field type selections
  For ii = 1 To MAX_RPT_FIELDS
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("Text")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("DateTime")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("Date")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("Time")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("YYMMDD")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("YYYYMMDD")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("HHMMSS")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("ProdName")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("SampID")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("Comment")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("ManEntry")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("SerNum")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("Input1")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("Input2")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("Input3")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("Input4")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("Input5")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("Input6")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("Input7")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("Input8")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("NumRecs")
    frm_dynRptCfg.combo_trlFieldType(ii).AddItem ("NumLines")
  Next ii
  
  ' Setup to display first tab frame
  show_frame 0
End Sub

Private Sub numInc_baseCounter_Change()

  If (m_rptNameMode = "Base") Then
    opt_baseName_Click
  End If
End Sub

Private Sub numInc_baseCounter_DblClick()

  unity_main.formfrom = 18
  unity_main.varfrom = 5
  frm_numpad.lbl_num.Caption = MLSupport.GSS("frm_dynRptCfg", "lbl1", "Base Name Counter Value")
  frm_numpad.txt_num.Text = frm_dynRptCfg.numInc_baseCounter.Text
  frm_numpad.Show 1
End Sub
  
Private Sub numInc_dateCounter_Change()
  
  If (m_rptNameMode = "Date") Then
    opt_dateName_Click
  End If
End Sub

Private Sub numInc_dateCounter_DblClick()

  unity_main.formfrom = 18
  unity_main.varfrom = 6
  frm_numpad.lbl_num.Caption = MLSupport.GSS("frm_dynRptCfg", "lbl2", "Date Name Counter Value")
  frm_numpad.txt_num.Text = frm_dynRptCfg.numInc_dateCounter.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_hdrMaxChars_DblClick(Index As Integer)

  unity_main.formfrom = 18
  unity_main.varfrom = 10 + Index
  frm_numpad.lbl_num.Caption = MLSupport.GSS("frm_dynRptCfg", "lbl_hdrMaxChrs(" & Str(Index) & ")", "Max Characters")
  frm_numpad.txt_num.Text = frm_dynRptCfg.numInc_hdrMaxChars(Index).Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_hdrNumFields_Change()
  Dim numFields As Integer
  Dim ii As Integer
  Dim jj As Integer

  numFields = frm_dynRptCfg.numInc_hdrNumFields.Text
  
  For ii = 1 To numFields
    frm_dynRptCfg.frame_hdr(ii).Visible = True
  Next ii
  
  For jj = ii To MAX_RPT_FIELDS
    frm_dynRptCfg.frame_hdr(jj).Visible = False
  Next jj
End Sub

Private Sub numInc_hdrNumFields_DblClick()

  unity_main.formfrom = 18
  unity_main.varfrom = 7
  frm_numpad.lbl_num.Caption = MLSupport.GSS("frm_dynRptCfg", "lbl_hdrNumFields", "Number of Data Fields")
  frm_numpad.txt_num.Text = frm_dynRptCfg.numInc_hdrNumFields.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_recMaxChars_DblClick(Index As Integer)

  unity_main.formfrom = 18
  unity_main.varfrom = 30 + Index
  frm_numpad.lbl_num.Caption = MLSupport.GSS("frm_dynRptCfg", "lbl_recMaxChrs(" & Str(Index) & ")", "Max Characters")
  frm_numpad.txt_num.Text = frm_dynRptCfg.numInc_recMaxChars(Index).Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_recNumFields_Change()
  Dim numFields As Integer
  Dim ii As Integer
  Dim jj As Integer

  numFields = frm_dynRptCfg.numInc_recNumFields.Text
  
  For ii = 1 To numFields
    frm_dynRptCfg.frame_rec(ii).Visible = True
  Next ii
  
  For jj = ii To MAX_RPT_FIELDS
    frm_dynRptCfg.frame_rec(jj).Visible = False
  Next jj
End Sub

Private Sub numInc_recNumFields_DblClick()

  unity_main.formfrom = 18
  unity_main.varfrom = 9
  frm_numpad.lbl_num.Caption = MLSupport.GSS("frm_dynRptCfg", "lbl_recNumFields", "Number of Data Fields")
  frm_numpad.txt_num.Text = frm_dynRptCfg.numInc_recNumFields.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_trlMaxChars_DblClick(Index As Integer)

  unity_main.formfrom = 18
  unity_main.varfrom = 40 + Index
  frm_numpad.lbl_num.Caption = MLSupport.GSS("frm_dynRptCfg", "lbl_trlMaxChrs(" & Str(Index) & ")", "Max Characters")
  frm_numpad.txt_num.Text = frm_dynRptCfg.numInc_trlMaxChars(Index).Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_trlNumFields_Change()
  Dim numFields As Integer
  Dim ii As Integer
  Dim jj As Integer

  numFields = frm_dynRptCfg.numInc_trlNumFields.Text
  
  For ii = 1 To numFields
    frm_dynRptCfg.frame_trl(ii).Visible = True
  Next ii
  
  For jj = ii To MAX_RPT_FIELDS
    frm_dynRptCfg.frame_trl(jj).Visible = False
  Next jj
End Sub

Private Sub numInc_trlNumFields_DblClick()

  unity_main.formfrom = 18
  unity_main.varfrom = 10
  frm_numpad.lbl_num.Caption = MLSupport.GSS("frm_dynRptCfg", "lbl_trlNumFields", "Number of Data Fields")
  frm_numpad.txt_num.Text = frm_dynRptCfg.numInc_trlNumFields.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_usrMaxChars_DblClick(Index As Integer)

  unity_main.formfrom = 18
  unity_main.varfrom = 20 + Index
  frm_numpad.lbl_num.Caption = MLSupport.GSS("frm_dynRptCfg", "lbl_usrMaxChrs(" & Str(Index) & ")", "Max Characters")
  frm_numpad.txt_num.Text = frm_dynRptCfg.numInc_usrMaxChars(Index).Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_usrNumFields_Change()
  Dim numFields As Integer
  Dim ii As Integer
  Dim jj As Integer

  numFields = frm_dynRptCfg.numInc_usrNumFields.Text
  
  If (numFields = 0) Then
    frm_dynRptCfg.opt_usrPosPre.Visible = False
    frm_dynRptCfg.opt_usrPosPost.Visible = False
  Else
    frm_dynRptCfg.opt_usrPosPre.Visible = True
    frm_dynRptCfg.opt_usrPosPost.Visible = True
  End If
  
  For ii = 1 To numFields
    frm_dynRptCfg.frame_usr(ii).Visible = True
  Next ii
  
  For jj = ii To MAX_RPT_FIELDS
    frm_dynRptCfg.frame_usr(jj).Visible = False
  Next jj

End Sub

Private Sub numInc_usrNumFields_DblClick()
  
  unity_main.formfrom = 18
  unity_main.varfrom = 8
  frm_numpad.lbl_num.Caption = MLSupport.GSS("frm_dynRptCfg", "lbl_usrNumFields", "Number of Data Fields")
  frm_numpad.txt_num.Text = frm_dynRptCfg.numInc_usrNumFields.Text
  frm_numpad.Show 1
End Sub

Private Sub opt_asciiFormat_Click()

  m_rptFileFormat = "ASCII"
End Sub

Private Sub opt_baseName_Click()

  frm_dynRptCfg.txt_rptName.Text = Trim(frm_dynRptCfg.txt_baseName.Text) & frm_dynRptCfg.numInc_baseCounter.Text & "." & Trim(frm_dynRptCfg.txt_fileExt.Text)
  m_rptNameMode = "Base"
End Sub

Private Sub opt_dateName_Click()
  Dim dateStrg As String
  
  Call frm_collect.rebuild_date(Date, dateStrg)
  frm_dynRptCfg.txt_rptName.Text = dateStrg & "_" & frm_dynRptCfg.numInc_dateCounter.Text & "." & Trim(frm_dynRptCfg.txt_fileExt.Text)
  m_rptNameMode = "Date"
End Sub

Private Sub opt_manualName_Click()
  Dim rptName As String

  If (frm_dynRptCfg.chk_addManualPrefix.Value <> 0) Then
    rptName = Trim(txt_manualPrefix.Text)
  End If
  
  rptName = rptName & MLSupport.GSS("frm_dynRptCfg", "txt_rptName", "Manual Entry")
  
  If (frm_dynRptCfg.chk_addManualSuffix.Value <> 0) Then
    rptName = rptName & Trim(txt_manualSuffix.Text)
  End If
  
  frm_dynRptCfg.txt_rptName.Text = rptName & "." & Trim(txt_fileExt.Text)
  m_rptNameMode = "Manual"
End Sub

Private Sub opt_sampleName_Click()

  txt_rptName.Text = MLSupport.GSS("frm_dynRptCfg", "txt_sampleName", "Sample ID") & "." & Trim(txt_fileExt.Text)
  m_rptNameMode = "Sample"
End Sub

Private Sub opt_uniFormat_Click()

  m_rptFileFormat = "Unicode"
End Sub

Private Sub tab_frame_TabChanged(Index As Integer)

  show_frame (Index - 1)
End Sub

Private Sub txt_baseName_DblClick(Button As Integer)

  unity_main.formfrom = 18
  unity_main.varfrom = 2
  frm_kybd.lbl_kybd.Caption = frm_dynRptCfg.lbl_baseName.Caption
  frm_kybd.txt_kybd.Text = frm_dynRptCfg.txt_baseName.Text
  frm_kybd.Show 1
End Sub

Private Sub txt_fileExt_DblClick(Button As Integer)
  
  unity_main.formfrom = 18
  unity_main.varfrom = 1
  frm_kybd.lbl_kybd.Caption = frm_dynRptCfg.lbl_fileExt.Caption
  frm_kybd.txt_kybd.Text = frm_dynRptCfg.txt_fileExt.Text
  frm_kybd.Show 1
End Sub

Private Sub txt_filePath_DblClick(Button As Integer)

  unity_main.formfrom = 18
  unity_main.varfrom = 0
  frm_kybd.lbl_kybd.Caption = frm_dynRptCfg.lbl_filePath.Caption
  frm_kybd.txt_kybd.Text = frm_dynRptCfg.txt_filePath.Text
  frm_kybd.Show 1
End Sub

Private Sub txt_hdrFieldTxt_DblClick(Index As Integer, Button As Integer)

  unity_main.formfrom = 18
  unity_main.varfrom = 50 + Index
  frm_kybd.lbl_kybd.Caption = MLSupport.GSS("frm_dynRptCfg", "lbl_hdrFieldTxt", "Header Text")
  frm_kybd.txt_kybd.Text = frm_dynRptCfg.txt_hdrFieldTxt(Index).Text
  frm_kybd.Show 1
End Sub

Private Sub txt_manualPrefix_Change()

  If (m_rptNameMode = "Manual") Then
    opt_manualName_Click
  End If
End Sub

Private Sub txt_manualPrefix_DblClick(Button As Integer)

  unity_main.formfrom = 18
  unity_main.varfrom = 3
  frm_kybd.lbl_kybd.Caption = MLSupport.GSS("frm_dynRptCfg", "lbl3", "File Name Prefix")
  frm_kybd.txt_kybd.Text = frm_dynRptCfg.txt_manualPrefix.Text
  frm_kybd.Show 1
End Sub

Private Sub txt_manualSuffix_Change()

  If (m_rptNameMode = "Manual") Then
    opt_manualName_Click
  End If
End Sub

Private Sub txt_manualSuffix_DblClick(Button As Integer)

  unity_main.formfrom = 18
  unity_main.varfrom = 4
  frm_kybd.lbl_kybd.Caption = MLSupport.GSS("frm_dynRptCfg", "lbl4", "File Name Suffix")
  frm_kybd.txt_kybd.Text = frm_dynRptCfg.txt_manualSuffix.Text
  frm_kybd.Show 1
End Sub

Private Sub txt_recFieldTxt_DblClick(Index As Integer, Button As Integer)

  unity_main.formfrom = 18
  unity_main.varfrom = 70 + Index
  frm_kybd.lbl_kybd.Caption = MLSupport.GSS("frm_dynRptCfg", "lbl_recFieldTxt", "Record Text")
  frm_kybd.txt_kybd.Text = frm_dynRptCfg.txt_recFieldTxt(Index).Text
  frm_kybd.Show 1
End Sub

Private Sub txt_trlFieldTxt_DblClick(Index As Integer, Button As Integer)

  unity_main.formfrom = 18
  unity_main.varfrom = 80 + Index
  frm_kybd.lbl_kybd.Caption = MLSupport.GSS("frm_dynRptCfg", "lbl_trlFieldTxt", "Trailer Text")
  frm_kybd.txt_kybd.Text = frm_dynRptCfg.txt_trlFieldTxt(Index).Text
  frm_kybd.Show 1
End Sub

Private Sub txt_usrFieldTxt_DblClick(Index As Integer, Button As Integer)

  unity_main.formfrom = 18
  unity_main.varfrom = 60 + Index
  frm_kybd.lbl_kybd.Caption = MLSupport.GSS("frm_dynRptCfg", "lbl_usrFieldTxt", "User Inputs Text")
  frm_kybd.txt_kybd.Text = frm_dynRptCfg.txt_usrFieldTxt(Index).Text
  frm_kybd.Show 1
End Sub








