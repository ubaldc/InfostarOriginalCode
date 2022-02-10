VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{488D84E9-A9A9-4C74-949B-F7752803ADEE}#99.2#0"; "EMDVirtualKeyBoard.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_edbias 
   Caption         =   "Bias Configuration"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
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
   Icon            =   "frm_edbias.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   765
      Left            =   7800
      TabIndex        =   17
      Top             =   9120
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1349
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
      Caption         =   "frm_edbias.frx":030A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_edbias.frx":0336
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_edbias.frx":0356
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   765
      Left            =   7800
      TabIndex        =   19
      Top             =   8160
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1349
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
      Caption         =   "frm_edbias.frx":0372
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_edbias.frx":03AA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_edbias.frx":03CA
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_newValue 
      Height          =   750
      Left            =   7800
      TabIndex        =   16
      Top             =   7185
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1323
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
      Caption         =   "frm_edbias.frx":03E6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_edbias.frx":0424
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_edbias.frx":0444
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_newval 
      Height          =   495
      Left            =   7665
      TabIndex        =   15
      Top             =   6360
      Width           =   2200
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_edbias.frx":0460
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
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
      Tip             =   "frm_edbias.frx":0480
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_edbias.frx":04A0
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_oldval 
      Height          =   555
      Left            =   7680
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5500
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   979
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      Enabled         =   0   'False
      Locked          =   0   'False
      Text            =   "frm_edbias.frx":04BC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
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
      Tip             =   "frm_edbias.frx":04DC
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_edbias.frx":04FC
   End
   Begin HexUniControls.ctlUniFrameXP Frame1 
      Height          =   5775
      Left            =   480
      Top             =   4680
      Width           =   3855
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
      Caption         =   "frm_edbias.frx":0518
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483642
      Tip             =   "frm_edbias.frx":0550
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_edbias.frx":0570
      Begin HexUniControls.ctlUniButtonImageXP cmd0 
         Height          =   795
         Left            =   240
         TabIndex        =   12
         Top             =   4680
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1402
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_edbias.frx":058C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_edbias.frx":05AE
         Style           =   -1
         RoundedBorders  =   -1  'True
         TransparentColor=   0
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_edbias.frx":05CE
      End
      Begin HexUniControls.ctlUniButtonImageXP cmd1 
         Height          =   795
         Left            =   240
         TabIndex        =   9
         Top             =   3600
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1402
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_edbias.frx":05EA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_edbias.frx":060C
         Style           =   -1
         RoundedBorders  =   -1  'True
         TransparentColor=   0
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_edbias.frx":062C
      End
      Begin HexUniControls.ctlUniButtonImageXP cmd8 
         Height          =   795
         Left            =   1440
         TabIndex        =   4
         Top             =   1440
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1402
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_edbias.frx":0648
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_edbias.frx":066A
         Style           =   -1
         RoundedBorders  =   -1  'True
         TransparentColor=   0
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_edbias.frx":068A
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdcomma 
         Height          =   800
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   1402
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_edbias.frx":06A6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_edbias.frx":06C8
         Style           =   -1
         RoundedBorders  =   -1  'True
         TransparentColor=   0
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_edbias.frx":06E8
      End
      Begin HexUniControls.ctlUniButtonImageXP cmddecimal 
         Height          =   795
         Left            =   2640
         TabIndex        =   14
         Top             =   4680
         Width           =   795
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_edbias.frx":0704
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_edbias.frx":0726
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_edbias.frx":0746
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdsign 
         Height          =   795
         Left            =   1440
         TabIndex        =   13
         Top             =   4680
         Width           =   795
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_edbias.frx":0762
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_edbias.frx":0788
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_edbias.frx":07A8
      End
      Begin HexUniControls.ctlUniButtonImageXP cmd3 
         Height          =   795
         Left            =   2640
         TabIndex        =   11
         Top             =   3600
         Width           =   795
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_edbias.frx":07C4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_edbias.frx":07E6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_edbias.frx":0806
      End
      Begin HexUniControls.ctlUniButtonImageXP cmd2 
         Height          =   795
         Left            =   1440
         TabIndex        =   10
         Top             =   3600
         Width           =   795
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_edbias.frx":0822
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_edbias.frx":0844
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_edbias.frx":0864
      End
      Begin HexUniControls.ctlUniButtonImageXP cmd6 
         Height          =   795
         Left            =   2640
         TabIndex        =   8
         Top             =   2520
         Width           =   795
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_edbias.frx":0880
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_edbias.frx":08A2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_edbias.frx":08C2
      End
      Begin HexUniControls.ctlUniButtonImageXP cmd5 
         Height          =   795
         Left            =   1440
         TabIndex        =   7
         Top             =   2520
         Width           =   795
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_edbias.frx":08DE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_edbias.frx":0900
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_edbias.frx":0920
      End
      Begin HexUniControls.ctlUniButtonImageXP cmd4 
         Height          =   795
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   795
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_edbias.frx":093C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_edbias.frx":095E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_edbias.frx":097E
      End
      Begin HexUniControls.ctlUniButtonImageXP cmd9 
         Height          =   795
         Left            =   2640
         TabIndex        =   5
         Top             =   1440
         Width           =   795
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_edbias.frx":099A
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_edbias.frx":09BC
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_edbias.frx":09DC
      End
      Begin HexUniControls.ctlUniButtonImageXP cmd7 
         Height          =   795
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   795
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_edbias.frx":09F8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_edbias.frx":0A1A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_edbias.frx":0A3A
      End
      Begin HexUniControls.ctlUniButtonImageXP cmd_bs 
         Height          =   800
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1402
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_edbias.frx":0A56
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_edbias.frx":0A88
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_edbias.frx":0AA8
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9240
      Top             =   10320
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   11010
      FormDesignWidth =   10335
   End
   Begin FPUSpreadADO.fpSpread ss_biases 
      Height          =   4095
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   9375
      _Version        =   458752
      _ExtentX        =   16536
      _ExtentY        =   7223
      _StockProps     =   64
      ColHeaderDisplay=   0
      DAutoHeadings   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      GrayAreaBackColor=   -2147483633
      MaxCols         =   3
      MaxRows         =   64
      OperationMode   =   1
      SelectBlockOptions=   2
      SpreadDesigner  =   "frm_edbias.frx":0AC4
      UserResize      =   1
   End
   Begin VBoard_EMD.KeySet_NumOnly KeySet_NumOnly1 
      Height          =   3165
      Left            =   6000
      Top             =   10920
      Visible         =   0   'False
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   5583
      FontName        =   "Arial Unicode MS"
      FontSize        =   12
      ButtonWidth     =   40
      ButtonHeight    =   40
      ShortRepeateDelay=   50
      LongRepeateDelay=   200
      SendKeysEnabled =   -1  'True
      RegAtDT         =   -1  'True
   End
   Begin HexUniControls.ctlUniLabel lbl_row 
      Height          =   495
      Left            =   5400
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_edbias.frx":131D
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_edbias.frx":133D
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_edbias.frx":135D
   End
   Begin HexUniControls.ctlUniLabel lbl_prop 
      Height          =   495
      Left            =   5520
      Top             =   4680
      Width           =   3975
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_edbias.frx":1379
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_edbias.frx":1399
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_edbias.frx":13B9
   End
   Begin HexUniControls.ctlUniLabel Label2 
      Height          =   495
      Left            =   5280
      Top             =   6360
      Width           =   2205
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_edbias.frx":13D5
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_edbias.frx":1407
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_edbias.frx":1427
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   495
      Left            =   5280
      Top             =   5520
      Width           =   2205
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_edbias.frx":1443
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_edbias.frx":147D
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_edbias.frx":149D
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   7680
      Top             =   10320
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
      Top             =   10320
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_edbias.frx":14B9
   End
End
Attribute VB_Name = "frm_edbias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub savebiases()
  Dim tempint As Integer
  Dim xx As Integer
  Dim tempstring As String
  Dim tempbias As Single
  
  frmedmod.grid_models.Col = 4  ' bias
  tempint = frmedmod.numprops.Text

  For xx = 1 To tempint
    frm_edbias.ss_biases.Row = xx
    frm_edbias.ss_biases.Col = 3
    tempbias = frm_edbias.ss_biases.Value
    frmedmod.grid_models.Row = xx
    frmedmod.grid_models.Text = tempbias
  Next xx
  
  frmedmod.runsavemodels
End Sub 'savebiases

Sub loadbiases()
  Dim tempint As Integer
  Dim yy As Integer
  Dim xx As Integer
  Dim tempstring As String
  Dim tempbias As Single

  frm_edbias.ss_biases.Row = 0
  frm_edbias.ss_biases.Col = 1
  frm_edbias.ss_biases.Text = MLSupport.GSS("Headers", "property", "Property")
  frm_edbias.ss_biases.Col = 2
  frm_edbias.ss_biases.Text = MLSupport.GSS("Headers", "curBias", "Current Bias")
  frm_edbias.ss_biases.Col = 3
  frm_edbias.ss_biases.Text = MLSupport.GSS("Headers", "newBias", "New Bias")
  tempint = frmedmod.numprops.Text
  frm_edbias.ss_biases.MaxRows = tempint
  
  For xx = 1 To tempint
    frmedmod.grid_models.Col = 1 ' property name
    frmedmod.grid_models.Row = xx
    tempstring = frmedmod.grid_models.Text
    frmedmod.grid_models.Col = 4 ' bias
    tempbias = frmedmod.grid_models.Value
    frm_edbias.ss_biases.Row = xx
    frm_edbias.ss_biases.Col = 1
    frm_edbias.ss_biases.Text = tempstring
    frm_edbias.ss_biases.Col = 2
    frm_edbias.ss_biases.Text = tempbias
    frm_edbias.ss_biases.Col = 3
    frm_edbias.ss_biases.Text = tempbias
  Next xx
  
  For yy = 1 To tempint
    frm_edbias.ss_biases.RowHeight(yy) = 20
  Next yy

  frm_edbias.ss_biases.Col = 1
  frm_edbias.ss_biases.Col2 = 2
  frm_edbias.ss_biases.Row = 1
  frm_edbias.ss_biases.Row2 = tempint
  frm_edbias.ss_biases.BlockMode = True
  frm_edbias.ss_biases.FontBold = True
  frm_edbias.ss_biases.BlockMode = False
  Call ss_biases_Click(1, 1)
End Sub  'loadbiases

Private Sub cmd_bs_Click()
  Dim tempstring As String
  Dim newstring As String
  Dim nchar As Integer

  tempstring = Trim(frm_edbias.txt_newval.Text)
  nchar = Len(tempstring)

  If (nchar = 0) Then
    Exit Sub
  End If
  
  newstring = Mid(tempstring, 1, (nchar - 1))
  frm_edbias.txt_newval.Text = newstring
End Sub

Private Sub cmd0_Click()
  
  frm_edbias.txt_newval.Text = Trim(frm_edbias.txt_newval.Text) & "0"
End Sub

Private Sub cmd1_Click()

  frm_edbias.txt_newval.Text = Trim(frm_edbias.txt_newval.Text) & "1"
End Sub

Private Sub cmd2_Click()

  frm_edbias.txt_newval.Text = Trim(frm_edbias.txt_newval.Text) & "2"
End Sub

Private Sub cmd3_Click()

  frm_edbias.txt_newval.Text = Trim(frm_edbias.txt_newval.Text) & "3"
End Sub

Private Sub cmd4_Click()

  frm_edbias.txt_newval.Text = Trim(frm_edbias.txt_newval.Text) & "4"
End Sub

Private Sub cmd5_Click()

  frm_edbias.txt_newval.Text = Trim(frm_edbias.txt_newval.Text) & "5"
End Sub

Private Sub cmd6_Click()

  frm_edbias.txt_newval.Text = Trim(frm_edbias.txt_newval.Text) & "6"
End Sub

Private Sub cmd7_Click()
  
  frm_edbias.txt_newval.Text = Trim(frm_edbias.txt_newval.Text) & "7"
End Sub

Private Sub cmd8_Click()

  frm_edbias.txt_newval.Text = Trim(frm_edbias.txt_newval.Text) & "8"
End Sub

Private Sub cmd9_Click()

  frm_edbias.txt_newval.Text = Trim(frm_edbias.txt_newval.Text) & "9"
End Sub

Private Sub cmdcomma_Click()

  frm_edbias.txt_newval.Text = Trim(frm_edbias.txt_newval.Text) & ","
End Sub

Private Sub cmddecimal_Click()

  frm_edbias.txt_newval.Text = Trim(frm_edbias.txt_newval.Text) & "."
End Sub

Private Sub cmdsign_Click()
  Dim tempstring As String
  Dim newstring As String
  Dim tempval As Single
  Dim lenval As Integer

  On Error GoTo donehere
  lenval = Len(Trim(frm_edbias.txt_newval.Text))
  
  If (lenval = 0) Then
    frm_edbias.txt_newval.Text = "-"
    GoTo donehere
  End If
  
  If (lenval = 1) Then
    If (Trim(frm_edbias.txt_newval.Text) = "-") Then
      frm_edbias.txt_newval.Text = ""
      GoTo donehere
    End If
  End If
  
  tempval = frm_edbias.txt_newval.Text
  tempval = -tempval
  frm_edbias.txt_newval.Text = tempval
donehere:
End Sub

Private Sub cmd_newValue_Click()
  Dim tempint As Integer
  Dim numcheck As Boolean

  unity_main.errorstring = "Bias Configuration screen " & cmd_newValue.Caption & " button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  tempint = frm_edbias.ss_biases.DataRowCnt
  
  If (Trim(frm_edbias.txt_newval.Text) = "") Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_edbias", "errMsg1", "You must enter a value for the new bias!"), vbCritical
    Exit Sub
  End If

  numcheck = IsNumeric(frm_edbias.txt_newval.Text)
  
  If (numcheck = False) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_edbias", "errMsg1", "You must enter a value for the new bias!"), vbCritical
    Exit Sub
  End If
  
  If (Trim(frm_edbias.lbl_row.Caption) = "") Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_edbias", "errMsg2", "No property selected; select a property to edit by clicking on the row of the property and then enter the new bias!"), vbCritical
    Exit Sub
  End If

  If (frm_edbias.lbl_row.Caption > tempint) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_edbias", "errMsg2", "No property selected select a property to edit by clicking on the row of the property and then enter the new bias!"), vbCritical
    Exit Sub
  End If
  
  If (frm_edbias.lbl_row.Caption < 1) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_edbias", "errMsg2", "No property selected; select a property to edit by clicking on the row of the property and then enter the new bias!"), vbCritical
    Exit Sub
  End If

  frm_edbias.ss_biases.Row = frm_edbias.lbl_row.Caption
  frm_edbias.ss_biases.Col = 3
  frm_edbias.ss_biases.Row = frm_edbias.lbl_row.Caption
  frm_edbias.ss_biases.Text = Trim(frm_edbias.txt_newval.Text)
End Sub

Private Sub cmd_save_Click()
  
  frm_edbias.savebiases
  unity_main.pw_open = False ' will have to reenter password next time
  Call unity_main.restart_loop(LOG_DBG_LEVEL1, "Bias Configuration screen Save Changes button selected")
  Unload frm_edbias
End Sub

Private Sub cmd_cancel_Click()
  
  unity_main.pw_open = False ' will have to reenter password next time
  Call unity_main.restart_loop(LOG_DBG_LEVEL3, "Bias Configuration screen Cancel button selected")
  Unload frm_edbias
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub

Private Sub ss_biases_Click(ByVal Col As Long, ByVal Row As Long)
  
  frm_edbias.ss_biases.Col = 1
  frm_edbias.lbl_row.Caption = frm_edbias.ss_biases.ActiveRow
  frm_edbias.ss_biases.Row = frm_edbias.lbl_row.Caption
  frm_edbias.lbl_prop.Caption = frm_edbias.ss_biases.Text
  frm_edbias.ss_biases.Col = 3
  frm_edbias.txt_oldval.Text = frm_edbias.ss_biases.Text
  frm_edbias.txt_newval.Text = ""
End Sub

Private Sub ss_biases_KeyUp(KeyCode As Integer, Shift As Integer)

  frm_edbias.ss_biases.Col = 1
  frm_edbias.lbl_prop.Caption = frm_edbias.ss_biases.Text
  frm_edbias.ss_biases.Col = 3
  frm_edbias.txt_oldval.Text = frm_edbias.ss_biases.Text
End Sub








