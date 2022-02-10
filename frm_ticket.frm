VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_ticket 
   Caption         =   "Ticket Printer Configuration"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
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
   Icon            =   "frm_ticket.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniFrameXP frame_preReport 
      Height          =   1215
      Left            =   240
      Top             =   1800
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_ticket.frx":0442
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_ticket.frx":0488
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_ticket.frx":04A8
      Begin HexUniControls.ctlUniLabel lbl_preNumLFs 
         Height          =   375
         Left            =   5040
         Top             =   480
         Width           =   2475
         _ExtentX        =   4366
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
         Caption         =   "frm_ticket.frx":04C4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_ticket.frx":0504
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":0524
      End
      Begin HexUniControls.ctlNumIncXP numInc_preLFs 
         Height          =   615
         Left            =   3720
         TabIndex        =   6
         Top             =   400
         Width           =   1215
         _ExtentX        =   2143
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
         Text            =   "0"
         Min             =   0
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
         MouseIcon       =   "frm_ticket.frx":0540
      End
      Begin HexUniControls.ctlUniCheckXP chk_preDelimiter 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   3000
         _ExtentX        =   5292
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
         Caption         =   "frm_ticket.frx":055C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":058E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":05AE
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_preFF 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   3000
         _ExtentX        =   5292
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
         Caption         =   "frm_ticket.frx":05CA
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":05FA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":061A
         ShowFocus       =   -1  'True
      End
   End
   Begin HexUniControls.ctlUniFrameXP frame_prtrSettings 
      Height          =   1575
      Left            =   240
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
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
      Caption         =   "frm_ticket.frx":0636
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_ticket.frx":0676
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_ticket.frx":0696
      Begin HexUniControls.ctlUniCheckXP chk_fontBold 
         Height          =   375
         Left            =   5040
         TabIndex        =   2
         Top             =   360
         Width           =   2520
         _ExtentX        =   4445
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
         Caption         =   "frm_ticket.frx":06B2
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":06E4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":0704
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlNumIncXP numInc_fontSize 
         Height          =   615
         Left            =   5040
         TabIndex        =   3
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
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
         Text            =   "12"
         Min             =   7
         Max             =   12
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
         MouseIcon       =   "frm_ticket.frx":0720
      End
      Begin HexUniControls.ctlUniLabel lbl_fontSize 
         Height          =   615
         Left            =   6360
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
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
         Caption         =   "frm_ticket.frx":073C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_ticket.frx":076E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":078E
      End
      Begin HexUniControls.ctlUniLabel lbl_prtSelect 
         Height          =   375
         Left            =   240
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
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
         Caption         =   "frm_ticket.frx":07AA
         BackColor       =   -2147483633
         ForeColor       =   0
         Alignment       =   2
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_ticket.frx":07EC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":080C
      End
Begin HexUniControls.ctlUniComboBoxXP combo_printer
         Height          =   405
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   714
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
         Tip             =   "frm_ticket.frx":0828
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
         MouseIcon       =   "frm_ticket.frx":0848
         DropDownOnTextClick=   -1  'True
         ManualStart     =   0   'False
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3840
      Top             =   8280
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9120
      FormDesignWidth =   8205
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   645
      Left            =   1080
      TabIndex        =   17
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
      Caption         =   "frm_ticket.frx":0864
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_ticket.frx":089C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_ticket.frx":08BC
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   645
      Left            =   5040
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
      Caption         =   "frm_ticket.frx":08D8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_ticket.frx":0904
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_ticket.frx":0924
   End
   Begin HexUniControls.ctlUniFrameXP frame_postReport 
      Height          =   1215
      Left            =   240
      Top             =   6720
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_ticket.frx":0940
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_ticket.frx":0988
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_ticket.frx":09A8
      Begin HexUniControls.ctlUniLabel lbl_postNumLFs 
         Height          =   375
         Left            =   5040
         Top             =   480
         Width           =   2475
         _ExtentX        =   4366
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
         Caption         =   "frm_ticket.frx":09C4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_ticket.frx":0A04
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":0A24
      End
      Begin HexUniControls.ctlNumIncXP numInc_postLFs 
         Height          =   615
         Left            =   3720
         TabIndex        =   16
         Top             =   400
         Width           =   1215
         _ExtentX        =   2143
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
         Text            =   "0"
         Min             =   0
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
         MouseIcon       =   "frm_ticket.frx":0A40
      End
      Begin HexUniControls.ctlUniCheckXP chk_postDelimiter 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   3000
         _ExtentX        =   5292
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
         Caption         =   "frm_ticket.frx":0A5C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":0A8E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":0AAE
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_postFF 
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   3000
         _ExtentX        =   5292
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
         Caption         =   "frm_ticket.frx":0ACA
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":0AFA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":0B1A
         ShowFocus       =   -1  'True
      End
   End
   Begin HexUniControls.ctlUniFrameXP frame_header 
      Height          =   1455
      Left            =   240
      Top             =   3120
      Width           =   7695
      _ExtentX        =   13573
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
      Caption         =   "frm_ticket.frx":0B36
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_ticket.frx":0B74
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_ticket.frx":0B94
      Begin HexUniControls.ctlUniTextBoxXP txt_header2 
         Height          =   400
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   714
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frm_ticket.frx":0BB0
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
         Tip             =   "frm_ticket.frx":0BE2
         NoHideSel       =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":0C02
      End
      Begin HexUniControls.ctlUniTextBoxXP txt_header1 
         Height          =   400
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   714
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frm_ticket.frx":0C1E
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
         Tip             =   "frm_ticket.frx":0C56
         NoHideSel       =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":0C76
      End
      Begin HexUniControls.ctlUniLabel lbl_header2 
         Height          =   375
         Left            =   4680
         Top             =   960
         Width           =   1515
         _ExtentX        =   2672
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
         Caption         =   "frm_ticket.frx":0C92
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_ticket.frx":0CBE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":0CDE
      End
      Begin HexUniControls.ctlUniLabel lbl_header1 
         Height          =   375
         Left            =   4680
         Top             =   480
         Width           =   1515
         _ExtentX        =   2672
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
         Caption         =   "frm_ticket.frx":0CFA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   1
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_ticket.frx":0D26
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":0D46
      End
   End
   Begin HexUniControls.ctlUniFrameXP frame_data 
      Height          =   1935
      Left            =   240
      Top             =   4680
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3413
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_ticket.frx":0D62
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_ticket.frx":0DC6
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_ticket.frx":0DE6
      Begin HexUniControls.ctlUniCheckXP chk_input 
         Height          =   375
         Index           =   8
         Left            =   5880
         TabIndex        =   25
         Top             =   1440
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "frm_ticket.frx":0E02
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":0E30
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":0E50
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input 
         Height          =   375
         Index           =   7
         Left            =   5880
         TabIndex        =   24
         Top             =   1080
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "frm_ticket.frx":0E6C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":0E9A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":0EBA
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input 
         Height          =   375
         Index           =   6
         Left            =   5880
         TabIndex        =   23
         Top             =   720
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "frm_ticket.frx":0ED6
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":0F04
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":0F24
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input 
         Height          =   375
         Index           =   5
         Left            =   5880
         TabIndex        =   22
         Top             =   360
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "frm_ticket.frx":0F40
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":0F6E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":0F8E
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input 
         Height          =   375
         Index           =   4
         Left            =   4080
         TabIndex        =   21
         Top             =   1440
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "frm_ticket.frx":0FAA
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":0FD8
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":0FF8
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input 
         Height          =   375
         Index           =   3
         Left            =   4080
         TabIndex        =   20
         Top             =   1080
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "frm_ticket.frx":1014
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":1042
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":1062
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input 
         Height          =   375
         Index           =   2
         Left            =   4080
         TabIndex        =   19
         Top             =   720
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "frm_ticket.frx":107E
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":10AC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":10CC
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input 
         Height          =   375
         Index           =   1
         Left            =   4080
         TabIndex        =   18
         Top             =   360
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "frm_ticket.frx":10E8
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":1116
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":1136
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_comment 
         Height          =   450
         Left            =   1920
         TabIndex        =   13
         Top             =   1360
         Width           =   2115
         _ExtentX        =   3731
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
         Caption         =   "frm_ticket.frx":1152
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":118E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":11AE
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_sampID 
         Height          =   450
         Left            =   1920
         TabIndex        =   12
         Top             =   860
         Width           =   2085
         _ExtentX        =   3678
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
         Caption         =   "frm_ticket.frx":11CA
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":11FC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":121C
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_product 
         Height          =   450
         Left            =   1920
         TabIndex        =   11
         Top             =   360
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "frm_ticket.frx":1238
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":1266
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":1286
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_time 
         Height          =   450
         Left            =   240
         TabIndex        =   10
         Top             =   860
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "frm_ticket.frx":12A2
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":12CA
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":12EA
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_date 
         Height          =   450
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1600
         _ExtentX        =   2831
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
         Caption         =   "frm_ticket.frx":1306
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":132E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":134E
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_serialNum 
         Height          =   450
         Left            =   240
         TabIndex        =   26
         Top             =   1360
         Width           =   1605
         _ExtentX        =   2831
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
         Caption         =   "frm_ticket.frx":136A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_ticket.frx":139E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_ticket.frx":13BE
         ShowFocus       =   -1  'True
      End
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   4440
      Top             =   8160
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
      Left            =   3360
      Top             =   8280
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_ticket.frx":13DA
   End
End
Attribute VB_Name = "frm_ticket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Configuration file parameters
Private m_tktComment As Integer
Private m_tktDate As Integer
Private m_tktHeader1 As String
Private m_tktHeader2 As String
Private m_tktPostDelimiter As Integer
Private m_tktPostFF As Integer
Private m_tktPostLFs As Integer
Private m_tktPreDelimiter As Integer
Private m_tktPreFF As Integer
Private m_tktPreLFs As Integer
Private m_tktPrinterBold As Integer
Private m_tktPrinterFont As Integer
Private m_tktPrinterName As String
Private m_tktProduct As Integer
Private m_tktSampId As Integer
Private m_tktSerialNum As Integer
Private m_tktTime As Integer
Private m_tktUserInputs(1 To MAX_MAN_INPUTS) As Boolean

' Misc variables
Public m_badTktIniVal As Boolean

Sub saveticketfile(setupCfg As Boolean)
  Dim ii As Integer
  Dim fileName As String
  Dim uniFile As New clsUniFile
  Dim errMsg As String
  Dim uniMsg As String

  m_tktPrinterName = frm_ticket.combo_printer.Text
  m_tktPrinterFont = frm_ticket.numInc_fontSize.Text
  m_tktPrinterBold = frm_ticket.chk_fontBold.value
  m_tktPreFF = frm_ticket.chk_preFF.value
  m_tktPreLFs = frm_ticket.numInc_preLFs.Text
  m_tktPreDelimiter = frm_ticket.chk_preDelimiter.value
  m_tktHeader1 = Trim(frm_ticket.txt_header1.Text)
  m_tktHeader2 = Trim(frm_ticket.txt_header2.Text)
  m_tktDate = frm_ticket.chk_date.value
  m_tktTime = frm_ticket.chk_time.value
  m_tktSerialNum = frm_ticket.chk_serialNum.value
  m_tktProduct = frm_ticket.chk_product.value
  m_tktSampId = frm_ticket.chk_sampID.value
  m_tktComment = frm_ticket.chk_comment.value
  
  For ii = 1 To MAX_MAN_INPUTS
    m_tktUserInputs(ii) = frm_ticket.chk_input(ii).value
  Next ii

  m_tktPostDelimiter = frm_ticket.chk_postDelimiter.value
  m_tktPostLFs = frm_ticket.numInc_postLFs.Text
  m_tktPostFF = frm_ticket.chk_postFF.value
  
  fileName = (CFG_DIR & PRINT_TICKET_CFG_FILE)
  
  If (uniFile.OpenFileWrite(fileName) = True) Then
    On Error GoTo FILE_ERROR
    uniFile.WriteBOM fe_UTF16LE
    uniFile.WriteUnicodeLine ("Version=" & INFOSTAR_VER)
    uniFile.WriteUnicodeLine ("Printer=" & m_tktPrinterName)
    uniFile.WriteUnicodeLine ("Font_Size=" & m_tktPrinterFont)
    uniFile.WriteUnicodeLine ("Font_Bold=" & m_tktPrinterBold)
    uniFile.WriteUnicodeLine ("PreReport_FF=" & m_tktPreFF)
    uniFile.WriteUnicodeLine ("PreReport_LFs=" & m_tktPreLFs)
    uniFile.WriteUnicodeLine ("PreReport_Delimiter=" & m_tktPreDelimiter)
    uniFile.WriteUnicodeLine ("Header_Line1=" & m_tktHeader1)
    uniFile.WriteUnicodeLine ("Header_Line2=" & m_tktHeader2)
    uniFile.WriteUnicodeLine ("Include_Date=" & m_tktDate)
    uniFile.WriteUnicodeLine ("Include_Time=" & m_tktTime)
    uniFile.WriteUnicodeLine ("Include_SerialNum=" & m_tktSerialNum)
    uniFile.WriteUnicodeLine ("Include_Product=" & m_tktProduct)
    uniFile.WriteUnicodeLine ("Include_Sample_ID=" & m_tktSampId)
    uniFile.WriteUnicodeLine ("Include_Comment=" & m_tktComment)
  
    For ii = 1 To MAX_MAN_INPUTS
      uniFile.WriteUnicodeLine ("Input" & ii & "=" & m_tktUserInputs(ii))
    Next ii
  
    uniFile.WriteUnicodeLine ("PostReport_Delimiter=" & m_tktPostDelimiter)
    uniFile.WriteUnicodeLine ("PostReport_LFs=" & m_tktPostLFs)
    uniFile.WriteUnicodeLine ("PostReport_FF=" & m_tktPostFF)
    uniFile.Flush

    If (setupCfg = True) Then
      ' Copy saved ticket config into system operational variables
      unity_main.m_tktComment = m_tktComment
      unity_main.m_tktDate = m_tktDate
      unity_main.m_tktHeader1 = m_tktHeader1
      unity_main.m_tktHeader2 = m_tktHeader2
      unity_main.m_tktPostDelimiter = m_tktPostDelimiter
      unity_main.m_tktPostLFs = m_tktPostLFs
      unity_main.m_tktPostFF = m_tktPostFF
      unity_main.m_tktPreDelimiter = m_tktPreDelimiter
      unity_main.m_tktPreLFs = m_tktPreLFs
      unity_main.m_tktPreFF = m_tktPreFF
      unity_main.m_tktPrinterBold = m_tktPrinterBold
      unity_main.m_tktPrinterFont = m_tktPrinterFont
      unity_main.m_tktPrinterName = m_tktPrinterName
      unity_main.m_tktProduct = m_tktProduct
      unity_main.m_tktSampId = m_tktSampId
      unity_main.m_tktSerialNum = m_tktSerialNum
      unity_main.m_tktTime = m_tktTime
            
      For ii = 1 To MAX_MAN_INPUTS
        TktUserInputs(ii) = m_tktUserInputs(ii)
      Next ii
    End If
      
    unity_main.errorstring = ("User saved new settings for configuration file: " & (CFG_DIR & PRINT_TICKET_CFG_FILE))
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

Sub loadticketfile(mustBeCfg As Boolean)
  Dim ii As Integer
  Dim fileName As String
  Dim p As Printer
  Dim uniMsg As String

  frm_ticket.combo_printer.clear
      
  For Each p In Printers
    frm_ticket.combo_printer.AddItem p.DeviceName
  Next
  
  m_badTktIniVal = False
  fileName = (CFG_DIR & PRINT_TICKET_CFG_FILE)
  
  If (CFile.st_FileExist(fileName) = False) Then
    If (mustBeCfg = True) Then
      uniMsg = MLSupport.GGS_Params("fileErrMsg6", "%1 file not found", fileName)
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_ticket.errMsg1", "%1. Loaded product configured for ticket printer; please configure the Ticket Printer settings and save them to file.", uniMsg), vbExclamation
    Else
      ' Setup screen's default values
      frm_ticket.chk_fontBold.value = 1
      frm_ticket.chk_preDelimiter.value = 1
      frm_ticket.chk_postDelimiter.value = 1
    End If
  Else
    ' Setup defaults
    unity_main.m_fileVersion = INFOSTAR_VER
    m_tktPrinterName = ""
    m_tktPrinterFont = 12
    m_tktPrinterBold = 1
    m_tktPreFF = 0
    m_tktPreLFs = frm_ticket.numInc_preLFs.Min
    m_tktPreDelimiter = 1
    m_tktHeader1 = ""
    m_tktHeader2 = ""
    m_tktDate = 0
    m_tktTime = 0
    m_tktSerialNum = 0
    m_tktProduct = 0
    m_tktSampId = 0
    m_tktComment = 0
    
    For ii = 1 To MAX_MAN_INPUTS
      m_tktUserInputs(ii) = False
    Next ii
    
    m_tktPostDelimiter = 1
    m_tktPostLFs = frm_ticket.numInc_postLFs.Min
    m_tktPostFF = 0
  
    If (load_ticket_file_vals(fileName) = True) Then
      ' Check for invalid file version
      If (unity_main.m_fileVersion <> INFOSTAR_VER) Then
        unity_main.errorstring = (fileName & " had incompatible value. Version was " & unity_main.m_fileVersion & "; updated to " & INFOSTAR_VER)
        unity_main.write_error
        unity_main.m_fileVersion = INFOSTAR_VER
        m_badTktIniVal = True
      End If
    
      ' Check for invalid printer selection
      For Each p In Printers
        If (p.DeviceName = m_tktPrinterName) Then
          GoTo FOUND_PRTR
        End If
      Next
      
      If (m_tktPrinterName <> "None") Then
        unity_main.errorstring = (fileName & " had incompatible value. Printer was " & m_tktPrinterName & "; updated to Default printer: " & frmReport.m_dfltPrinterName)
        unity_main.write_error
        m_tktPrinterName = frmReport.m_dfltPrinterName
        m_badTktIniVal = True
      End If
    
FOUND_PRTR:
      frm_ticket.combo_printer.Text = m_tktPrinterName
    
      ' Check for invalid font size
      If ((m_tktPrinterFont < frm_ticket.numInc_fontSize.Min) Or (m_tktPrinterFont > frm_ticket.numInc_fontSize.Max)) Then
        unity_main.errorstring = (fileName & " had incompatible value. Font_Size was " & m_tktPrinterFont & "; updated to " & frm_ticket.numInc_fontSize.Max)
        unity_main.write_error
        m_tktPrinterFont = frm_ticket.numInc_fontSize.Max
        m_badTktIniVal = True
      End If
    
      frm_ticket.numInc_fontSize.Text = m_tktPrinterFont
      
      If (m_tktPrinterBold <> 0) Then
        frm_ticket.chk_fontBold.value = 1
      Else
        frm_ticket.chk_fontBold.value = 0
      End If
    
      If (m_tktPreFF <> 0) Then
        frm_ticket.chk_preFF.value = 1
      Else
        frm_ticket.chk_preFF.value = 0
      End If
    
      ' Check for invalid # of pre-report linefeeds
      If ((m_tktPreLFs < frm_ticket.numInc_preLFs.Min) Or (m_tktPreLFs > frm_ticket.numInc_preLFs.Max)) Then
        unity_main.errorstring = (fileName & " had incompatible value. PreReport_LFs was " & m_tktPreLFs & "; updated to " & frm_ticket.numInc_preLFs.Min)
        unity_main.write_error
        m_tktPreLFs = frm_ticket.numInc_preLFs.Min
        m_badTktIniVal = True
      End If
    
      frm_ticket.numInc_preLFs.Text = m_tktPreLFs
    
      If (m_tktPreDelimiter <> 0) Then
        frm_ticket.chk_preDelimiter.value = 1
      Else
        frm_ticket.chk_preDelimiter.value = 0
      End If
    
      frm_ticket.txt_header1.Text = m_tktHeader1
      frm_ticket.txt_header2.Text = m_tktHeader2
      
      If (m_tktDate <> 0) Then
        frm_ticket.chk_date.value = 1
      Else
        frm_ticket.chk_date.value = 0
      End If

      If (m_tktTime <> 0) Then
        frm_ticket.chk_time.value = 1
      Else
        frm_ticket.chk_time.value = 0
      End If

      If (m_tktSerialNum <> 0) Then
        frm_ticket.chk_serialNum.value = 1
      Else
        frm_ticket.chk_serialNum.value = 0
      End If

      If (m_tktProduct <> 0) Then
        frm_ticket.chk_product.value = 1
      Else
        frm_ticket.chk_product.value = 0
      End If

      If (m_tktSampId <> 0) Then
        frm_ticket.chk_sampID.value = 1
      Else
        frm_ticket.chk_sampID.value = 0
      End If

      If (m_tktComment <> 0) Then
        frm_ticket.chk_comment.value = 1
      Else
        frm_ticket.chk_comment.value = 0
      End If

      For ii = 1 To MAX_MAN_INPUTS
        If (m_tktUserInputs(ii) = True) Then
          frm_ticket.chk_input(ii).value = 1
        Else
          frm_ticket.chk_input(ii).value = 0
        End If
      Next ii

      If (m_tktPostDelimiter <> 0) Then
        frm_ticket.chk_postDelimiter.value = 1
      Else
        frm_ticket.chk_postDelimiter.value = 0
      End If

      ' Check for invalid # of post-report linefeeds
      If ((m_tktPostLFs < frm_ticket.numInc_postLFs.Min) Or (m_tktPostLFs > frm_ticket.numInc_postLFs.Max)) Then
        unity_main.errorstring = (fileName & " had incompatible value. PostReport_LFs was " & m_tktPostLFs & "; updated to " & frm_ticket.numInc_postLFs.Min)
        unity_main.write_error
        m_tktPostLFs = frm_ticket.numInc_postLFs.Min
        m_badTktIniVal = True
      End If
    
      frm_ticket.numInc_postLFs.Text = m_tktPostLFs
    
      If (m_tktPostFF <> 0) Then
        frm_ticket.chk_postFF.value = 1
      Else
        frm_ticket.chk_postFF.value = 0
      End If
    
      ' Check if ini file had bad value
      If (m_badTktIniVal = True) Then
        unity_main.errorstring = (fileName & " had incompatible value(s). Updated file with proper/default value(s).")
        unity_main.write_error
        uniMsg = MLSupport.GGS_Params("fileErrMsg7", "%1 had incompatible value(s). Updated file with proper/default value(s)", fileName)
        CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg3", "%1. Refer to system log for more details", uniMsg), vbOKOnly
        Call saveticketfile(False)
      End If
      
      If (mustBeCfg = True) Then
        ' Copy loaded ticket config into system operational variables
        unity_main.m_tktComment = m_tktComment
        unity_main.m_tktDate = m_tktDate
        unity_main.m_tktHeader1 = m_tktHeader1
        unity_main.m_tktHeader2 = m_tktHeader2
        unity_main.m_tktPostDelimiter = m_tktPostDelimiter
        unity_main.m_tktPostLFs = m_tktPostLFs
        unity_main.m_tktPostFF = m_tktPostFF
        unity_main.m_tktPreDelimiter = m_tktPreDelimiter
        unity_main.m_tktPreLFs = m_tktPreLFs
        unity_main.m_tktPreFF = m_tktPreFF
        unity_main.m_tktPrinterBold = m_tktPrinterBold
        unity_main.m_tktPrinterFont = m_tktPrinterFont
        unity_main.m_tktPrinterName = m_tktPrinterName
        unity_main.m_tktProduct = m_tktProduct
        unity_main.m_tktSampId = m_tktSampId
        unity_main.m_tktSerialNum = m_tktSerialNum
        unity_main.m_tktTime = m_tktTime
    
        For ii = 1 To MAX_MAN_INPUTS
          TktUserInputs(ii) = m_tktUserInputs(ii)
        Next ii
      End If
    End If
  End If
End Sub

Function load_ticket_file_vals(srcFile As String) As Boolean
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
          Case "printer"
            m_tktPrinterName = Trim(varVal)
          Case "font_size"
            m_tktPrinterFont = CInt(varVal)
          Case "font_bold"
            m_tktPrinterBold = CInt(varVal)
          Case "prereport_ff"
            m_tktPreFF = CInt(varVal)
          Case "prereport_lfs"
            m_tktPreLFs = CInt(varVal)
          Case "prereport_delimiter"
            m_tktPreDelimiter = CInt(varVal)
          Case "header_line1"
            m_tktHeader1 = Trim(varVal)
          Case "header_line2"
            m_tktHeader2 = Trim(varVal)
          Case "include_date"
            m_tktDate = CInt(varVal)
          Case "include_time"
            m_tktTime = CInt(varVal)
          Case "include_serialnum"
            m_tktSerialNum = CInt(varVal)
          Case "include_product"
            m_tktProduct = CInt(varVal)
          Case "include_sample_id"
            m_tktSampId = CInt(varVal)
          Case "include_comment"
            m_tktComment = CInt(varVal)
          Case "input1"
            m_tktUserInputs(1) = CBool(varVal)
          Case "input2"
            m_tktUserInputs(2) = CBool(varVal)
          Case "input3"
            m_tktUserInputs(3) = CBool(varVal)
          Case "input4"
            m_tktUserInputs(4) = CBool(varVal)
          Case "input5"
            m_tktUserInputs(5) = CBool(varVal)
          Case "input6"
            m_tktUserInputs(6) = CBool(varVal)
          Case "input7"
            m_tktUserInputs(7) = CBool(varVal)
          Case "input8"
            m_tktUserInputs(8) = CBool(varVal)
          Case "postreport_delimiter"
            m_tktPostDelimiter = CInt(varVal)
          Case "postreport_lfs"
            m_tktPostLFs = CInt(varVal)
          Case "postreport_ff"
            m_tktPostFF = CInt(varVal)
        End Select
      End If
    Wend
  
    load_ticket_file_vals = True
  Else
FILE_ERROR:
    errMsg = (srcFile & " file read error. " & Error$)
    unity_main.errorstring = errMsg
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg4", "%1 file read error. %2", srcFile, Error$)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    load_ticket_file_vals = False
  End If
  
  uniFile.CloseFile
  Exit Function
  
BAD_INI_VALUE:
  unity_main.errorstring = (srcFile & " line " & CStr(lineCnt) & " had incompatible value. " & Trim(tmpStrg) & " - " & varVal & "; will use default value")
  unity_main.write_error
  m_badTktIniVal = True
  Resume Next
End Function

Sub writeticket()
  Dim ii As Integer
  Dim tempint As Integer
  Dim tempstring1 As String
  Dim tempstring2 As String
  Dim int2 As Integer
  Dim errMsg As String
  Dim p As Printer
  Dim jj As Integer
  Dim uniMsg As String
  Dim fontTwips As Long
  Dim msg As String
  Dim rc As Long
  Dim lineCnt As Integer
  
  uniMsg = MLSupport.GSS("frm_ticket", "statMsg1", "Writing sample results to ticket printer")
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Writing sample results to ticket printer", uniMsg)
  
  On Error GoTo PRINT_ERROR
  
  For Each p In Printers
    If (p.DeviceName = unity_main.m_tktPrinterName) Then
      ' Set printer as system default.
      Set Printer = p
      GoTo FOUND_PRTR
    End If
  Next
  
  If (unity_main.m_tktPrinterName = "None") Then
    errMsg = ("Ticket printer not configured or installed")
    uniMsg = MLSupport.GGS_Params("prtrErrMsg1", "%1 printer not configured or installed", MLSupport.GSS("frm_ticket", "statMsg2", "Ticket"))
  Else
    errMsg = (unity_main.m_tktPrinterName & " printer not configured or installed")
    uniMsg = MLSupport.GGS_Params("prtrErrMsg1", "%1 printer not configured or installed", unity_main.m_tktPrinterName)
  End If
  
  unity_main.errorstring = errMsg
  unity_main.write_error
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  Exit Sub
 
FOUND_PRTR:
  Printer.FontName = "Arial Unicode MS"
  Printer.FontBold = CBool(unity_main.m_tktPrinterBold)
  Printer.FontSize = unity_main.m_tktPrinterFont
  
  ' Initialize to get the hdc
  Printer.Print " "
  
  ' Convert font size (points) to twips (1/20 pts per twips)
  fontTwips = (Printer.FontSize * 1.3333) / (0.05 * Printer.TwipsPerPixelY)
  
  ' Check if to print pre-report formfeed
  If (unity_main.m_tktPreFF <> 0) Then
    Printer.NewPage
  End If
  
  ' Print any pre-report linefeeds
  For jj = 1 To unity_main.m_tktPreLFs
    msg = ""
    rc = TextOutW(Printer.hdc, 0, fontTwips * lineCnt, StrConv(msg, vbUnicode), Len(msg))
    lineCnt = lineCnt + 1
  Next jj
  
  ' Check if to print pre-report delimiter
  If (unity_main.m_tktPreDelimiter <> 0) Then
    msg = "**************************"
    rc = TextOutW(Printer.hdc, 0, fontTwips * lineCnt, StrConv(msg, vbUnicode), Len(msg))
    lineCnt = lineCnt + 1
  End If
  
  msg = MLSupport.GSS("frm_ticket", "printMsg1", "Unity InfoStar")
  rc = TextOutW(Printer.hdc, 0, fontTwips * lineCnt, StrConv(msg, vbUnicode), Len(msg))
  lineCnt = lineCnt + 1
  
  msg = MLSupport.GSS("frm_ticket", "printMsg2", "Sample Ticket Report")
  rc = TextOutW(Printer.hdc, 0, fontTwips * lineCnt, StrConv(msg, vbUnicode), Len(msg))
  lineCnt = lineCnt + 1
  
  ' Check if to print header line 1
  If (unity_main.m_tktHeader1 <> "") Then
    msg = unity_main.m_tktHeader1
    rc = TextOutW(Printer.hdc, 0, fontTwips * lineCnt, StrConv(msg, vbUnicode), Len(msg))
    lineCnt = lineCnt + 1
  End If
  
  ' Check if to print header line 2
  If (unity_main.m_tktHeader2 <> "") Then
    msg = unity_main.m_tktHeader2
    rc = TextOutW(Printer.hdc, 0, fontTwips * lineCnt, StrConv(msg, vbUnicode), Len(msg))
    lineCnt = lineCnt + 1
  End If
  
  ' Check if to print sample date
  If (unity_main.m_tktDate <> 0) Then
    msg = (MLSupport.GSS("Headers", "date", "Date") & ": " & unity_main.lbl_date.Caption)
    rc = TextOutW(Printer.hdc, 0, fontTwips * lineCnt, StrConv(msg, vbUnicode), Len(msg))
    lineCnt = lineCnt + 1
  End If
  
  ' Check if to print sample time
  If (unity_main.m_tktTime <> 0) Then
    msg = (MLSupport.GSS("Headers", "time", "Time") & ": " & unity_main.lbl_time.Caption)
    rc = TextOutW(Printer.hdc, 0, fontTwips * lineCnt, StrConv(msg, vbUnicode), Len(msg))
    lineCnt = lineCnt + 1
  End If
  
  ' Check if to print system serial number
  If (unity_main.m_tktSerialNum <> 0) Then
    msg = (MLSupport.GSS("Headers", "serNum", "Serial No.") & ": " & unity_main.m_sysSerialNum)
    rc = TextOutW(Printer.hdc, 0, fontTwips * lineCnt, StrConv(msg, vbUnicode), Len(msg))
    lineCnt = lineCnt + 1
  End If
  
  ' Check if to print product name
  If (unity_main.m_tktProduct <> 0) Then
    msg = (MLSupport.GSS("Headers", "product", "Product") & ": " & unity_main.lblProd1.Caption)
    rc = TextOutW(Printer.hdc, 0, fontTwips * lineCnt, StrConv(msg, vbUnicode), Len(msg))
    lineCnt = lineCnt + 1
  End If
  
  ' Check if to print sample ID
  If (unity_main.m_tktSampId <> 0) Then
    msg = (MLSupport.GSS("Headers", "sampleID", "Sample ID") & ": " & unity_main.txtsamplename.Text)
    rc = TextOutW(Printer.hdc, 0, fontTwips * lineCnt, StrConv(msg, vbUnicode), Len(msg))
    lineCnt = lineCnt + 1
  End If
  
  ' Check if to print sample comment
  If (unity_main.m_tktComment <> 0) Then
    msg = (MLSupport.GSS("Headers", "comment", "Comment") & ": " & unity_main.txtsampcomment.Text)
    rc = TextOutW(Printer.hdc, 0, fontTwips * lineCnt, StrConv(msg, vbUnicode), Len(msg))
    lineCnt = lineCnt + 1
  End If

  ' Check if to print user input 1 - 8
  For ii = 1 To MAX_MAN_INPUTS
    If (TktUserInputs(ii) = True) Then
      tempstring1 = Trim(frm_scanname.lbl(ii).Caption)
    
      If Trim(tempstring1) = "" Then
        tempstring1 = MLSupport.GSS("Headers", "input", "Input") & " " & CStr(ii)
      End If
            
      ' Setup for input enable field
      frm_buttoncfg.ss_buttonconfig.Col = ii
      frm_buttoncfg.ss_buttonconfig.Row = 1
  
      ' Check if input enabled
      If (frm_buttoncfg.ss_buttonconfig.value = 1) Then
        ' Setup for text entry/list box selection field
        frm_buttoncfg.ss_buttonconfig.Col = ii
        frm_buttoncfg.ss_buttonconfig.Row = 2
  
        ' Check if using text entry
        If (frm_buttoncfg.ss_buttonconfig.value = 0) Then
          tempstring2 = Trim(frm_scanname.txtbx(ii).Text)
        Else    ' Using list
          tempstring2 = Trim(frm_scanname.combo(ii).Text)
        End If
      Else
        tempstring2 = MLSupport.GSS("Headers", "na", "NA")
      End If
      
      msg = (tempstring1 & ": " & tempstring2)
      rc = TextOutW(Printer.hdc, 0, fontTwips * lineCnt, StrConv(msg, vbUnicode), Len(msg))
      lineCnt = lineCnt + 1
    End If
  Next ii

  ' Print each property name and value
  tempint = frmedmod.numprops.Text

  For jj = 1 To tempint
    unity_main.fpspread_pred.Row = jj
    unity_main.fpspread_pred.Col = 1
    tempstring2 = Trim(unity_main.fpspread_pred.Text)
    tempstring2 = tempstring2 & " = "
    unity_main.fpspread_pred.Col = 2
    tempstring2 = tempstring2 & Trim(unity_main.fpspread_pred.Text)
    msg = tempstring2
    rc = TextOutW(Printer.hdc, 0, fontTwips * lineCnt, StrConv(msg, vbUnicode), Len(msg))
    lineCnt = lineCnt + 1
  Next jj
  
  ' Check if to print post-report delimiter
  If (unity_main.m_tktPostDelimiter <> 0) Then
    msg = "**************************"
    rc = TextOutW(Printer.hdc, 0, fontTwips * lineCnt, StrConv(msg, vbUnicode), Len(msg))
    lineCnt = lineCnt + 1
  End If
  
  ' Print any post-report linefeeds
  For jj = 1 To unity_main.m_tktPreLFs
    msg = ""
    rc = TextOutW(Printer.hdc, 0, fontTwips * lineCnt, StrConv(msg, vbUnicode), Len(msg))
    lineCnt = lineCnt + 1
  Next jj

  ' Check if to print post-report formfeed
  If (unity_main.m_tktPreFF <> 0) Then
    Printer.NewPage
  End If
  
  Printer.EndDoc
  Exit Sub
  
PRINT_ERROR:
  errMsg = ("Ticket printer print error. " & Error$)
  unity_main.errorstring = errMsg
  unity_main.write_error
  uniMsg = MLSupport.GGS_Params("prtrErrMsg2", "Ticket printer print error. %1", Error$)
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub

Private Sub cmd_cancel_Click()
  
  unity_main.errorstring = "Ticket Printer Configuration screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Unload frm_ticket
End Sub

Private Sub cmd_save_Click()
  
  unity_main.errorstring = "Ticket Printer Configuration screen Save Changes button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Call saveticketfile(True)
  Unload frm_ticket
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub

Private Sub numInc_fontSize_DblCLick()

  unity_main.formfrom = 14
  unity_main.varfrom = 3
  frm_numpad.lbl_num.Caption = lbl_fontSize.Caption
  frm_numpad.txt_num.Text = numInc_fontSize.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_postLFs_DblCLick()

  unity_main.formfrom = 14
  unity_main.varfrom = 5
  frm_numpad.lbl_num.Caption = lbl_postNumLFs.Caption
  frm_numpad.txt_num.Text = numInc_postLFs.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_preLFs_DblCLick()

  unity_main.formfrom = 14
  unity_main.varfrom = 4
  frm_numpad.lbl_num.Caption = lbl_preNumLFs.Caption
  frm_numpad.txt_num.Text = numInc_preLFs.Text
  frm_numpad.Show 1
End Sub

Private Sub txt_header1_DblCLick(Button As Integer)

  unity_main.formfrom = 14
  unity_main.varfrom = 1
  frm_kybd.lbl_kybd.Caption = lbl_header1.Caption
  frm_kybd.txt_kybd.Text = txt_header1.Text
  frm_kybd.Show 1
End Sub

Private Sub txt_header2_DblCLick(Button As Integer)

  unity_main.formfrom = 14
  unity_main.varfrom = 2
  frm_kybd.lbl_kybd.Caption = lbl_header2.Caption
  frm_kybd.txt_kybd.Text = txt_header2.Text
  frm_kybd.Show 1
End Sub








