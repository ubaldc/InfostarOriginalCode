VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Object = "{488D84E9-A9A9-4C74-949B-F7752803ADEE}#99.2#0"; "EMDVirtualKeyBoard.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frmedmod 
   Caption         =   "Product Properties Management"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11655
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
   Icon            =   "frmedmod.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   11655
   WindowState     =   2  'Maximized
   Begin HexUniControls.ctlUniListBoxXP lst_prdConstituentNames 
      Height          =   255
      Left            =   10560
      TabIndex        =   13
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TrapTab         =   0   'False
      Tip             =   "frmedmod.frx":0442
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frmedmod.frx":0462
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   4440
      TabIndex        =   8
      Top             =   7080
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
      Caption         =   "frmedmod.frx":047E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmedmod.frx":04AA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmedmod.frx":04CA
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   650
      Left            =   2280
      TabIndex        =   9
      Top             =   7080
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
      Caption         =   "frmedmod.frx":04E6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmedmod.frx":051E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmedmod.frx":053E
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_add 
      Height          =   650
      Left            =   120
      TabIndex        =   10
      Top             =   7080
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
      Caption         =   "frmedmod.frx":055A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmedmod.frx":0592
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmedmod.frx":05B2
   End
   Begin FPUSpreadADO.fpSpread grid_models 
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   11355
      _Version        =   458752
      _ExtentX        =   20029
      _ExtentY        =   5318
      _StockProps     =   64
      ColHeaderDisplay=   0
      DAutoHeadings   =   0   'False
      EditModePermanent=   -1  'True
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      GrayAreaBackColor=   -2147483633
      MaxCols         =   15
      MaxRows         =   64
      ProcessTab      =   -1  'True
      SelectBlockOptions=   2
      SpreadDesigner  =   "frmedmod.frx":05CE
      UserResize      =   1
   End
   Begin HexUniControls.ctlUniFrameXP Frame1 
      Height          =   1335
      Left            =   6600
      Top             =   6720
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2355
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmedmod.frx":0860
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frmedmod.frx":089C
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmedmod.frx":08BC
      Begin HexUniControls.ctlUniCheckXP chk_nd 
         Height          =   450
         Left            =   2400
         TabIndex        =   3
         Top             =   860
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmedmod.frx":08D8
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmedmod.frx":0922
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmedmod.frx":0942
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_value 
         Height          =   450
         Left            =   120
         TabIndex        =   2
         Top             =   860
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmedmod.frx":095E
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmedmod.frx":0988
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmedmod.frx":09A8
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_rr 
         Height          =   450
         Left            =   2400
         TabIndex        =   1
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmedmod.frx":09C4
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmedmod.frx":09F4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmedmod.frx":0A14
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_md 
         Height          =   450
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmedmod.frx":0A30
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmedmod.frx":0A5C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmedmod.frx":0A7C
         ShowFocus       =   -1  'True
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_edit 
      Height          =   650
      Left            =   2280
      TabIndex        =   6
      Top             =   6360
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
      Caption         =   "frmedmod.frx":0A98
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmedmod.frx":0AD2
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmedmod.frx":0AF2
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9240
      Top             =   600
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8085
      FormDesignWidth =   11655
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_delete 
      Height          =   650
      Left            =   120
      TabIndex        =   5
      Top             =   6360
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
      Caption         =   "frmedmod.frx":0B0E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmedmod.frx":0B4C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmedmod.frx":0B6C
   End
   Begin HexUniControls.ctlUniTextBoxXP numprops 
      Height          =   375
      Left            =   6165
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7800
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmedmod.frx":0B88
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
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
      Tip             =   "frmedmod.frx":0BB2
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmedmod.frx":0BD2
   End
   Begin VBoard_EMD.KeySet_STD KeySet_STD1 
      Height          =   2055
      Left            =   4455
      Top             =   4150
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   3625
      BackColor       =   14737632
      LargeFontName   =   "Arial Unicode MS"
      LargeFontSize   =   6.75
      SmallFontName   =   "Arial Unicode MS"
      ButtonWidth     =   25
      ButtonHeight    =   25
      SendKeysEnabled =   -1  'True
      RegAtDT         =   -1  'True
   End
   Begin HexUniControls.ctlUniLabel Label4 
      Height          =   855
      Left            =   120
      Top             =   4680
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmedmod.frx":0BEE
      BackColor       =   -2147483633
      ForeColor       =   12583104
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmedmod.frx":0CCA
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmedmod.frx":0CEA
   End
   Begin HexUniControls.ctlUniLabel lblfile 
      Height          =   375
      Left            =   7560
      Top             =   120
      Width           =   4020
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmedmod.frx":0D06
      BackColor       =   -2147483633
      ForeColor       =   0
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmedmod.frx":0D26
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmedmod.frx":0D46
   End
   Begin HexUniControls.ctlUniLabel Label3 
      Height          =   375
      Left            =   6600
      Top             =   120
      Width           =   840
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmedmod.frx":0D62
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmedmod.frx":0D8A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmedmod.frx":0DAA
   End
   Begin HexUniControls.ctlUniLabel Label2 
      Height          =   375
      Left            =   30
      Top             =   600
      Width           =   1665
      _ExtentX        =   2937
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
      Caption         =   "frmedmod.frx":0DC6
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmedmod.frx":0DF4
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmedmod.frx":0E14
   End
   Begin HexUniControls.ctlUniLabel lblsampmode 
      Height          =   375
      Left            =   1800
      Top             =   600
      Width           =   4680
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmedmod.frx":0E30
      BackColor       =   -2147483633
      ForeColor       =   0
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmedmod.frx":0E52
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmedmod.frx":0E72
   End
   Begin HexUniControls.ctlUniLabel lblproduct 
      Height          =   375
      Left            =   1800
      Top             =   120
      Width           =   4680
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmedmod.frx":0E8E
      BackColor       =   -2147483633
      ForeColor       =   0
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmedmod.frx":0EAE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmedmod.frx":0ECE
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Left            =   30
      Top             =   120
      Width           =   1665
      _ExtentX        =   2937
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
      Caption         =   "frmedmod.frx":0EEA
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmedmod.frx":0F18
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmedmod.frx":0F38
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_deleteAll 
      Height          =   645
      Left            =   120
      TabIndex        =   11
      Top             =   5640
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
      Caption         =   "frmedmod.frx":0F54
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmedmod.frx":0F9E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmedmod.frx":0FBE
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_importPRD 
      Height          =   645
      Left            =   2280
      TabIndex        =   12
      Top             =   5640
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
      Caption         =   "frmedmod.frx":0FDA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmedmod.frx":101C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmedmod.frx":103C
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   10680
      Top             =   600
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
      Left            =   9840
      Top             =   600
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frmedmod.frx":1058
   End
End
Attribute VB_Name = "frmedmod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_addProp As Boolean
Public m_prdModelType As Boolean
Public m_prdFileName As String
Public m_stfFileName As String

Sub runsavemodels()
  
  checktablespread

  If (unity_main.tableok = True) Then
    frmedmod.Visible = False
    Call frm_collect.savescansettings(True)
  
    If (frmedmod.lblproduct.Caption = unity_main.current_product) Then
      Call unity_main.load_prod_file("", True)
      setupnewpredss
      unity_main.setup_olcols
    Else
      Call unity_main.load_prod_file("", False)
    End If
    
    frmProduct.txt_inifile.Text = ""
    frmProduct.txt_product.Text = ""
    frmProduct.txt_sampmode.Text = ""
    frmProduct.LSTPRODUCTS.ListIndex = -1
  End If
End Sub

Sub clearmodtable()
  
  frmedmod.grid_models.Row = 1
  frmedmod.grid_models.Col = 1
  frmedmod.grid_models.Row2 = MAX_NUM_PROPS
  frmedmod.grid_models.Col2 = frmedmod.grid_models.MaxCols
  frmedmod.grid_models.BlockMode = True
  frmedmod.grid_models.Action = ActionClear
  frmedmod.grid_models.BlockMode = False
  frmedmod.fixthesize
End Sub

Sub fixthesize()
  Dim xx As Integer
  
  frmedmod.grid_models.Col = 1
  frmedmod.grid_models.Col2 = 1
  frmedmod.grid_models.Row = 1
  frmedmod.grid_models.Row2 = MAX_NUM_PROPS
  frmedmod.grid_models.BlockMode = True
  frmedmod.grid_models.FontSize = 14
  frmedmod.grid_models.ColWidth(1) = 18
  frmedmod.grid_models.FontBold = True
  frmedmod.grid_models.BlockMode = False
  
  frmedmod.grid_models.ColWidth(2) = 25
  frmedmod.grid_models.ColWidth(15) = 25
  
  For xx = 3 To 10
    frmedmod.grid_models.ColWidth(xx) = 7
  Next xx
  
  frmedmod.grid_models.Col = 1
  frmedmod.grid_models.Row = 0
  frmedmod.grid_models.FontBold = False
  frmedmod.grid_models.FontSize = 8
End Sub

Sub fillproplist()

  ' Show Add Property button if max number of properties has not been configured
  If (frmedmod.numprops.Text < MAX_NUM_PROPS) Then
    frmedmod.cmd_add.Visible = True
  Else
    frmedmod.cmd_add.Visible = False
  End If

  frmedmod.fixthesize
End Sub

Private Sub setupnewpredss()
  Dim nn As Integer
  Dim tempstring As String
  
  For nn = 1 To unity_main.fpspread_pred.MaxRows
    unity_main.fpspread_pred.Row = nn
    unity_main.fpspread_pred.Col = 1
    frmedmod.grid_models.Col = 1
    frmedmod.grid_models.Row = nn
    tempstring = frmedmod.grid_models.Text
    unity_main.fpspread_pred.Text = Trim(tempstring)
  Next nn
  
  frmedmod.fixthesize
End Sub

Private Sub checktablespread()
  Dim modFileName As String
  Dim npropsm As Integer
  Dim lmint As Integer
  Dim lmval As Double
  Dim checknum As Boolean
  Dim modExt As String
  Dim lennum As Integer
  Dim onechar As String
  Dim rebuildit As String
  Dim wasf As String
  Dim fieldcounter As Integer
  Dim buildcounter As Integer
  Dim colcounter As Integer
  Dim curcol As Integer
  Dim pathspec As Boolean
  Dim zz As Integer
  Dim prdFileName As String
  Dim modStfFileName As String
  Dim stfFileName As String
  Dim constitName As String
  Dim startWaveln As Double
  Dim endWaveln As Double
  Dim wavelnInc As Double
  Dim specsChanged As Boolean
  Dim stdOK As Boolean
  Dim predOK As Boolean
  Dim numWaves As Long
  Dim rc As Long
  Dim uniMsg As String

  '1 = "Property"
  '2 = "Model"
  '3 = "Prop Index"
  '4 = "Intercept"
  '5 = "Slope"
  '6 = "Sig Figs"
  '7 = "MD Warn"
  '8 = "MD Fail"
  '9 = "Resid Warn"
  '10 = "Resid Fail"
  '11 = low warn
  '12 = low alarm
  '13 = high warn
  '14 = high alarm

  unity_main.tableok = True
  npropsm = frmedmod.grid_models.MaxRows

  modStfFileName = "NA"
#If ABBFT Then
  startWaveln = frm_collect.m_startWavenum
  endWaveln = frm_collect.m_endWavenum
  wavelnInc = frm_collect.m_waveNumIncr
#Else
  startWaveln = frm_collect.m_smplStartWvln
  endWaveln = frm_collect.m_smplEndWvln
  wavelnInc = MS11CfgData.wvlnIncr
#End If

  For zz = 1 To npropsm
    'check model name
    frmedmod.grid_models.Row = zz
    frmedmod.grid_models.Col = 2
    modFileName = Trim(frmedmod.grid_models.Text)
    
    If (modFileName = "") Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg1", "Please enter the model name including extension"), vbOKOnly
      GoTo Failedcheck
    End If
    
    'check for proper extension
    modExt = ("." & LCase(CFile.st_FileExt(modFileName)))
    
    If (modExt <> GRAMS_MODEL_FILE_EXT) And (modExt <> MLR_MODEL_FILE_EXT) And (modExt <> SEC_MODEL_FILE_EXT) And _
       (modExt <> CALSTAR_MODEL_FILE_EXT) And (modExt <> PRD_MODEL_FILE_EXT) Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg2-1", "Model name must include extension of either .mlr (for multiple linear regression), .cal (for Grams PLSiQ models) , .smf (for secondary models), .prd (for UCal PRD models) or .cpf (for CalStar Models)"), vbOKOnly
      frmedmod.grid_models.SetFocus
      GoTo Failedcheck
    Else
      ' Check if CalStar run-time software installed
      If (modExt = CALSTAR_MODEL_FILE_EXT) And (unity_main.calstar_enabled = False) Then
        frmedmod.grid_models.SetFocus
        CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg3", "Problem with product trying to use a CalStar model without CalStar software installed"), vbOKOnly
        GoTo Failedcheck
      Else
        If (modExt = PRD_MODEL_FILE_EXT) Then
          ' Check if PRD run-time software installed
          If (unity_main.m_prdEnabled = False) Then
            frmedmod.grid_models.SetFocus
            CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg21", "Problem with product trying to use a PRD model without UCal PRD software installed"), vbOKOnly
            GoTo Failedcheck
          Else
            ' Check if PRD model file the same for all properties
            If (prdFileName = "") Then
              prdFileName = modFileName
            Else
              If (prdFileName <> modFileName) Then
                frmedmod.grid_models.SetFocus
                CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg22", "UCal PRD model configured for each property must be same for product"), vbOKOnly
                GoTo Failedcheck
              End If
            End If
           
            ' Check if STF file the same for all properties
            frmedmod.grid_models.Col = 15
            stfFileName = Trim(frmedmod.grid_models.Text)
            
            If (modStfFileName = "NA") Then
              modStfFileName = stfFileName
            Else
              If (modStfFileName <> stfFileName) Then
                frmedmod.grid_models.SetFocus
                CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg23", "UCal STF file configured for each property must be same for product"), vbOKOnly
                GoTo Failedcheck
              End If
            End If

            ' Get property/constituent name
            frmedmod.grid_models.Col = 3
            constitName = Trim(frmedmod.grid_models.Text)

            ' Check if PRD & STF file spectrum compatible for product
            On Error GoTo OBJECT_ERROR
            rc = PRDObject.chkWavelnSpecs_4(prdFileName, stfFileName, constitName, startWaveln, endWaveln, wavelnInc, specsChanged, stdOK, predOK)
            
            If (rc <> 0) Then
              uniMsg = MLSupport.GGS_Params("errMsg8", "%1 %2 error: %3", prdFileName, "chkWavelnSpecs_4()", CStr(rc))
              CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
              GoTo Failedcheck
            Else
              If (predOK = False) Then
                frmedmod.grid_models.SetFocus
                uniMsg = MLSupport.GGS_Params("frmedmod.errMsg1", "%1 UCal PRD model file wavelength specs: (%2-%3, %4) are not valid for product wavelength specs:", prdFileName, CStr(startWaveln), CStr(endWaveln), CStr(wavelnInc))
#If ABBFT Then
                uniMsg = uniMsg & (": (" & frm_collect.m_startWavenum & "-" & frm_collect.m_endWavenum & ", " & frm_collect.m_waveNumIncr & ")")
#Else
                uniMsg = uniMsg & (": (" & frm_collect.m_smplStartWvln & "-" & frm_collect.m_smplEndWvln & ", " & MS11CfgData.wvlnIncr & ")")
#End If
                CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
                GoTo Failedcheck
              Else
                If (stdOK = False) And (stfFileName <> "") Then
                  rc = PRDObject.getStfWavelnSpecs(stfFileName, numWaves, startWaveln, endWaveln, wavelnInc, specsChanged)
                
                  If (rc <> 0) Then
                    uniMsg = MLSupport.GGS_Params("errMsg8", "%1 %2 error: %3", stfFileName, "getStfWavelnSpecs()", CStr(rc))
                  Else
                    frmedmod.grid_models.SetFocus
                    uniMsg = MLSupport.GGS_Params("frmedmod.errMsg2", "%1 UCal STF file wavelength specs: (%2-%3, %4) are not valid for product wavelength specs:", stfFileName, CStr(startWaveln), CStr(endWaveln), CStr(wavelnInc))
#If ABBFT Then
                    uniMsg = uniMsg & (": (" & frm_collect.m_startWavenum & "-" & frm_collect.m_endWavenum & ", " & frm_collect.m_waveNumIncr & ")")
#Else
                    uniMsg = uniMsg & (": (" & frm_collect.m_smplStartWvln & "-" & frm_collect.m_smplEndWvln & ", " & MS11CfgData.wvlnIncr & ")")
#End If
                  End If
                  
                  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
                  GoTo Failedcheck
                End If
              End If
            End If
            
            rc = PRDObject.chkSerialNumberCompatibility(prdFileName, constitName, stfFileName, unity_main.m_sysSerialNum)
            
            If (rc <> 0) Then
              uniMsg = MLSupport.GGS_Params("errMsg8", "%1 %2 error: %3", prdFileName, "chkSerialNumberCompatibility()", CStr(rc))
              CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
              GoTo Failedcheck
            End If
          End If
        End If
      End If
    End If
    
    ' Confirm path contains '\' instead of '/'
    check_filepathname_delimiters modFileName
    frmedmod.grid_models.Col = 2
    frmedmod.grid_models.Text = modFileName
    
    If (InStr(frmedmod.grid_models.Text, "\") = 0) Then
      frmedmod.grid_models.Text = (MODELS_DIR & modFileName)
    End If
    
    'check model index
    frmedmod.grid_models.Col = 3
    frmedmod.grid_models.Row = zz
    
    If (IsNumeric(Trim(frmedmod.grid_models.Text)) = True) Then
      lmint = CInt(Trim(frmedmod.grid_models.Text))
      lmval = Trim(frmedmod.grid_models.Text)
      
      If (lmval <> lmint) And (modExt <> CALSTAR_MODEL_FILE_EXT) And (modExt <> PRD_MODEL_FILE_EXT) Then
        frmedmod.grid_models.SetFocus
        CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg4", "Property index must be whole number > 0"), vbOKOnly
        GoTo Failedcheck
      End If
    End If
           
    'check property name
    frmedmod.grid_models.Col = 1
    frmedmod.grid_models.Row = zz
    
    If (Trim(frmedmod.grid_models.Text) = "") Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg5", "Please enter the property name"), vbOKOnly
      GoTo Failedcheck
    End If

    'check bias
    frmedmod.grid_models.Col = 4
    frmedmod.grid_models.Row = zz
  
    If (Trim(frmedmod.grid_models.Text) = "") Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg6", "Please enter an intercept value, 0 by default"), vbOKOnly
      GoTo Failedcheck
    End If
    
    If (IsNumeric(Trim(frmedmod.grid_models.Text)) = False) Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg7", "Please enter a numeric value for the intercept, 0 by default"), vbOKOnly
      GoTo Failedcheck
    End If

    'check skew
    frmedmod.grid_models.Col = 5
    frmedmod.grid_models.Row = zz
    
    If (IsNumeric(Trim(frmedmod.grid_models.Text)) = False) Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg8", "Please enter a numeric value for the slope, 1 by default"), vbOKOnly
      GoTo Failedcheck
    End If

    'check sigfigs
    frmedmod.grid_models.Col = 6
    frmedmod.grid_models.Row = zz
    
    If (Trim(frmedmod.grid_models.Text) = "") Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg9", "Please enter a whole number for the significant figures"), vbOKOnly
      GoTo Failedcheck
    End If
 
    'sigfigs numeric?
    If (IsNumeric(Trim(frmedmod.grid_models.Text)) = True) Then
      lmint = CInt(Trim(frmedmod.grid_models.Text))
      lmval = Trim(frmedmod.grid_models.Text)
      
      If (lmval <> lmint) Then
        frmedmod.grid_models.SetFocus
        CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg9", "Please enter a whole number for the significant figures"), vbOKOnly
        GoTo Failedcheck
      End If
    Else
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg9", "Please enter a whole number for the significant figures"), vbOKOnly
      GoTo Failedcheck
    End If

    'check mdwarn
    frmedmod.grid_models.Col = 7
    frmedmod.grid_models.Row = zz
    
    If (Trim(frmedmod.grid_models.Text) = "") Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg10", "Please enter a value for md warning, 1.2 by default"), vbOKOnly
      GoTo Failedcheck
    End If
    
    If (IsNumeric(Trim(frmedmod.grid_models.Text)) = False) Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg10", "Please enter a numeric value for md warning, 1.2 by default"), vbOKOnly
      GoTo Failedcheck
    End If

    'check mdfail
    frmedmod.grid_models.Col = 8
    frmedmod.grid_models.Row = zz
    
    If (Trim(frmedmod.grid_models.Text) = "") Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg11", "Please enter a value for md fail, 1.5 by default"), vbOKOnly
      GoTo Failedcheck
    End If
    
    If (IsNumeric(Trim(frmedmod.grid_models.Text)) = False) Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg11", "Please enter a numeric value for md fail, 1.5 by default"), vbOKOnly
      GoTo Failedcheck
    End If
 
    'check rrwarn
    frmedmod.grid_models.Col = 9
    frmedmod.grid_models.Row = zz
    
    If (Trim(frmedmod.grid_models.Text) = "") Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg12", "Please enter a value for Resid warning, 1.5 by default"), vbOKOnly
      GoTo Failedcheck
    End If
    
    If (IsNumeric(Trim(frmedmod.grid_models.Text)) = False) Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg12", "Please enter a numeric value for Resid warning, 1.5 by default"), vbOKOnly
      GoTo Failedcheck
    End If

    'check rrfail
    frmedmod.grid_models.Col = 10
    frmedmod.grid_models.Row = zz
    
    If (Trim(frmedmod.grid_models.Text) = "") Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg13", "Please enter a value for Resid fail, 3 by default"), vbOKOnly
      GoTo Failedcheck
    End If
    
    If (IsNumeric(Trim(frmedmod.grid_models.Text)) = False) Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg13", "Please enter a numeric value for Resid fail, 3 by default"), vbOKOnly
      GoTo Failedcheck
    End If

    'check value low warn
    frmedmod.grid_models.Col = 11
    frmedmod.grid_models.Row = zz
  
    If (Trim(frmedmod.grid_models.Text) = "") Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg14", "Please enter a value for Low Value Warn"), vbOKOnly
      GoTo Failedcheck
    End If
    
    If (IsNumeric(Trim(frmedmod.grid_models.Text)) = False) Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg14", "Please enter a value for Low Value Warn"), vbOKOnly
      GoTo Failedcheck
    End If
  
    'check value low alarm
    frmedmod.grid_models.Col = 12
    frmedmod.grid_models.Row = zz
    
    If (Trim(frmedmod.grid_models.Text) = "") Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg15", "Please enter a value for Low Value Alarm"), vbOKOnly
      GoTo Failedcheck
    End If
    
    If (IsNumeric(Trim(frmedmod.grid_models.Text)) = False) Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg15", "Please enter a value for Low Value Alarm"), vbOKOnly
      GoTo Failedcheck
    End If

    'check value high warn
    frmedmod.grid_models.Col = 13
    frmedmod.grid_models.Row = zz
    
    If (Trim(frmedmod.grid_models.Text) = "") Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg16", "Please enter a value for High Value Warn"), vbOKOnly
      GoTo Failedcheck
    End If
    
    If (IsNumeric(Trim(frmedmod.grid_models.Text)) = False) Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg16", "Please enter a value for High Value Warn"), vbOKOnly
      GoTo Failedcheck
    End If

    'check value high alarm
    frmedmod.grid_models.Col = 14
    frmedmod.grid_models.Row = zz
    
    If (Trim(frmedmod.grid_models.Text) = "") Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg17", "Please enter a value for High Value Alarm"), vbOKOnly
      GoTo Failedcheck
    End If
    
    If (IsNumeric(Trim(frmedmod.grid_models.Text)) = False) Then
      frmedmod.grid_models.SetFocus
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg17", "Please enter a value for High Value Alarm"), vbOKOnly
      GoTo Failedcheck
    End If
  Next zz
  
  GoTo looksok
  
OBJECT_ERROR:
  unity_main.errorstring = "Unity PRDComponent.dll component not installed or registered"
  unity_main.write_error (LOG_DBG_LEVEL1)
  uniMsg = MLSupport.GGS_Params("errMsg7", "Unity %1 component not installed or registered", "PRDComponent.dll")
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  
Failedcheck:
  unity_main.tableok = False
  
looksok:
  frmedmod.fixthesize
End Sub

Private Sub cmd_add_Click()
  
  unity_main.errorstring = "Product Properties Management screen Add Property button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  frm_modtype.Show 1
End Sub

Private Sub cmd_cancel_Click()
  
  unity_main.errorstring = "Product Properties Management screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  unity_main.setup_olcols
  frmedmod.Visible = False
  
  frmProduct.txt_inifile.Text = ""
  frmProduct.txt_product.Text = ""
  frmProduct.txt_sampmode.Text = ""
  frmProduct.LSTPRODUCTS.ListIndex = -1
End Sub

Private Sub cmd_delete_Click()
  Dim actRow As Integer
  Dim PropName As String
  Dim cnt As Integer
  Dim optVal As Integer
  Dim tempstring As String
  Dim modExt As String
  Dim constituentName As String
  Dim ii As Integer
  
  unity_main.errorstring = "Product Properties Management screen Delete Property button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  actRow = frmedmod.grid_models.ActiveRow
  
  If (actRow < 1) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg18", "Please Select a property to delete from the list"), vbOKOnly
    Exit Sub
  End If

  frmedmod.grid_models.Col = 1
  frmedmod.grid_models.Row = actRow
  PropName = frmedmod.grid_models.Text

  optVal = CWrap.ShowMessageBoxW(MLSupport.GGS_Params("frmedmod.statMsg1", "Are you sure you want to delete property %1?", PropName), vbYesNo)
  
  If (optVal = vbYes) Then
    ' Get model file name extension
    frmedmod.grid_models.Row = actRow
    frmedmod.grid_models.Col = 2
    tempstring = frmedmod.grid_models.Text
    modExt = ("." & LCase(CFile.st_FileExt(tempstring)))
  
    ' Check if PRD model file
    If (modExt = PRD_MODEL_FILE_EXT) Then
      ' Get constituent name
      frmedmod.grid_models.Col = 3
      constituentName = Trim(frmedmod.grid_models.Text)
    
      For ii = 0 To lst_prdConstituentNames.ListCount - 1
        If (constituentName = lst_prdConstituentNames.List(ii)) Then
          lst_prdConstituentNames.RemoveItem (ii)
          Exit For
        End If
      Next ii
    End If
  
    cnt = frmedmod.numprops.Text - 1
    frmedmod.numprops.Text = cnt
    frmedmod.grid_models.Row = actRow
    frmedmod.grid_models.Action = ActionDeleteRow
    frmedmod.grid_models.MaxRows = cnt
    frmedmod.fillproplist
    
    If (lst_prdConstituentNames.ListCount = 0) Then
      frmedmod.m_prdModelType = False
      frmedmod.m_prdFileName = ""
      frmedmod.m_stfFileName = ""
    End If
  End If
  
  frmedmod.grid_models.SetActiveCell 0, 0
End Sub

Private Sub cmd_deleteAll_Click()
  Dim optVal As Integer

  unity_main.errorstring = "Product Properties Management screen Delete All Properties button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  optVal = CWrap.ShowMessageBoxW(MLSupport.GSS("frmedmod", "statMsg1", "Are you sure you want to delete all properties?"), vbYesNo)
  
  If (optVal = vbYes) Then
    frmedmod.grid_models.Row = 1
    frmedmod.grid_models.Col = 2
    frmedmod.grid_models.Row2 = frmedmod.grid_models.MaxRows
    frmedmod.grid_models.Col2 = frmedmod.grid_models.MaxCols
    frmedmod.grid_models.BlockMode = True
    frmedmod.grid_models.Action = ActionClear
    frmedmod.grid_models.BlockMode = False
    frmedmod.grid_models.MaxRows = 0
    
    frmedmod.m_prdModelType = False
    frmedmod.m_prdFileName = ""
    frmedmod.m_stfFileName = ""
    frmedmod.lst_prdConstituentNames.Clear
  End If
  
  frmedmod.grid_models.SetActiveCell 0, 0
End Sub

Private Sub cmd_edit_Click()
  Dim tempstring As String
  Dim modExt As String
  Dim actRow As Integer
  Dim zz As Integer
  
  unity_main.errorstring = "Product Properties Management screen Edit Property button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  actRow = frmedmod.grid_models.ActiveRow
  
  If (actRow < 1) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg19", "Please Select a property to edit from the list"), vbOKOnly
    Exit Sub
  End If
  
  ' Get model file name extension
  frmedmod.grid_models.Col = 2
  frmedmod.grid_models.Row = actRow
  tempstring = frmedmod.grid_models.Text
  modExt = ("." & LCase(CFile.st_FileExt(tempstring)))
 
  If (modExt = PRD_MODEL_FILE_EXT) Then
    ' Check if .NET Framework 2.0 is not installed
    If (unity_main.m_netFWInstalled = False) Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg20", "Cannot edit an UCal PRD model without MS .NET Framework 2.0 installed!"), vbCritical
      frmedmod.grid_models.SetActiveCell 0, 0
      Exit Sub
    End If
    
    frm_prd.uniTitle = MLSupport.GSS("frm_prd", "uniTitle2", "UCal PRD Model Property Configuration")
    frm_prd.cmd_browsePRD.Visible = False
    frm_prd.Label16.Visible = True
    frm_prd.txt_propname.Visible = True
    frm_prd.Picture1.Visible = True
    frm_prd.Label1(1).Visible = True
    frm_prd.Label1(0).Visible = False
    frm_prd.List1.MultiSelect = 0
    frm_prd.List1.enabled = False
    frm_prd.cmd_selectAll.Visible = False
    frm_prd.cmd_unselectAll.Visible = False
    frm_prd.cmd_import.Visible = False
    frm_prd.cmd_save.Visible = True
    
    frm_prd.txtmrow.Text = actRow
    frm_prd.m_prdFileName = Trim(tempstring)
    frm_prd.lbl_prdName.Caption = Trim(tempstring)
    frmedmod.grid_models.Col = 1
    frm_prd.txt_propname.Text = Trim(frmedmod.grid_models.Text)
    frmedmod.grid_models.Col = 3
    Call frm_prd.set_constituent_list_index(Trim(frmedmod.grid_models.Text))
    
    frmedmod.grid_models.Col = 15
    
    If (Trim(frmedmod.grid_models.Text) <> "NA") Then
      frm_prd.m_stfFileName = Trim(frmedmod.grid_models.Text)
    Else
      frm_prd.m_stfFileName = ""
    End If
  
    frm_prd.txt_stfName.Text = frm_prd.m_stfFileName
  
    For zz = 8 To 9
      frm_prd.txt_modvar(zz).Visible = False
      frm_prd.Label1(zz).Visible = False
    Next zz
      
    frm_prd.Picture9.Visible = False
    frm_prd.Picture10.Visible = False
  
    For zz = 4 To frmedmod.grid_models.MaxCols - 1
      frmedmod.grid_models.Col = zz
      frm_prd.txt_modvar(zz - 1).Text = frmedmod.grid_models.Text
    Next zz
    
    frm_prd.Show 1
    frmedmod.grid_models.SetActiveCell 0, 0
    Exit Sub
  End If
  
  If (modExt = CALSTAR_MODEL_FILE_EXT) Then
    frmMain.txtmrow.Text = actRow
    frmMain.pname2 = Trim(tempstring)
    frmMain.LBL_Calibname.Caption = Trim(tempstring)
    Call frmMain.opentoedit
    frmedmod.grid_models.Col = 1
    frmMain.txt_propname.Text = Trim(frmedmod.grid_models.Text)
    frmedmod.grid_models.Col = 3
    unity_main.slcal = Trim(frmedmod.grid_models.Text)
    Call frmMain.parse_cs
  
    For zz = 3 To frmedmod.grid_models.MaxCols - 1
      frmedmod.grid_models.Col = zz
      frmMain.txt_modvar(zz - 1).Text = frmedmod.grid_models.Text
    Next zz
    
    frmMain.Show 1
    frmedmod.grid_models.SetActiveCell 0, 0
    Exit Sub
  End If
  
  frm_1model.txtmrow.Text = actRow
  frm_1model.modtypex = 0
  
  If (modExt = MLR_MODEL_FILE_EXT) Or (modExt = SEC_MODEL_FILE_EXT) Then
    If (modExt = MLR_MODEL_FILE_EXT) Then
      frm_1model.uniTitle = MLSupport.GSS("frm_1model", "uniTitle2", "MLR Model Property Configuration")
      frm_1model.modtypex = 2
    Else
      frm_1model.uniTitle = MLSupport.GSS("frm_1model", "uniTitle3", "Secondary Model Property Configuration")
      frm_1model.modtypex = 3
    End If
    
    For zz = 6 To 9
      frm_1model.txt_modvar(zz).Visible = False
      frm_1model.Label1(zz).Visible = False
    Next zz
      
    frm_1model.Picture7.Visible = False
    frm_1model.Picture8.Visible = False
    frm_1model.Picture9.Visible = False
    frm_1model.Picture10.Visible = False
  Else
    frm_1model.uniTitle = MLSupport.GSS("frm_1model", "uniTitle1", "GRAMS PLSIQ Model Property Configuration")
    frm_1model.modtypex = 1
  End If

  For zz = 1 To frmedmod.grid_models.MaxCols - 1
    frmedmod.grid_models.Col = zz
    frm_1model.txt_modvar(zz - 1).Text = frmedmod.grid_models.Text
  Next zz
  
  frm_1model.Show 1
  frmedmod.grid_models.SetActiveCell 0, 0
End Sub

Private Sub cmd_importPRD_Click()
  Dim ff As Integer

  unity_main.errorstring = "Product Properties Management screen Import PRD Model button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  ' Check if .NET Framework 2.0 is not installed
  If (unity_main.m_netFWInstalled = False) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "errMsg20", "Cannot edit an UCal PRD model without MS .NET Framework 2.0 installed!"), vbCritical
    frmedmod.grid_models.SetActiveCell 0, 0
    Exit Sub
  End If

  frm_prd.lbl_prdName.Caption = m_prdFileName
  frm_prd.m_prdFileName = m_prdFileName
  frm_prd.txt_stfName.Text = m_stfFileName
  frm_prd.m_stfFileName = m_stfFileName

  If (m_prdModelType = True) Then
    If (frm_prd.build_partial_constituent_list() <> 0) Then
      Exit Sub
    Else
      If (frm_prd.List1.ListCount = 0) Then
        CWrap.ShowMessageBoxW MLSupport.GSS("frmedmod", "statMsg2", "No more constituents to import"), vbOKOnly
        Exit Sub
      Else
        frm_prd.cmd_browsePRD.Visible = False
      End If
    End If
  Else
    frm_prd.cmd_browsePRD.Visible = True
  End If

  frm_prd.uniTitle = MLSupport.GSS("frm_prd", "uniTitle1", "UCal PRD Model Import")
  
  frm_prd.Label16.Visible = False
  frm_prd.txt_propname.Visible = False
  frm_prd.Picture1.Visible = False
  frm_prd.Label1(0).Visible = True
  frm_prd.Label1(1).Visible = False
  frm_prd.List1.MultiSelect = 1
  frm_prd.List1.enabled = True
  frm_prd.cmd_selectAll.Visible = True
  frm_prd.cmd_unselectAll.Visible = True
  frm_prd.cmd_import.Visible = True
  frm_prd.cmd_save.Visible = False
  
  For ff = 8 To 9
    frm_prd.txt_modvar(ff).Visible = False
    frm_prd.Label1(ff).Visible = False
  Next ff
      
  frm_prd.Picture9.Visible = False
  frm_prd.Picture10.Visible = False
    
  frm_prd.txt_modvar(3).Text = "0" ' intercept
  frm_prd.txt_modvar(4).Text = "1" ' slope
  frm_prd.txt_modvar(5).Text = "2" ' sig figs
  frm_prd.txt_modvar(6).Text = "3.0" ' m dist warn
  frm_prd.txt_modvar(7).Text = "7.0" ' m dist fail
  frm_prd.txt_modvar(8).Text = "3" ' resid rat warn
  frm_prd.txt_modvar(9).Text = "7" ' rr fail
  frm_prd.txt_modvar(10).Text = "1" ' prop low warn
  frm_prd.txt_modvar(11).Text = "0.5" ' prop low fail
  frm_prd.txt_modvar(12).Text = "99" ' prop hi warn
  frm_prd.txt_modvar(13).Text = "100" ' prop hi fail
  frm_prd.txtmrow.Text = frmedmod.grid_models.MaxRows
  frm_prd.Show 1
End Sub

Private Sub cmd_save_Click()
  
  unity_main.errorstring = "Product Properties Management screen Save Changes button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  frmedmod.runsavemodels
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me

   ' Label columns
  frmedmod.grid_models.Row = 0
  frmedmod.grid_models.Col = 1
  frmedmod.grid_models.Text = MLSupport.GSS("Headers", "property", "Property")
  frmedmod.grid_models.Col = 2
  frmedmod.grid_models.Text = MLSupport.GSS("Headers", "model", "Model")
  frmedmod.grid_models.Col = 3
  frmedmod.grid_models.Text = MLSupport.GSS("Headers", "propIndx", "Prop Indx")
  frmedmod.grid_models.Col = 4
  frmedmod.grid_models.Text = MLSupport.GSS("Headers", "intercept", "Intercept")
  frmedmod.grid_models.Col = 5
  frmedmod.grid_models.Text = MLSupport.GSS("Headers", "slope", "Slope")
  frmedmod.grid_models.Col = 6
  frmedmod.grid_models.Text = MLSupport.GSS("Headers", "sigFigs", "Sig Figs")
  frmedmod.grid_models.Col = 7
  frmedmod.grid_models.Text = MLSupport.GSS("Headers", "mdWarn", "MD Warn")
  frmedmod.grid_models.Col = 8
  frmedmod.grid_models.Text = MLSupport.GSS("Headers", "mdFail", "MD Fail")
  frmedmod.grid_models.Col = 9
  frmedmod.grid_models.Text = MLSupport.GSS("Headers", "resWarn", "Res Warn")
  frmedmod.grid_models.Col = 10
  frmedmod.grid_models.Text = MLSupport.GSS("Headers", "resFail", "Res Fail")
  frmedmod.grid_models.Col = 11
  frmedmod.grid_models.Text = MLSupport.GSS("Headers", "lowWarn", "Low Warn")
  frmedmod.grid_models.Col = 12
  frmedmod.grid_models.Text = MLSupport.GSS("Headers", "lowFail", "Low Fail")
  frmedmod.grid_models.Col = 13
  frmedmod.grid_models.Text = MLSupport.GSS("Headers", "highWarn", "High Warn")
  frmedmod.grid_models.Col = 14
  frmedmod.grid_models.Text = MLSupport.GSS("Headers", "highFail", "High Fail")
  frmedmod.grid_models.Col = 15
  frmedmod.grid_models.Text = MLSupport.GSS("Headers", "stfFile", "STF File")
End Sub











