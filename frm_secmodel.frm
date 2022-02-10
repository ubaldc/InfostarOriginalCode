VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_secmodel 
   Caption         =   "Secondary Model Design"
   ClientHeight    =   7155
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   10725
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
   Icon            =   "frm_secmodel.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin FPUSpreadADO.fpSpread ss_cprop 
      Height          =   2295
      Left            =   960
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
      Width           =   8775
      _Version        =   458752
      _ExtentX        =   15478
      _ExtentY        =   4048
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   50
      MaxRows         =   35
      SpreadDesigner  =   "frm_secmodel.frx":08CA
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_eqn 
      Height          =   735
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5280
      Width           =   10215
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483633
      ForeColor       =   0
      Enabled         =   0   'False
      Locked          =   0   'False
      Text            =   "frm_secmodel.frx":0B17
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
      MultiLine       =   -1  'True
      Alignment       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frm_secmodel.frx":0B37
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":0B57
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_clearLastEntry 
      Height          =   650
      Left            =   4200
      TabIndex        =   10
      Top             =   4320
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
      Caption         =   "frm_secmodel.frx":0B73
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_secmodel.frx":0BB3
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":0BD3
   End
   Begin HexUniControls.ctlUniTextBoxXP Text1 
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      Top             =   6720
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_secmodel.frx":0BEF
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
      Tip             =   "frm_secmodel.frx":0C19
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":0C39
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   650
      Left            =   6240
      TabIndex        =   12
      Top             =   6120
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
      Caption         =   "frm_secmodel.frx":0C55
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_secmodel.frx":0C8D
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":0CAD
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_mname2 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1200
      Width           =   3135
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_secmodel.frx":0CC9
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
      Tip             =   "frm_secmodel.frx":0CE9
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":0D09
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_browse 
      Height          =   650
      Left            =   6120
      TabIndex        =   4
      Top             =   1080
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
      Caption         =   "frm_secmodel.frx":0D25
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_secmodel.frx":0D51
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":0D71
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_mname 
      Height          =   375
      Left            =   3480
      TabIndex        =   20
      Top             =   8160
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_secmodel.frx":0D8D
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
      Tip             =   "frm_secmodel.frx":0DAD
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":0DCD
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modinfo 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   6015
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_secmodel.frx":0DE9
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
      Tip             =   "frm_secmodel.frx":0E09
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":0E29
   End
   Begin HexUniControls.ctlUniTextBoxXP Text6 
      Height          =   615
      Left            =   2520
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1935
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      Enabled         =   0   'False
      Locked          =   0   'False
      Text            =   "frm_secmodel.frx":0E45
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   14.25
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
      Tip             =   "frm_secmodel.frx":0E65
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":0E85
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_test 
      Height          =   650
      Left            =   360
      TabIndex        =   11
      Top             =   6120
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
      Caption         =   "frm_secmodel.frx":0EA1
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_secmodel.frx":0EDB
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":0EFB
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_clearModel 
      Height          =   650
      Left            =   4220
      TabIndex        =   9
      Top             =   3600
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
      Caption         =   "frm_secmodel.frx":0F17
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_secmodel.frx":0F4D
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":0F6D
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9720
      Top             =   240
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7155
      FormDesignWidth =   10725
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_cprop 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   720
      Width           =   3135
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_secmodel.frx":0F89
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
      Tip             =   "frm_secmodel.frx":0FA9
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":0FC9
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_addEquNum 
      Height          =   650
      Left            =   4220
      TabIndex        =   8
      Top             =   2880
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
      Caption         =   "frm_secmodel.frx":0FE5
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_secmodel.frx":1031
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":1051
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_value 
      Height          =   375
      Left            =   4220
      TabIndex        =   7
      Top             =   2400
      Width           =   1920
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_secmodel.frx":106D
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
      Tip             =   "frm_secmodel.frx":108D
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":10AD
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   8520
      TabIndex        =   0
      Top             =   6120
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
      Caption         =   "frm_secmodel.frx":10C9
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_secmodel.frx":10F5
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":1115
   End
   Begin HexUniControls.ctlUniListBoxXP lst_props 
      Height          =   2160
      Left            =   1920
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TrapTab         =   0   'False
      Tip             =   "frm_secmodel.frx":1131
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   -1  'True
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":1151
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniListBoxXP lst_functs 
      Height          =   2160
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TrapTab         =   0   'False
      Tip             =   "frm_secmodel.frx":116D
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   -1  'True
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":118D
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txtnvar 
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_secmodel.frx":11A9
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
      Tip             =   "frm_secmodel.frx":11D3
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":11F3
   End
   Begin HexUniControls.ctlUniListBoxXP lst_type 
      Height          =   2205
      Left            =   2640
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
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
      Tip             =   "frm_secmodel.frx":120F
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":122F
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniListBoxXP lst_what 
      Height          =   1425
      Left            =   720
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TrapTab         =   0   'False
      Tip             =   "frm_secmodel.frx":124B
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":126B
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniLabel Label11 
      Height          =   1935
      Left            =   6480
      Top             =   3240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3413
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_secmodel.frx":1287
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_secmodel.frx":14FD
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":151D
   End
   Begin HexUniControls.ctlUniLabel Label10 
      Height          =   1095
      Left            =   6480
      Top             =   1920
      Width           =   3615
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
      Caption         =   "frm_secmodel.frx":1539
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_secmodel.frx":16DF
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":16FF
   End
   Begin HexUniControls.ctlUniLabel Label9 
      Height          =   300
      Left            =   1920
      Top             =   1920
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_secmodel.frx":171B
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_secmodel.frx":174F
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":176F
   End
   Begin HexUniControls.ctlUniLabel Label8 
      Height          =   300
      Left            =   240
      Top             =   1920
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_secmodel.frx":178B
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_secmodel.frx":17BD
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":17DD
   End
   Begin HexUniControls.ctlUniLabel Label5 
      Height          =   300
      Left            =   150
      Top             =   4920
      Width           =   3015
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_secmodel.frx":17F9
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_secmodel.frx":1835
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":1855
   End
   Begin HexUniControls.ctlUniLabel Label2 
      Height          =   300
      Left            =   720
      Top             =   1245
      Width           =   1935
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_secmodel.frx":1871
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_secmodel.frx":18AF
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":18CF
   End
   Begin HexUniControls.ctlUniLabel Label6 
      Height          =   300
      Left            =   150
      Top             =   765
      Width           =   2535
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_secmodel.frx":18EB
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_secmodel.frx":193B
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":195B
   End
   Begin HexUniControls.ctlUniLabel Label3 
      Height          =   255
      Left            =   4215
      Top             =   1920
      Width           =   1920
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_secmodel.frx":1977
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_secmodel.frx":19B3
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":19D3
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   300
      Left            =   360
      Top             =   285
      Width           =   2295
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_secmodel.frx":19EF
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_secmodel.frx":1A31
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_secmodel.frx":1A51
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   9840
      Top             =   1320
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
      Left            =   9720
      Top             =   840
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_secmodel.frx":1A6D
   End
End
Attribute VB_Name = "frm_secmodel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub domath()
  
  Select Case unity_main.curfunct
    Case "+"
        unity_main.val1 = unity_main.val1 + unity_main.val2
    Case "-"
        unity_main.val1 = unity_main.val1 - unity_main.val2
    Case "/"
        unity_main.val1 = unity_main.val1 / unity_main.val2
    Case "*"
        unity_main.val1 = unity_main.val1 * unity_main.val2
  End Select
End Sub

Private Sub cmd_cancel_Click()
  
  unity_main.errorstring = "Secondary Model Design screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Unload frm_secmodel
End Sub

Private Sub cmd_addEquNum_Click()
  unity_main.errorstring = "Secondary Model Design screen Add Number to Equation button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  If IsNumeric(frm_secmodel.txt_value.Text) Then
    frm_secmodel.lst_type.AddItem ("C")
    frm_secmodel.lst_what.AddItem (frm_secmodel.txt_value.Text)
    frm_secmodel.txt_eqn.Text = frm_secmodel.txt_eqn.Text & frm_secmodel.txt_value.Text
  Else
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_secmodel", "errMsg1", "Please enter a valid number"), vbExclamation
  End If
End Sub

Private Sub cmd_clearLastEntry_Click()
  Dim ff, zz As Integer
  Dim modstring As String
  
  unity_main.errorstring = "Secondary Model Design screen Clear Last Entry button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  ff = frm_secmodel.lst_type.ListCount
  
  If (ff > 0) Then
    frm_secmodel.lst_type.RemoveItem (ff - 1)
    ff = frm_secmodel.lst_what.ListCount
    frm_secmodel.lst_what.RemoveItem (ff - 1)
    
    modstring = ""
    
    For zz = 0 To (frm_secmodel.lst_what.ListCount - 1)
      modstring = modstring & frm_secmodel.lst_what.List(zz)
    Next zz
    
    frm_secmodel.txt_eqn.Text = modstring
  End If
End Sub

Private Sub cmd_clearModel_Click()
  
  unity_main.errorstring = "Secondary Model Design screen Clear Model button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  frm_secmodel.ss_cprop.Row = 0
  frm_secmodel.ss_cprop.Col = 0
  frm_secmodel.ss_cprop.Row2 = 35
  frm_secmodel.ss_cprop.Col2 = 16
  frm_secmodel.ss_cprop.BlockMode = True
  frm_secmodel.ss_cprop.Action = ActionClear
  frm_secmodel.ss_cprop.BlockMode = False
  txt_eqn.Text = ""
  lst_type.Clear
  lst_what.Clear
  Unload frm_secmodel
  frm_secmodel.Show 1
End Sub

Private Sub cmd_browse_Click()
  Dim dialog As New clsCommDialogs
  Dim fileDir As String
  Dim sFilter As String
  Dim dlgTitle As String
  Dim fileName As String
  Dim tempstring As String
  Dim nVar As Integer
  Dim fileToOpen As String
  Dim zz As Integer
  Dim lineCnt As Integer
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim rc As Boolean
  Dim varStr As Variant
  Dim errMsg As String
  Dim uniMsg As String
  
  unity_main.errorstring = "Secondary Model Design screen Browse button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  dialog.InitDialogs
  fileDir = MODELS_DIR
  sFilter = ("Secondary Model File (*" & SEC_MODEL_FILE_EXT & ")" & Chr(0) & "*" & SEC_MODEL_FILE_EXT & Chr(0))
  dlgTitle = MLSupport.GSS("frm_secmodel", "dlgTitle", "Select Secondary Model File to Edit")
  fileName = dialog.ShowOpen(Me.hwnd, sFilter, dlgTitle, OFN_FILEMUSTEXIST + OFN_HIDEREADONLY + OFN_NOCHANGEDIR, fileDir)
 
  If (fileName <> "") Then
    txt_mname.Text = fileName
    txt_mname2.Text = CFile.st_FileName(fileName)
    fileToOpen = txt_mname.Text
    
    If (uniFile.OpenFileRead(fileToOpen) = True) Then
      On Error GoTo BAD_FILE
      fEncoding = uniFile.ReadBOM
      lineCnt = lineCnt + 1
      
      If (fEncoding = fe_ANSI) Then
        rc = uniFile.ReadAnsiLine(tempstring)
      Else
        rc = uniFile.ReadUnicodeLine(tempstring)
      End If
      
      If (rc = False) Then GoTo BAD_FILE
  
      txt_cprop.Text = tempstring
  
      lineCnt = lineCnt + 1
      
      If (fEncoding = fe_ANSI) Then
        rc = uniFile.ReadAnsiLine(tempstring)
      Else
        rc = uniFile.ReadUnicodeLine(tempstring)
      End If
      
      If (rc = False) Then GoTo BAD_FILE
  
      txt_modinfo.Text = tempstring
      lineCnt = lineCnt + 1
      
      If (fEncoding = fe_ANSI) Then
        rc = uniFile.ReadAnsiLine(tempstring)
      Else
        rc = uniFile.ReadUnicodeLine(tempstring)
      End If
      
      If (rc = False) Then GoTo BAD_FILE
  
      txtnvar.Text = tempstring
      nVar = tempstring
  
      lst_type.Clear
      lst_what.Clear

      For zz = 1 To nVar
        lineCnt = lineCnt + 1
        
        If (fEncoding = fe_ANSI) Then
          rc = uniFile.ReadAnsiLine(tempstring)
        Else
          rc = uniFile.ReadUnicodeLine(tempstring)
        End If
      
        If (rc = False) Then GoTo BAD_FILE
      
        varStr = Split(tempstring, ",")
        ss_cprop.Col = zz + 1
        ss_cprop.Row = 1
        ss_cprop.Text = Trim(varStr(1))
        lst_type.AddItem Trim(varStr(1))
      
        ss_cprop.Row = 2
        ss_cprop.Text = Trim(varStr(0))
        lst_what.AddItem Trim(varStr(0))
      Next zz
  
      tempstring = ""
      frm_secmodel.ss_cprop.Row = 2
  
      For zz = 2 To nVar + 1
        frm_secmodel.ss_cprop.Col = zz
    
        If (zz = 2) Then
          tempstring = Trim(frm_secmodel.ss_cprop.Text)
        Else
          tempstring = (tempstring & " " & Trim(frm_secmodel.ss_cprop.Text))
        End If
      Next zz
  
      txt_eqn.Text = tempstring
    Else
BAD_FILE:
      If (lineCnt = 0) Then
        errMsg = (fileToOpen & " file open error." & Error$)
        uniMsg = MLSupport.GGS_Params("fileErrMsg2", "%1 file open error. %2", fileToOpen, Error$)
      Else
        errMsg = (fileToOpen & " file has error on line " & CStr(lineCnt) & ". " & Error$)
        uniMsg = MLSupport.GGS_Params("fileErrMsg3", "%1 file has error on line %2. %3", fileToOpen, CStr(lineCnt), Error$)
      End If
  
      unity_main.errorstring = errMsg
      unity_main.write_error
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    End If
    
    uniFile.CloseFile
  End If
End Sub

Private Sub cmd_test_Click()
  Dim x As Integer
  
  unity_main.errorstring = "Secondary Model Design screen Test Equation button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  If (lst_type.ListCount = 0) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_secmodel", "errMsg2", "No model loaded or defined"), vbExclamation
  Else
    txtnvar.Text = lst_type.ListCount
    
    For x = 0 To lst_type.ListCount - 1
      ss_cprop.Col = x + 2
      ss_cprop.Row = 1
      ss_cprop.Text = lst_type.List(x)
      
      ss_cprop.Row = 2
      ss_cprop.Text = lst_what.List(x)
    Next x
  
    Call dosecpred2(True)
  End If
End Sub
Sub dosecpred2(offLine As Boolean)
  Dim tempstring As Variant
  Dim tempstring2 As String
  Dim callmath As Boolean
  Dim firstone As Boolean
  Dim typefunct As Boolean
  Dim lbindx As Integer
  Dim rbindx As Integer
  Dim countval As Integer
  Dim countfunct As Integer
  Dim freefilex As Integer
  Dim nt, zz, nn As Integer
  Dim errMsg As String
  Dim uniMsg As String

  If (offLine = False) Then
    uniMsg = MLSupport.GSS("frm_secmodel", "statMsg1", "Performing Secondary model equation")
    Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Performing Secondary model equation", uniMsg)
  End If
  
  On Error GoTo probsecx
  nt = Trim(txtnvar.Text)
  
  For zz = 2 To nt + 1
    ss_cprop.Row = 1
    ss_cprop.Col = zz
    
    Select Case (Trim(ss_cprop.Text))
      Case "C"
        ss_cprop.Row = 2
        tempstring = ss_cprop.Text
        ss_cprop.Row = 3
        ss_cprop.Text = tempstring
      
      Case "F"
        ss_cprop.Row = 2
        tempstring = ss_cprop.Text
        ss_cprop.Row = 3
        ss_cprop.Text = tempstring
   
      Case "B"
        ss_cprop.Row = 2
        tempstring = ss_cprop.Text
        ss_cprop.Row = 3
        ss_cprop.Text = tempstring
          
      Case "P"    ' need to deal with calc terms and properties
        ss_cprop.Row = 2
        tempstring = Trim(ss_cprop.Text)
      
        For nn = 1 To unity_main.fpspread_pred.MaxRows
          unity_main.fpspread_pred.Col = 1
          unity_main.fpspread_pred.Row = nn
          tempstring2 = Trim(unity_main.fpspread_pred.Text)
        
          If (Trim(tempstring) = Trim(tempstring2)) Then
            
            If (offLine = True) Then
              unity_main.fpspread_pred.Col = 2
              tempstring = Trim(unity_main.fpspread_pred.Text)
              
              If (tempstring = "") Then
                tempstring = "0"
              End If
            Else
              tempstring = unity_main.preds.List(nn - 1)
            End If
            
            ss_cprop.Row = 3
            ss_cprop.Text = tempstring
            GoTo NEXT_VAR
          End If
        Next nn
    End Select
NEXT_VAR:
  Next zz

  unity_main.calc_string = ""
  frm_secmodel.ss_cprop.Row = 3
  
  For zz = 2 To nt + 1
    frm_secmodel.ss_cprop.Col = zz
    unity_main.calc_string = Trim(unity_main.calc_string) & Trim(frm_secmodel.ss_cprop.Text)
  Next zz

  frm_secmodel.Text1 = unity_main.calc_string
  frmParse2.txtExpression.Text = Trim(unity_main.calc_string)
  Call frmParse2.runparse
  Text6.Text = unity_main.sec_value
  
  If (offLine = True) Then
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_secmodel.statMsg1", "For equation: %1, the result was %2", txt_eqn.Text, frm_secmodel.Text6.Text)
  End If
  
  Exit Sub

probsecx:
  errMsg = (unity_main.fullmodelname & " file has equation problem. " & Error$)
  uniMsg = MLSupport.GGS_Params("frm_secmodel.errMsg1", "%1 file has equation problem. %2", unity_main.fullmodelname, Error$)
  
  If (offLine = True) Then
    CWrap.ShowMessageBoxW uniMsg, vbCritical
  Else
    Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  End If
  
  unity_main.pukedonpred = True
End Sub

Sub savesec()
  Dim fileName As String
  Dim fileExt As String
  Dim fToSave As String
  Dim nameLen As Integer
  Dim tmpText As String
  Dim x As Integer
  Dim uniFile As New clsUniFile
  Dim errMsg As String
  Dim uniMsg As String
  
  If (lst_type.ListCount = 0) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_secmodel", "errMsg3", "No equation defined"), vbCritical
    Exit Sub
  End If
  
  If (txt_mname2.Text = "") Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_secmodel", "errMsg4", "No model file name defined"), vbCritical
    Exit Sub
  End If
  
  ' Check model file name extension
  fileName = Trim(txt_mname2.Text)
  fileExt = ("." & LCase(uniFile.st_FileExt(fileName)))
    
  If (fileExt <> SEC_MODEL_FILE_EXT) Then
    fileName = uniFile.st_FileNameNoExt(fileName) & SEC_MODEL_FILE_EXT
  End If
  
  fToSave = (MODELS_DIR & fileName)
  
  If (uniFile.st_FileExist(fToSave) = True) Then
    uniFile.st_SetFileAttr (fToSave), vbNormal
  End If
  
  txtnvar.Text = lst_type.ListCount

  If (uniFile.OpenFileWrite(fToSave) = True) Then
    On Error GoTo FILE_ERROR
    uniFile.WriteBOM fe_UTF16LE
    uniFile.WriteUnicodeLine txt_cprop.Text
    uniFile.WriteUnicodeLine txt_modinfo.Text
    uniFile.WriteUnicodeLine lst_type.ListCount
  
    For x = 0 To lst_type.ListCount - 1
      uniFile.WriteUnicodeLine lst_what.List(x) & "," & lst_type.List(x)
    Next x

    uniFile.Flush
    unity_main.errorstring = ("User saved new settings for model file: " & fToSave)
    unity_main.write_error (LOG_DBG_LEVEL1)
    Unload frm_secmodel
  Else
FILE_ERROR:
    errMsg = (fToSave & " file write error. " & Error$)
    unity_main.errorstring = errMsg
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", fToSave, Error$)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  End If
  
  uniFile.CloseFile
End Sub

Sub dosecpred()  '3/21
  Dim tempstring As String
  Dim nVar As Integer
  Dim zz As Integer
  Dim lineCnt As Integer
  Dim varStr As Variant
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim rc As Boolean
  Dim errMsg As String
  Dim uniMsg As String

  uniMsg = MLSupport.GGS_Params("frm_secmodel.statMsg2", "Loading Secondary model: %1", unity_main.modlname)
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Loading Secondary model: " & unity_main.modlname), uniMsg)
  
  If (uniFile.OpenFileRead(unity_main.fullmodelname) = True) Then
    On Error GoTo FILE_ERROR
    fEncoding = uniFile.ReadBOM
    lineCnt = lineCnt + 1

    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(tempstring)
    Else
      rc = uniFile.ReadUnicodeLine(tempstring)
    End If
      
    If (rc = False) Then GoTo FILE_ERROR
  
    txt_cprop.Text = tempstring
  
    lineCnt = lineCnt + 1
    
    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(tempstring)
    Else
      rc = uniFile.ReadUnicodeLine(tempstring)
    End If
      
    If (rc = False) Then GoTo FILE_ERROR
  
    txt_modinfo.Text = tempstring
    uniMsg = MLSupport.GGS_Params("frm_secmodel.statMsg3", "Model info: %1", tempstring)
    Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Model info: " & tempstring), uniMsg)
  
    lineCnt = lineCnt + 1
    
    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(tempstring)
    Else
      rc = uniFile.ReadUnicodeLine(tempstring)
    End If
      
    If (rc = False) Then GoTo FILE_ERROR
  
    txtnvar.Text = tempstring
    nVar = tempstring
  
    For zz = 1 To nVar
      lineCnt = lineCnt + 1
      
      If (fEncoding = fe_ANSI) Then
        rc = uniFile.ReadAnsiLine(tempstring)
      Else
        rc = uniFile.ReadUnicodeLine(tempstring)
      End If
      
      If (rc = False) Then GoTo FILE_ERROR
    
      varStr = Split(tempstring, ",")
      frm_secmodel.ss_cprop.Col = zz + 1
      frm_secmodel.ss_cprop.Row = 1
      frm_secmodel.ss_cprop.Text = varStr(1)
    
      frm_secmodel.ss_cprop.Row = 2
      frm_secmodel.ss_cprop.Text = varStr(0)
    Next zz

    tempstring = ""
    frm_secmodel.ss_cprop.Row = 2
  
    For zz = 2 To nVar + 1
      frm_secmodel.ss_cprop.Col = zz
      tempstring = (tempstring & Trim(frm_secmodel.ss_cprop.Text))
    Next zz

    txt_eqn.Text = tempstring
    Call dosecpred2(False)
  Else
FILE_ERROR:
    If (lineCnt = 0) Then
      errMsg = unity_main.fullmodelname & " file open error." & Error$
      uniMsg = MLSupport.GGS_Params("fileErrMsg2", "%1 file open error. %2", unity_main.fullmodelname, Error$)
    Else
      errMsg = unity_main.fullmodelname & " file has error on line " & CStr(lineCnt) & ". " & Error$
      uniMsg = MLSupport.GGS_Params("fileErrMsg3", "%1 file has error on line %2. %3", unity_main.fullmodelname, CStr(lineCnt), Error$)
    End If
  
    Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    unity_main.pukedonpred = True
  End If
  
  uniFile.CloseFile
End Sub

Private Sub cmd_save_Click()
  
  unity_main.errorstring = "Secondary Model Design screen Save Changes button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Call savesec
End Sub

Private Sub Form_Load()
  
  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
  Call loadfunctions
  Call loadprops
End Sub

Sub loadprops()
  Dim zz As Integer
  Dim tempstring As String
  
  frmedmod.grid_models.Col = 1
  
  For zz = 1 To frmedmod.grid_models.MaxRows
    frmedmod.grid_models.Row = zz
    tempstring = frmedmod.grid_models.Text
    
    If (Trim(tempstring) <> "") Then
      frm_secmodel.lst_props.AddItem (tempstring)
    End If
  Next zz
  
  Call frmedmod.fixthesize
End Sub

Sub loadfunctions()
  
  frm_secmodel.lst_functs.AddItem (" + ")
  frm_secmodel.lst_functs.AddItem (" - ")
  frm_secmodel.lst_functs.AddItem (" / ")
  frm_secmodel.lst_functs.AddItem (" * ")
  frm_secmodel.lst_functs.AddItem (" ( ")
  frm_secmodel.lst_functs.AddItem (" ) ")
End Sub

Private Sub lst_functs_Click()
  
  If (Trim(frm_secmodel.lst_functs.Text) = "(") Then
    frm_secmodel.lst_type.AddItem ("B")
  ElseIf (Trim(frm_secmodel.lst_functs.Text) = ")") Then
    frm_secmodel.lst_type.AddItem ("B")
  Else
    frm_secmodel.lst_type.AddItem ("F")
  End If
  
  frm_secmodel.lst_what.AddItem (lst_functs.Text)
  frm_secmodel.txt_eqn.Text = frm_secmodel.txt_eqn.Text & frm_secmodel.lst_functs.Text
End Sub

Private Sub lst_props_Click()
  
  frm_secmodel.lst_type.AddItem ("P")
  frm_secmodel.lst_what.AddItem (lst_props.Text)
  frm_secmodel.txt_eqn.Text = frm_secmodel.txt_eqn.Text & frm_secmodel.lst_props.Text
End Sub

Private Sub txt_cprop_DblClick(Button As Integer)
  
  unity_main.formfrom = 5
  unity_main.varfrom = 2
  frm_kybd.lbl_kybd.Caption = frm_secmodel.Label6.Caption
  frm_kybd.txt_kybd.Text = txt_cprop.Text
  frm_kybd.Show 1
End Sub

Private Sub txt_mname2_DblClick(Button As Integer)
  
  unity_main.formfrom = 5
  unity_main.varfrom = 3
  frm_kybd.lbl_kybd.Caption = frm_secmodel.Label2.Caption
  frm_kybd.txt_kybd.Text = txt_mname2.Text
  frm_kybd.Show 1
End Sub

Private Sub txt_modinfo_DblClick(Button As Integer)
  
  unity_main.formfrom = 5
  unity_main.varfrom = 1
  frm_kybd.lbl_kybd.Caption = frm_secmodel.Label1.Caption
  frm_kybd.txt_kybd.Text = txt_modinfo.Text
  frm_kybd.Show 1
End Sub

Private Sub txt_value_DblClick(Button As Integer)
  
  unity_main.formfrom = 5
  unity_main.varfrom = 1
  frm_numpad.lbl_num.Caption = frm_secmodel.Label3.Caption
  frm_numpad.txt_num.Text = txt_value.Text
  frm_numpad.Show 1
End Sub








