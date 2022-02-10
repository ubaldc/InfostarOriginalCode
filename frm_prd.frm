VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "resize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_prd 
   Caption         =   "UCal PRD Model Property Configuration"
   ClientHeight    =   9255
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   12975
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
   Icon            =   "frm_prd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   12975
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8880
      Picture         =   "frm_prd.frx":0442
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   12120
      Picture         =   "frm_prd.frx":0884
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
   End
   Begin HexUniControls.ctlUniTextBoxXP Text1 
      Height          =   375
      Left            =   12120
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   8640
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_prd.frx":0CC6
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
      Tip             =   "frm_prd.frx":0CF0
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":0D10
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_propname 
      Height          =   405
      Left            =   9120
      TabIndex        =   2
      Top             =   600
      Width           =   2895
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_prd.frx":0D2C
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
      Tip             =   "frm_prd.frx":0D4C
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":0D6C
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   645
      Left            =   10080
      TabIndex        =   17
      Top             =   5040
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
      Caption         =   "frm_prd.frx":0D88
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_prd.frx":0DC0
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":0DE0
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Height          =   645
      Left            =   10080
      TabIndex        =   0
      Top             =   5760
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
      Caption         =   "frm_prd.frx":0DFC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_prd.frx":0E28
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":0E48
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   7
      Left            =   10080
      TabIndex        =   11
      Top             =   7200
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_prd.frx":0E64
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
      Tip             =   "frm_prd.frx":0E84
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":0EA4
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   9
      Left            =   10080
      TabIndex        =   12
      Top             =   8640
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_prd.frx":0EC0
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
      Tip             =   "frm_prd.frx":0EE0
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":0F00
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   11
      Left            =   10080
      TabIndex        =   13
      Top             =   7680
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_prd.frx":0F1C
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
      Tip             =   "frm_prd.frx":0F3C
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":0F5C
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   13
      Left            =   10080
      TabIndex        =   14
      Top             =   8160
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_prd.frx":0F78
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
      Tip             =   "frm_prd.frx":0F98
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":0FB8
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11400
      Picture         =   "frm_prd.frx":0FD4
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7200
      Width           =   495
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11400
      Picture         =   "frm_prd.frx":1416
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   8640
      Width           =   495
   End
   Begin VB.PictureBox Picture12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11400
      Picture         =   "frm_prd.frx":1858
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7680
      Width           =   495
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11400
      Picture         =   "frm_prd.frx":1C9A
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8160
      Width           =   495
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   6
      Left            =   4005
      TabIndex        =   7
      Top             =   7200
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_prd.frx":20DC
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
      Tip             =   "frm_prd.frx":20FC
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":211C
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   8
      Left            =   4005
      TabIndex        =   8
      Top             =   8640
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_prd.frx":2138
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
      Tip             =   "frm_prd.frx":2158
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":2178
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   10
      Left            =   4005
      TabIndex        =   9
      Top             =   7680
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_prd.frx":2194
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
      Tip             =   "frm_prd.frx":21B4
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":21D4
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   12
      Left            =   4005
      TabIndex        =   10
      Top             =   8160
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_prd.frx":21F0
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
      Tip             =   "frm_prd.frx":2210
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":2230
   End
   Begin HexUniControls.ctlUniTextBoxXP txtmrow 
      Height          =   495
      Left            =   15120
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_prd.frx":224C
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
      Tip             =   "frm_prd.frx":226C
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":228C
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      Picture         =   "frm_prd.frx":22A8
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7200
      Width           =   495
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      Picture         =   "frm_prd.frx":26EA
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   8640
      Width           =   495
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      Picture         =   "frm_prd.frx":2B2C
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7680
      Width           =   495
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      Picture         =   "frm_prd.frx":2F6E
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   8160
      Width           =   495
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   3
      Left            =   4005
      TabIndex        =   4
      Top             =   5040
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_prd.frx":33B0
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
      Tip             =   "frm_prd.frx":33D0
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":33F0
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   4
      Left            =   4005
      TabIndex        =   5
      Top             =   5520
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_prd.frx":340C
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
      Tip             =   "frm_prd.frx":342C
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":344C
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   5
      Left            =   4005
      TabIndex        =   6
      Top             =   6000
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_prd.frx":3468
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
      Tip             =   "frm_prd.frx":3488
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":34A8
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      Picture         =   "frm_prd.frx":34C4
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5040
      Width           =   495
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      Picture         =   "frm_prd.frx":3906
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5520
      Width           =   495
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      Picture         =   "frm_prd.frx":3D48
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6000
      Width           =   495
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   1440
      Top             =   120
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9255
      FormDesignWidth =   12975
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_browsePRD 
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2295
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
      Caption         =   "frm_prd.frx":418A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_prd.frx":41B6
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":41D6
   End
   Begin HexUniControls.ctlUniListBoxXP List1 
      Height          =   2025
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3572
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
      TrapTab         =   0   'False
      Tip             =   "frm_prd.frx":41F2
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   -1  'True
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4212
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin HexUniControls.ctlUniLabel Label16 
      Height          =   405
      Left            =   9480
      Top             =   120
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":422E
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":425E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":427E
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   7
      Left            =   6795
      Top             =   7200
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":429A
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":42D8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":42F8
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   9
      Left            =   6795
      Top             =   8640
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":4314
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":435A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":437A
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   11
      Left            =   6795
      Top             =   7680
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":4396
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":43E4
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4404
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   13
      Left            =   6795
      Top             =   8175
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":4420
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":4470
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4490
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   6
      Left            =   705
      Top             =   7200
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":44AC
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":44EA
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":450A
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   8
      Left            =   705
      Top             =   8640
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":4526
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":456C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":458C
   End
   Begin HexUniControls.ctlUniLabel Label17 
      Height          =   375
      Left            =   4080
      Top             =   6600
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":45A8
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":45EE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":460E
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   10
      Left            =   705
      Top             =   7680
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":462A
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":4678
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4698
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   12
      Left            =   705
      Top             =   8175
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":46B4
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":4704
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4724
   End
   Begin HexUniControls.ctlUniLabel Label14 
      Height          =   375
      Left            =   15525
      Top             =   9120
      Visible         =   0   'False
      Width           =   375
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
      Caption         =   "frm_prd.frx":4740
      BackColor       =   -2147483633
      ForeColor       =   12582912
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":4762
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4782
   End
   Begin HexUniControls.ctlUniLabel Label13 
      Height          =   375
      Left            =   16965
      Top             =   9120
      Visible         =   0   'False
      Width           =   375
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
      Caption         =   "frm_prd.frx":479E
      BackColor       =   -2147483633
      ForeColor       =   12582912
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":47C0
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":47E0
   End
   Begin HexUniControls.ctlUniLabel Label12 
      Height          =   375
      Left            =   18405
      Top             =   9120
      Visible         =   0   'False
      Width           =   375
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
      Caption         =   "frm_prd.frx":47FC
      BackColor       =   -2147483633
      ForeColor       =   12582912
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":481E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":483E
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   3
      Left            =   705
      Top             =   5040
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":485A
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":488C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":48AC
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   4
      Left            =   705
      Top             =   5520
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":48C8
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":48F2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4912
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   5
      Left            =   705
      Top             =   6000
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":492E
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":4974
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4994
   End
   Begin HexUniControls.ctlUniLabel Label4 
      Height          =   375
      Left            =   2760
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":49B0
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":4A00
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4A20
   End
   Begin HexUniControls.ctlUniLabel Label8 
      Height          =   195
      Left            =   15600
      Top             =   10320
      Visible         =   0   'False
      Width           =   270
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
      Caption         =   "frm_prd.frx":4A3C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   -1  'True
      Tip             =   "frm_prd.frx":4A64
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4A84
   End
   Begin HexUniControls.ctlUniLabel Label6 
      Height          =   255
      Left            =   14640
      Top             =   10680
      Visible         =   0   'False
      Width           =   1095
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
      Caption         =   "frm_prd.frx":4AA0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":4AD8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4AF8
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   0
      Left            =   240
      Top             =   1320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":4B14
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":4B5A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4B7A
   End
   Begin HexUniControls.ctlUniTextBoxXP PredSpecNum 
      Height          =   285
      Left            =   7200
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   8760
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_prd.frx":4B96
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
      Tip             =   "frm_prd.frx":4BB8
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4BD8
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_browseSTF 
      Height          =   615
      Left            =   240
      TabIndex        =   31
      Top             =   4200
      Width           =   2295
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
      Caption         =   "frm_prd.frx":4BF4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_prd.frx":4C20
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4C40
   End
   Begin HexUniControls.ctlUniLabel Label2 
      Height          =   375
      Left            =   2640
      Top             =   3840
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":4C5C
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":4CB6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4CD6
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   1
      Left            =   240
      Top             =   1320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":4CF2
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":4D3C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4D5C
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_import 
      Height          =   645
      Left            =   10080
      TabIndex        =   32
      Top             =   5040
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
      Caption         =   "frm_prd.frx":4D78
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_prd.frx":4DBE
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4DDE
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_selectAll 
      Height          =   645
      Left            =   7080
      TabIndex        =   33
      Top             =   1920
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
      Caption         =   "frm_prd.frx":4DFA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_prd.frx":4E2E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4E4E
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_unselectAll 
      Cancel          =   -1  'True
      Height          =   645
      Left            =   7080
      TabIndex        =   34
      Top             =   2760
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
      Caption         =   "frm_prd.frx":4E6A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_prd.frx":4EA2
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4EC2
   End
   Begin HexUniControls.ctlUniLabel lbl_prdName 
      Height          =   405
      Left            =   2760
      Top             =   600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_prd.frx":4EDE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_prd.frx":4EFE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4F1E
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_stfName 
      Height          =   405
      Left            =   2640
      TabIndex        =   35
      Top             =   4320
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   714
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_prd.frx":4F3A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
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
      Tip             =   "frm_prd.frx":4F5A
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_prd.frx":4F7A
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   0
      Top             =   0
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
      Left            =   720
      Top             =   120
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_prd.frx":4F96
   End
End
Attribute VB_Name = "frm_prd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_prdFileName As String
Public m_stfFileName As String
  
Public m_predVal As Double
Public m_globalVal As Double
Public m_ndVal As Double
Public m_pFExpCodes As Long
  
Private m_numConstituents As Long
Private m_constituentNames() As String

Public Sub set_constituent_list_index(constituentName As String)
  Dim ii As Long

  If (build_full_constituent_list = 0) Then
    For ii = 0 To m_numConstituents - 1
      If (m_constituentNames(ii) = constituentName) Then
        List1.ListIndex = ii
        Exit Sub
      End If
    Next ii
  End If
End Sub

Public Function build_full_constituent_list() As Long
  Dim rc As Long
  Dim ii As Long
    
  ' Get constituents names
  rc = get_constituent_names
    
  List1.Clear
    
  If (rc = 0) Then
    ' Build list of constituents
    For ii = 0 To m_numConstituents - 1
      List1.AddItem m_constituentNames(ii)
    Next ii
  End If
  
  build_full_constituent_list = rc
End Function

Public Function build_partial_constituent_list() As Long
  Dim rc As Long
  Dim ii As Long
  Dim jj As Long
  Dim notListed As Boolean
    
  ' Get constituents names
  rc = get_constituent_names
    
  If (rc = 0) Then
    List1.Clear
      
    ' Build list of constituents
    For ii = 0 To m_numConstituents - 1
      notListed = True
      
      ' Check if constituent already imported
      For jj = 0 To frmedmod.lst_prdConstituentNames.ListCount - 1
        If (m_constituentNames(ii) = frmedmod.lst_prdConstituentNames.List(jj)) Then
          notListed = False
          Exit For
        End If
      Next jj
      
      If (notListed = True) Then
        List1.AddItem m_constituentNames(ii)
      End If
    Next ii
  End If
  
  build_partial_constituent_list = rc
End Function

Public Function get_constituent_names() As Long
  Dim rc As Long
  Dim errMsg As String
  Dim uniMsg As String
    
  ' Get # of constituents
  On Error GoTo OBJECT_ERROR
  rc = PRDObject.getNumConstituents(frm_prd.m_prdFileName, m_numConstituents)
    
  If (rc = -200) Then
  ' since mahendra's encryption code generates its own errors, we are just exiting. i dont agree with this.
    'errMsg = frm_prd.m_prdFileName & " UCal PRD model license has expired.: " & rc
    'uniMsg = MLSupport.GGS_Params("frm_prd.errMsg", "%1 UCal PRD model file error: UCal PRD model license has expired. %2", frm_prd.m_prdFileName, CStr(rc))
    'Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
    'CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg", "%1. Please contact your supervisor!", uniMsg), vbCritical
    'get_constituent_names = -200
    'm_numConstituents = 0
    'frm_prd.lbl_prdName.Caption = ""
    Exit Function
  End If
  
  If (rc = 0) Then
    ' Get constituent names
    ReDim m_constituentNames(m_numConstituents - 1)
    rc = PRDObject.getConstituentNames(frm_prd.m_prdFileName, 0, m_numConstituents, m_constituentNames)
    
    If (rc = 0) Then GoTo LEAVE_RTN
  End If
    frm_prd.lbl_prdName.Caption = ""
  errMsg = frm_prd.m_prdFileName & " UCal PRD model file error: " & rc
  uniMsg = MLSupport.GGS_Params("frm_prd.errMsg1", "%1 UCal PRD model file error: %2", frm_prd.m_prdFileName, CStr(rc))
  Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  
LEAVE_RTN:
  get_constituent_names = rc
  
  Exit Function
  
OBJECT_ERROR:
  unity_main.errorstring = "checktablespread:Unity PRDComponent.dll component not installed or registered"
  unity_main.write_error (LOG_DBG_LEVEL1)
  uniMsg = MLSupport.GGS_Params("errMsg7", "Unity %1 component not installed or registered", "PRDComponent.dll")
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  get_constituent_names = False
End Function

Public Sub prd_predict()
  Dim rc As Long
  Dim predVal As Double
  Dim globalVal As Double
  Dim ndVal As Double
  Dim pFExpCodes As Long
  Dim errMsg As String
  Dim uniMsg As String
  
  On Error GoTo OBJECT_ERROR
    
#If ABBFT Then
  ' Perform PRD model prediction
  rc = PRDObject.predict(unity_main.modlname, unity_main.m_stfFileName, unity_main.m_prdConstituent, unity_main.m_sysSerialNum, unity_main.m_mb3000.m_startWavenum, unity_main.m_mb3000.m_endWavenum, unity_main.m_mb3000.m_waveNumIncr, ProdAbsYVals, predVal, globalVal, ndVal, pFExpCodes)
#Else
  ' Check if spectrum data treated
  If (unity_main.m_enableTreatment = True) Then
    ' Perform PRD model prediction
    rc = PRDObject.predict(unity_main.modlname, unity_main.m_stfFileName, unity_main.m_prdConstituent, unity_main.m_sysSerialNum, unity_main.m_smplStartWvln, unity_main.m_smplEndWvln, MS11CfgData.wvlnIncr, ProdTreatAbsYVals, predVal, globalVal, ndVal, pFExpCodes)
  Else
    ' Perform PRD model prediction
    rc = PRDObject.predict(unity_main.modlname, unity_main.m_stfFileName, unity_main.m_prdConstituent, unity_main.m_sysSerialNum, unity_main.m_smplStartWvln, unity_main.m_smplEndWvln, MS11CfgData.wvlnIncr, ProdAbsYVals, predVal, globalVal, ndVal, pFExpCodes)
  End If
#End If

  If (rc = 0) Then
    ' Save prediction values
    frm_prd.m_predVal = predVal
    frm_prd.m_globalVal = globalVal
    frm_prd.m_ndVal = ndVal
    frm_prd.m_pFExpCodes = pFExpCodes
    Exit Sub
  End If

PREDICT_ERR:
  ' Check if no transfer file defined
  If (unity_main.m_stfFileName = "") Then
    errMsg = (unity_main.modlname & " UCal PRD model " & unity_main.m_prdConstituent & " constituent had prediction error: " & rc)
    uniMsg = MLSupport.GGS_Params("frm_prd.errMsg2", "%1 UCal PRD model %2 constituent had prediction error: %3", unity_main.modlname, unity_main.m_prdConstituent, CStr(rc))
  Else
    errMsg = (unity_main.modlname & " UCal PRD model " & unity_main.m_prdConstituent & " constituent " & unity_main.m_stfFileName & " STF file had prediction error: " & rc)
    uniMsg = MLSupport.GGS_Params("frm_prd.errMsg3", "%1 UCal PRD model %2 constituent %3 STF file had prediction error: %4", unity_main.modlname, unity_main.m_prdConstituent, unity_main.m_stfFileName, CStr(rc))
  End If
  
LEAVE_RTN:
  Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
  CWrap.ShowMessageBoxW uniMsg, vbCritical
  unity_main.pukedonpred = True
  Exit Sub
  
OBJECT_ERROR:
  errMsg = "Unity PRDComponent.dll component not installed or registered"
  uniMsg = MLSupport.GGS_Params("errMsg7", "Unity %1 component not installed or registered", "PRDComponent.dll")
  GoTo LEAVE_RTN
End Sub

Private Sub cmd_browsePRD_Click()
  Dim dialog As New clsCommDialogs
  Dim fileDir As String
  Dim sFilter As String
  Dim dlgTitle As String
  Dim fileName As String

  unity_main.errorstring = "UCal PRD Model Property Configuration screen Browse PRD button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
    
  dialog.InitDialogs
  fileDir = MODELS_DIR
  sFilter = ("UCal PRD Model Files (*" & PRD_MODEL_FILE_EXT & ")" & Chr(0) & "*" & PRD_MODEL_FILE_EXT & Chr(0))
  dlgTitle = MLSupport.GSS("frm_prd", "dlgTitle", "Select Model File")
  fileName = dialog.ShowOpen(Me.hWnd, sFilter, dlgTitle, OFN_FILEMUSTEXIST + OFN_HIDEREADONLY + OFN_NOCHANGEDIR, fileDir)
 
  If (fileName <> "") Then
    frm_prd.m_prdFileName = fileName
    lbl_prdName.Caption = frm_prd.m_prdFileName
    build_full_constituent_list
  End If
End Sub

Private Sub cmd_browseSTF_Click()
  Dim dialog As New clsCommDialogs
  Dim fileDir As String
  Dim sFilter As String
  Dim dlgTitle As String
  Dim fileName As String
  Dim ii As Integer

  unity_main.errorstring = "UCal PRD Model Property Configuration screen Browse STF button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
    
  dialog.InitDialogs
  fileDir = MODELS_DIR
  sFilter = ("Transfer Files (*" & STF_FILE_EXT & ")" & Chr(0) & "*" & STF_FILE_EXT & Chr(0))
  dlgTitle = MLSupport.GSS("frm_prd", "dlgTitle2", "Select PRD Transfer File")
  fileName = dialog.ShowOpen(Me.hWnd, sFilter, dlgTitle, OFN_FILEMUSTEXIST + OFN_HIDEREADONLY + OFN_NOCHANGEDIR, fileDir)
 
  If (fileName <> "") Then
    frm_prd.m_stfFileName = fileName
    txt_stfName.Text = frm_prd.m_stfFileName
  End If
End Sub

Private Sub cmd_cancel_Click()
  
  unity_main.errorstring = "UCal PRD Model Property Configuration screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  If (frmedmod.m_addProp = True) Then
    frmedmod.m_addProp = False
    frmedmod.grid_models.MaxRows = frmedmod.grid_models.MaxRows - 1
  End If
  
  Unload frm_prd
End Sub

Private Sub cmd_import_Click()
  Dim Row As Integer
  Dim ii As Integer
  Dim jj As Integer

  unity_main.errorstring = "UCal PRD Model Property Configuration screen Import Constituents button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  If (List1.SelCount = 0) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_prd", "errMsg2", "You must select a constituent"), vbCritical
    frm_prd.List1.SetFocus
    Exit Sub
  End If

  Row = txtmrow.Text
  
  If ((List1.SelCount + Row) > MAX_NUM_PROPS) Then
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_prd.errMsg4", "The number of selected constituents (%1) plus existing defined properties (%2) exceed the maximum number of properties supported (%3)", CStr(List1.SelCount), CStr(Row), CStr(MAX_NUM_PROPS)), vbCritical
    frm_prd.List1.SetFocus
    Exit Sub
  End If

  For ii = 3 To 13
    If (Trim(frm_prd.txt_modvar(ii).Text) = "") Or (IsNumeric(frm_prd.txt_modvar(ii).Text) = False) Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_prd", "errMsg3", "Please enter a valid number"), vbCritical
      frm_prd.txt_modvar(ii).SetFocus
      Exit Sub
    End If
  Next ii

  frmedmod.grid_models.MaxRows = List1.SelCount + Row
  
  For ii = 0 To List1.ListCount - 1
    If (List1.Selected(ii) = True) Then
      Row = Row + 1
      frmedmod.grid_models.Row = Row
      frmedmod.grid_models.Col = 1
      frmedmod.grid_models.Text = List1.List(ii)
      frmedmod.grid_models.Col = 2
      frmedmod.grid_models.Text = Trim(frm_prd.lbl_prdName.Caption)
      frmedmod.m_prdFileName = frmedmod.grid_models.Text
      frmedmod.grid_models.Col = 3
      frmedmod.grid_models.Text = List1.List(ii)
      frmedmod.lst_prdConstituentNames.AddItem List1.List(ii)
  
      For jj = 4 To frmedmod.grid_models.MaxCols - 1
       frmedmod.grid_models.Col = jj
       frmedmod.grid_models.Text = frm_prd.txt_modvar(jj - 1).Text
      Next jj

      frmedmod.grid_models.Col = frmedmod.grid_models.MaxCols
      frmedmod.grid_models.Text = Trim(frm_prd.txt_stfName.Text)
      frmedmod.m_stfFileName = frmedmod.grid_models.Text
    End If
  Next ii

  frmedmod.m_prdModelType = True
  
  frmedmod.numprops.Text = frmedmod.grid_models.MaxRows
  frmedmod.fillproplist
  Unload frm_prd
End Sub

Private Sub cmd_save_Click()
  Dim ii As Integer
  
  unity_main.errorstring = "UCal PRD Model Property Configuration screen Save Changes button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  If (Trim(txt_propname.Text) = "") Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_prd", "errMsg1", "You must enter a property name"), vbCritical
    frm_prd.txt_propname.SetFocus
    Exit Sub
  End If
  
  If (List1.ListIndex = -1) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_prd", "errMsg2", "You must select a constituent"), vbCritical
    frm_prd.List1.SetFocus
    Exit Sub
  End If
  
  For ii = 3 To 13
    If (Trim(frm_prd.txt_modvar(ii).Text) = "") Or (IsNumeric(frm_prd.txt_modvar(ii).Text) = False) Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_prd", "errMsg3", "Please enter a valid number"), vbCritical
      frm_prd.txt_modvar(ii).SetFocus
      Exit Sub
    End If
  Next ii

  frmedmod.grid_models.Row = CInt(frm_prd.txtmrow.Text)
  frmedmod.grid_models.Col = 1
  frmedmod.grid_models.Text = frm_prd.txt_propname.Text
  frmedmod.grid_models.Col = 2
  frmedmod.grid_models.Text = Trim(frm_prd.lbl_prdName.Caption)
  frmedmod.grid_models.Col = 3
  frmedmod.grid_models.Text = List1.List(List1.ListIndex)

  For ii = 4 To frmedmod.grid_models.MaxCols - 1
    frmedmod.grid_models.Col = ii
    frmedmod.grid_models.Text = frm_prd.txt_modvar(ii - 1).Text
  Next ii

  frmedmod.grid_models.Col = frmedmod.grid_models.MaxCols
  frmedmod.grid_models.Text = Trim(frm_prd.txt_stfName.Text)

  If (frmedmod.m_addProp = True) Then
    frmedmod.m_addProp = False
    frmedmod.numprops.Text = frmedmod.numprops.Text + 1
  End If
  
  frmedmod.fillproplist
  Unload frm_prd
End Sub

Private Sub cmd_selectAll_Click()
  Dim ii As Integer
  
  For ii = 0 To List1.ListCount - 1
    List1.Selected(ii) = True
  Next ii
End Sub

Private Sub cmd_unselectAll_Click()
  Dim ii As Integer
  
  For ii = 0 To List1.ListCount - 1
    List1.Selected(ii) = False
  Next ii
End Sub

Private Sub Form_Load()
    
  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
  Label8.Caption = ""
End Sub











Private Sub Picture1_Click()
  
  unity_main.formfrom = 17
  unity_main.varfrom = 1
  frm_kybd.lbl_kybd.Caption = frm_prd.Label16.Caption
  frm_kybd.txt_kybd.Text = Trim(frm_prd.txt_propname.Text)
  frm_kybd.Show 1
End Sub

Private Sub Picture2_Click()

  unity_main.formfrom = 17
  unity_main.varfrom = 2
  frm_kybd.lbl_kybd.Caption = frm_prd.Label2.Caption
  frm_kybd.txt_kybd.Text = Trim(frm_prd.txt_stfName.Text)
  frm_kybd.Show 1
End Sub

Private Sub Picture4_Click()
  
  unity_main.formfrom = 17
  unity_main.varfrom = 3
  frm_numpad.lbl_num.Caption = frm_prd.Label1(3).Caption
  frm_numpad.txt_num.Text = txt_modvar(3).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture5_Click()
  
  unity_main.formfrom = 17
  unity_main.varfrom = 4
  frm_numpad.lbl_num.Caption = frm_prd.Label1(4).Caption
  frm_numpad.txt_num.Text = txt_modvar(4).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture6_Click()
  
  unity_main.formfrom = 17
  unity_main.varfrom = 5
  frm_numpad.lbl_num.Caption = frm_prd.Label1(5).Caption
  frm_numpad.txt_num.Text = txt_modvar(5).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture7_Click()
  
  unity_main.formfrom = 17
  unity_main.varfrom = 6
  frm_numpad.lbl_num.Caption = frm_prd.Label1(6).Caption
  frm_numpad.txt_num.Text = txt_modvar(6).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture8_Click()
  
  unity_main.formfrom = 17
  unity_main.varfrom = 7
  frm_numpad.lbl_num.Caption = frm_prd.Label1(7).Caption
  frm_numpad.txt_num.Text = txt_modvar(7).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture9_Click()
  
  unity_main.formfrom = 17
  unity_main.varfrom = 8
  frm_numpad.lbl_num.Caption = frm_prd.Label1(8).Caption
  frm_numpad.txt_num.Text = txt_modvar(8).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture10_Click()
  
  unity_main.formfrom = 17
  unity_main.varfrom = 9
  frm_numpad.lbl_num.Caption = frm_prd.Label1(9).Caption
  frm_numpad.txt_num.Text = txt_modvar(9).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture11_Click()
  
  unity_main.formfrom = 17
  unity_main.varfrom = 10
  frm_numpad.lbl_num.Caption = frm_prd.Label1(10).Caption
  frm_numpad.txt_num.Text = txt_modvar(10).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture12_Click()
  
  unity_main.formfrom = 17
  unity_main.varfrom = 11
  frm_numpad.lbl_num.Caption = frm_prd.Label1(11).Caption
  frm_numpad.txt_num.Text = txt_modvar(11).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture13_Click()
  
  unity_main.formfrom = 17
  unity_main.varfrom = 12
  frm_numpad.lbl_num.Caption = frm_prd.Label1(12).Caption
  frm_numpad.txt_num.Text = txt_modvar(12).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture14_Click()
  
  unity_main.formfrom = 17
  unity_main.varfrom = 13
  frm_numpad.lbl_num.Caption = frm_prd.Label1(13).Caption
  frm_numpad.txt_num.Text = txt_modvar(13).Text
  frm_numpad.Show 1
End Sub

Private Sub txt_propname_DblClick(Button As Integer)
  
  unity_main.formfrom = 17
  unity_main.varfrom = 1
  frm_kybd.txt_kybd.Text = Trim(frm_prd.txt_propname.Text)
  frm_kybd.Show 1
End Sub








