VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frmMain 
   Caption         =   "CalStar Model Property Configuration"
   ClientHeight    =   8940
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
   Icon            =   "frm_slogic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   12975
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   12120
      Picture         =   "frm_slogic.frx":0442
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   840
      Width           =   495
   End
   Begin HexUniControls.ctlUniTextBoxXP Text1 
      Height          =   375
      Left            =   11760
      TabIndex        =   17
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
      Text            =   "frm_slogic.frx":0884
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
      Tip             =   "frm_slogic.frx":08AE
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":08CE
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_propname 
      Height          =   400
      Left            =   9120
      TabIndex        =   2
      Top             =   810
      Width           =   2895
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_slogic.frx":08EA
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
      Tip             =   "frm_slogic.frx":090A
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":092A
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   650
      Left            =   10080
      TabIndex        =   18
      Top             =   4200
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
      Caption         =   "frm_slogic.frx":0946
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_slogic.frx":097E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":099E
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   10080
      TabIndex        =   0
      Top             =   4920
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
      Caption         =   "frm_slogic.frx":09BA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_slogic.frx":09E6
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":0A06
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   7
      Left            =   10080
      TabIndex        =   11
      Top             =   6240
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_slogic.frx":0A22
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
      Tip             =   "frm_slogic.frx":0A42
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":0A62
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   9
      Left            =   10080
      TabIndex        =   12
      Top             =   6840
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_slogic.frx":0A7E
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
      Tip             =   "frm_slogic.frx":0A9E
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":0ABE
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   11
      Left            =   10080
      TabIndex        =   13
      Top             =   7440
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_slogic.frx":0ADA
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
      Tip             =   "frm_slogic.frx":0AFA
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":0B1A
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   13
      Left            =   10080
      TabIndex        =   14
      Top             =   8040
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_slogic.frx":0B36
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
      Tip             =   "frm_slogic.frx":0B56
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":0B76
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11400
      Picture         =   "frm_slogic.frx":0B92
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6240
      Width           =   495
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11400
      Picture         =   "frm_slogic.frx":0FD4
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox Picture12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11400
      Picture         =   "frm_slogic.frx":1416
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7440
      Width           =   495
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11400
      Picture         =   "frm_slogic.frx":1858
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   8040
      Width           =   495
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   6
      Left            =   4005
      TabIndex        =   7
      Top             =   6240
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_slogic.frx":1C9A
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
      Tip             =   "frm_slogic.frx":1CBA
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":1CDA
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   8
      Left            =   4005
      TabIndex        =   8
      Top             =   6840
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_slogic.frx":1CF6
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
      Tip             =   "frm_slogic.frx":1D16
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":1D36
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   10
      Left            =   4005
      TabIndex        =   9
      Top             =   7440
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_slogic.frx":1D52
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
      Tip             =   "frm_slogic.frx":1D72
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":1D92
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   12
      Left            =   4005
      TabIndex        =   10
      Top             =   8040
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_slogic.frx":1DAE
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
      Tip             =   "frm_slogic.frx":1DCE
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":1DEE
   End
   Begin HexUniControls.ctlUniTextBoxXP txtmrow 
      Height          =   495
      Left            =   15120
      TabIndex        =   24
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
      Text            =   "frm_slogic.frx":1E0A
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
      Tip             =   "frm_slogic.frx":1E2A
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":1E4A
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      Picture         =   "frm_slogic.frx":1E66
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6240
      Width           =   495
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      Picture         =   "frm_slogic.frx":22A8
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6840
      Width           =   495
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      Picture         =   "frm_slogic.frx":26EA
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7440
      Width           =   495
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      Picture         =   "frm_slogic.frx":2B2C
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   8040
      Width           =   495
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   2
      Left            =   10080
      TabIndex        =   32
      Top             =   1800
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_slogic.frx":2F6E
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
      Tip             =   "frm_slogic.frx":2F8E
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":2FAE
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   3
      Left            =   4005
      TabIndex        =   4
      Top             =   4200
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_slogic.frx":2FCA
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
      Tip             =   "frm_slogic.frx":2FEA
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":300A
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   4
      Left            =   4005
      TabIndex        =   5
      Top             =   4680
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_slogic.frx":3026
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
      Tip             =   "frm_slogic.frx":3046
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":3066
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   5
      Left            =   4005
      TabIndex        =   6
      Top             =   5160
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_slogic.frx":3082
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
      Tip             =   "frm_slogic.frx":30A2
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":30C2
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      Picture         =   "frm_slogic.frx":30DE
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4200
      Width           =   495
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      Picture         =   "frm_slogic.frx":3520
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4680
      Width           =   495
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      Picture         =   "frm_slogic.frx":3962
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5160
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
      FormDesignHeight=   8940
      FormDesignWidth =   12975
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_browse 
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   720
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
      Caption         =   "frm_slogic.frx":3DA4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_slogic.frx":3DD0
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":3DF0
   End
   Begin HexUniControls.ctlUniListBoxXP List3 
      Height          =   255
      Left            =   15840
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   11400
      Visible         =   0   'False
      Width           =   975
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
      Tip             =   "frm_slogic.frx":3E0C
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":3E2C
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniListBoxXP List2 
      Height          =   255
      Left            =   9120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
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
      Tip             =   "frm_slogic.frx":3E48
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":3E68
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniListBoxXP List1 
      Height          =   2025
      Left            =   120
      TabIndex        =   3
      Top             =   1920
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
      Tip             =   "frm_slogic.frx":3E84
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   -1  'True
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":3EA4
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniLabel Label16 
      Height          =   405
      Left            =   9480
      Top             =   360
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
      Caption         =   "frm_slogic.frx":3EC0
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":3EF0
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":3F10
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   7
      Left            =   6795
      Top             =   6240
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
      Caption         =   "frm_slogic.frx":3F2C
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":3F6A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":3F8A
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   9
      Left            =   6795
      Top             =   6840
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
      Caption         =   "frm_slogic.frx":3FA6
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":3FEC
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":400C
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   11
      Left            =   6795
      Top             =   7440
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
      Caption         =   "frm_slogic.frx":4028
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":4076
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":4096
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   13
      Left            =   6795
      Top             =   8055
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
      Caption         =   "frm_slogic.frx":40B2
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":4102
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":4122
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   6
      Left            =   705
      Top             =   6240
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
      Caption         =   "frm_slogic.frx":413E
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":417C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":419C
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   8
      Left            =   705
      Top             =   6840
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
      Caption         =   "frm_slogic.frx":41B8
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":41FE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":421E
   End
   Begin HexUniControls.ctlUniLabel Label17 
      Height          =   375
      Left            =   4080
      Top             =   5760
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
      Caption         =   "frm_slogic.frx":423A
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":4280
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":42A0
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   10
      Left            =   705
      Top             =   7440
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
      Caption         =   "frm_slogic.frx":42BC
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":430A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":432A
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   12
      Left            =   705
      Top             =   8055
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
      Caption         =   "frm_slogic.frx":4346
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":4396
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":43B6
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
      Caption         =   "frm_slogic.frx":43D2
      BackColor       =   -2147483633
      ForeColor       =   12582912
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":43F4
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":4414
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
      Caption         =   "frm_slogic.frx":4430
      BackColor       =   -2147483633
      ForeColor       =   12582912
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":4452
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":4472
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
      Caption         =   "frm_slogic.frx":448E
      BackColor       =   -2147483633
      ForeColor       =   12582912
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":44B0
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":44D0
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   2
      Left            =   10080
      Top             =   1440
      Visible         =   0   'False
      Width           =   1995
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
      Caption         =   "frm_slogic.frx":44EC
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":4528
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":4548
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   3
      Left            =   705
      Top             =   4200
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
      Caption         =   "frm_slogic.frx":4564
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":4596
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":45B6
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   4
      Left            =   705
      Top             =   4680
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
      Caption         =   "frm_slogic.frx":45D2
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":45FC
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":461C
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   5
      Left            =   705
      Top             =   5160
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
      Caption         =   "frm_slogic.frx":4638
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":467E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":469E
   End
   Begin HexUniControls.ctlUniLabel LBL_Calibname 
      Height          =   405
      Left            =   6840
      Top             =   2400
      Width           =   5805
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
      Caption         =   "frm_slogic.frx":46BA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":46DA
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":46FA
   End
   Begin HexUniControls.ctlUniLabel lbl_cpfname 
      Height          =   405
      Left            =   2760
      Top             =   840
      Width           =   5925
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
      Caption         =   "frm_slogic.frx":4716
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":4736
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":4756
   End
   Begin HexUniControls.ctlUniLabel Label9 
      Height          =   375
      Left            =   6840
      Top             =   1920
      Width           =   2775
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
      Caption         =   "frm_slogic.frx":4772
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":47B2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":47D2
   End
   Begin HexUniControls.ctlUniLabel Label4 
      Height          =   375
      Left            =   2760
      Top             =   360
      Width           =   5895
      _ExtentX        =   10398
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
      Caption         =   "frm_slogic.frx":47EE
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":4834
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":4854
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
      Caption         =   "frm_slogic.frx":4870
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   -1  'True
      Tip             =   "frm_slogic.frx":4898
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":48B8
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
      Caption         =   "frm_slogic.frx":48D4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":490C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":492C
   End
   Begin HexUniControls.ctlUniLabel Label2 
      Height          =   255
      Left            =   7080
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
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
      Caption         =   "frm_slogic.frx":4948
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":4984
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":49A4
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   1560
      Width           =   2175
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
      Caption         =   "frm_slogic.frx":49C0
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_slogic.frx":4A06
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":4A26
   End
   Begin HexUniControls.ctlUniTextBoxXP PredSpecNum 
      Height          =   285
      Left            =   7200
      TabIndex        =   33
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
      Text            =   "frm_slogic.frx":4A42
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
      Tip             =   "frm_slogic.frx":4A64
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_slogic.frx":4A84
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
      Caption         =   "frm_slogic.frx":4AA0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pname2 As String

Private m_projectName As String

Private Const OUT_CALIBPROTECTED = 60       ' CalStar calibration is protected error
Private Const OUT_ERRORNODONGLE = -100      ' CalStar calibration is protected error, no hardlock found

Sub lmslpred()
  Dim serName As String
  Dim calName As String
  Dim serInfo As TSeriesInfo
  Dim i As Integer
  Dim inputData As TVBInputData
  Dim predInput As TPredictInput
  Dim predResult As TResultData
  Dim x As Double
  Dim slwl() As Single
  Dim slspec() As Single
  Dim slmin, slmax, slinc As Single
  Dim slpts, ff As Integer
  Dim errMsg As String
  Dim uniMsg As String

  uniMsg = MLSupport.GGS_Params("frmMain.statMsg1", "Loading CalStar model: %1", unity_main.modlname)
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Loading CalStar model: " & unity_main.modlname), uniMsg)

  m_projectName = unity_main.fullmodelname  'Senslogic FILE
  unity_main.pukedonpred = False
  open_project
  
  If (unity_main.pukedonpred = True) Then
    GoTo BAD_PRED
  End If
  
  frmMain.parse_cs
  
  If (unity_main.pukedonpred = True) Then
    GoTo BAD_PRED
  End If
  
  calName = Trim(frmMain.List1.Text)
  
  uniMsg = MLSupport.GGS_Params("frm_calstar.statMsg2", "Using Calibration name: %1", calName)
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Using Calibration name: " & calName), uniMsg)
  
#If ABBFT Then
  slmin = CInt(unity_main.m_mb3000.m_startWavenum)
  slinc = CInt(unity_main.m_mb3000.m_waveNumIncr)
  slpts = unity_main.m_mb3000.m_endWavenumIndx - unity_main.m_mb3000.m_startWavenumIndx
#Else
  slmin = unity_main.m_smplStartWvln
  slmax = unity_main.m_smplEndWvln
  slinc = MS11CfgData.wvlnIncr
  slpts = (slmax - slmin) / slinc
#End If

  ReDim slspec(slpts)
  ReDim slwl(slpts)
    
#If ABBFT Then
  For ff = 0 To slpts                   ' put our current spectrum into the senslogic
    slwl(ff) = slmin + (ff * slinc)     ' set the wavelength value
    slspec(ff) = FTSmplAbsYVals(unity_main.m_mb3000.m_startWavenumIndx + ff)
  Next ff
#Else
  ' Check if spectrum data treated
  If (unity_main.m_enableTreatment = True) Then
    For ff = 0 To slpts                   ' put our current spectrum into the senslogic
      slwl(ff) = slmin + (ff * slinc)     ' set the wavelength value
      slspec(ff) = ProdTreatAbsYVals(ff)
    Next ff
  Else
    For ff = 0 To slpts                   ' put our current spectrum into the senslogic
      slwl(ff) = slmin + (ff * slinc)     ' set the wavelength value
      slspec(ff) = ProdAbsYVals(ff)
    Next ff
  End If
#End If

  inputData.calibName = calName + Chr$(0)
  inputData.ProjectFileName = m_projectName + Chr$(0)
  inputData.ConstH = 3
  inputData.ConstSR = 3
  inputData.FullRange = 1
  inputData.AbsFlag = 1
  inputData.NumData = slpts + 1

  i = FillQuanPredictionStruct(inputData, slspec(0), slwl(0), predInput) 'from senslogic - returns 30 if bad
  
  If (i = 0) Then
    i = QuanPrediction(predInput, predResult)
    
    If (i = 0) Then
      ' Check for infinity value
      On Error Resume Next
      x = predResult.Prediction
      
      If ((x - x) <> 0) Then
        unity_main.tempval = -9999
      Else
        unity_main.tempval = x
      End If
    
      unity_main.tempval = unity_main.tempval * unity_main.tempskew '3/31
      unity_main.tempval = unity_main.tempval + unity_main.tempbias '3/31
      unity_main.preds.AddItem (unity_main.tempval)
      
      ' Check for infinity value
      On Error Resume Next
      x = predResult.Leverage
      
      If ((x - x) <> 0) Then
        unity_main.tempmdist = -9999
      Else
        unity_main.tempmdist = x
      End If
      
      unity_main.tempmdist = Int(unity_main.tempmdist * 100)
      unity_main.tempmdist = (unity_main.tempmdist / 100)
      unity_main.lstmd.AddItem unity_main.tempmdist
      
      ' Check for infinity value
      On Error Resume Next
      x = predResult.SpecRecon
      
      If ((x - x) <> 0) Then
        unity_main.lstrr.AddItem -9999
      Else
        unity_main.lstrr.AddItem x
      End If
      
      unity_main.lstrr2.AddItem "1"
      unity_main.lst_nd.AddItem unity_main.m_noOLVal
      unity_main.lst_pfexp.AddItem unity_main.m_noOLVal
    Else
      errMsg = ("CalStar QuanPrediction error: " & i)
      uniMsg = MLSupport.GGS_Params("frmMain.errMsg3", "CalStar QuanPrediction error: %1", CStr(i))
      Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
      unity_main.pukedonpred = True
    End If
  Else
    If (i = OUT_CALIBPROTECTED) Or (i = OUT_ERRORNODONGLE) Then
      errMsg = "CalStar calibration protection error; missing dongle or invalid code number for dongle"
      uniMsg = MLSupport.GSS("frmMain", "errMsg4", "CalStar calibration protection error; missing dongle or invalid code number for dongle")
    Else
      errMsg = ("CalStar FillQuanPredictionStruct() error: " & i)
      uniMsg = MLSupport.GGS_Params("frmMain.errMsg4", "CalStar FillQuanPredictionStruct() error: %1", CStr(i))
    End If
    
    Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    unity_main.pukedonpred = True
  End If
  
BAD_PRED:
  i = FreeQuanPredictionStruct(predInput)
  close_project
End Sub

Public Sub opentoedit()
  Dim sFile As String
  
  sFile = frmMain.pname2
  m_projectName = pname2
  open_project
  lbl_cpfname.Caption = (m_projectName)
End Sub

Public Sub parse_cs()
  Dim l1, l2, ls, zz As Integer
  Dim tempstring, inString, onechar As String
  Dim gotcal As Boolean
  Dim tempint As Integer
  Dim start2 As Integer
  Dim errMsg As String
  Dim uniMsg As String

  inString = UCase(unity_main.slcal)  '7/9/04
  ls = Len(inString)
  tempstring = ""
  
  For zz = 2 To ls
    onechar = Mid(inString, zz, 1)
    If (onechar <> "S") Then
      tempstring = tempstring & onechar
    Else
      l1 = tempstring
      start2 = zz
      Exit For
    End If
  Next zz
  
  l2 = Right(inString, (ls - start2))
  tempint = l1

  If (tempint >= frmMain.List1.ListCount) Then
    errMsg = m_projectName & " CalStar model project file invalid"
    uniMsg = MLSupport.GGS_Params("frmMain.errMsg5", "%1 model project file invalid", m_projectName)
    Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    unity_main.pukedonpred = True
    Exit Sub
  End If

  frmMain.List1.ListIndex = tempint
  tempint = l2

  If (tempint >= frmMain.List2.ListCount) Then
    errMsg = m_projectName & " CalStar model project file invalid"
    uniMsg = MLSupport.GGS_Params("frmMain.errMsg5", "%1 model project file invalid", m_projectName)
    Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    unity_main.pukedonpred = True
    Exit Sub
  End If

  frmMain.List2.ListIndex = tempint
    
End Sub

Private Sub open_project()
  Dim i As Integer
  Dim CalInfo As TCalibInfo
  Dim pWave() As Single
  Dim errMsg As String
  Dim tmpStrg, TempStr As String
  Dim pos As Long
  Dim uniMsg As String
        
  If (IsDatabaseOpen > 0) Then
    CloseCTreeDb
  End If
    
  If (IsProjectOpen > 0) Then
    CloseProject
  End If
    
  If (IsDatabaseOpen > 0) Then
    CloseCTreeDb
  End If
    
  List1.clear
  List2.clear
  List3.clear
    
  i = IsDatabaseOpen
  i = OpenCTreeDb
  i = IsDatabaseOpen
  i = IsProjectOpen
  
  If (CFile.st_FileExist(m_projectName) = False) Then
    errMsg = m_projectName & " CalStar model file not found"
    uniMsg = MLSupport.GGS_Params("fileErrMsg6", "%1 file not found", m_projectName)
    Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    unity_main.pukedonpred = True
    Exit Sub
  End If
  
  i = OpenProject(m_projectName)
    
  If (i <> 0) Then
    errMsg = m_projectName & " CalStar model file open error: " & i
    uniMsg = MLSupport.GGS_Params("frmMain.errMsg1", "%1 CalStar model file open error: %2", m_projectName, CStr(i))
    Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    unity_main.pukedonpred = True
    Exit Sub
  End If
    
  i = IsProjectOpen
    
  If (i = 0) Then
    errMsg = unity_main.fullmodelname & " CalStar model project open error"
    uniMsg = MLSupport.GGS_Params("frmMain.errMsg2", "%1 CalStar model project open error", m_projectName)
    Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    unity_main.pukedonpred = True
    Exit Sub
  End If
    
  ' Build list of calibration names
  TempStr = Space$(80)
  i = GetCalibName(TempStr, 1)
    
  If (i = 0) Then
    pos = InStrRev(TempStr, ChrW(0))
    
    If (pos <> 0) Then
      tmpStrg = Trim(Left(TempStr, pos - 1))
    Else
      tmpStrg = Trim(TempStr)
    End If
  
    List1.AddItem (tmpStrg)
  End If
    
  While (i = 0)
    TempStr = Space$(80)
    i = GetCalibName(TempStr, 0)
      
    If (i = 0) Then
      pos = InStrRev(TempStr, ChrW(0))
    
      If (pos <> 0) Then
        tmpStrg = Trim(Left(TempStr, pos - 1))
      Else
        tmpStrg = Trim(TempStr)
      End If
    
      List1.AddItem (tmpStrg)
    End If
  Wend
    
  If (List1.ListCount > 0) Then
    List1.ListIndex = 0
  End If
    
  ' Build list of series names
  TempStr = Space$(80)
  i = GetSeriesName(TempStr, 1)
  
  If (i = 0) Then
    pos = InStrRev(TempStr, ChrW(0))
    
    If (pos <> 0) Then
      tmpStrg = Trim(Left(TempStr, pos - 1))
    Else
      tmpStrg = Trim(TempStr)
    End If
  
    List2.AddItem (tmpStrg)
  End If
    
  While i = 0
    TempStr = Space$(80)
    i = GetSeriesName(TempStr, 0)
      
    If (i = 0) Then
      pos = InStrRev(TempStr, ChrW(0))
    
      If (pos <> 0) Then
        tmpStrg = Trim(Left(TempStr, pos - 1))
      Else
        tmpStrg = Trim(TempStr)
      End If

      List2.AddItem (tmpStrg)
    End If
  Wend
    
  If (List2.ListCount > 0) Then
    List2.ListIndex = 0
  End If
End Sub

Private Sub close_project()
    
    If IsProjectOpen > 0 Then
      CloseProject
    End If
    
    If IsDatabaseOpen > 0 Then
      CloseCTreeDb
    End If

    List1.clear
    List2.clear
    List3.clear
End Sub

Private Sub cmd_browse_Click()
  Dim dialog As New clsCommDialogs
  Dim fileDir As String
  Dim sFilter As String
  Dim dlgTitle As String
  Dim fileName As String
  Dim pos As Long

  unity_main.errorstring = "CalStar Model Property Configuration screen Browse button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
    
  dialog.InitDialogs
  fileDir = MODELS_DIR
  sFilter = ("CWS Project Files (*" & CALSTAR_MODEL_FILE_EXT & ")" & Chr(0) & "*" & CALSTAR_MODEL_FILE_EXT & Chr(0))
  dlgTitle = MLSupport.GSS("frmMain", "dlgTitle", "Select Model File")
  fileName = dialog.ShowOpen(Me.hwnd, sFilter, dlgTitle, OFN_FILEMUSTEXIST + OFN_HIDEREADONLY + OFN_NOCHANGEDIR, fileDir)
 
  If (fileName <> "") Then
    m_projectName = fileName
    open_project
    lbl_cpfname.Caption = (m_projectName)
  End If
End Sub

Private Sub cmd_cancel_Click()
  
  unity_main.errorstring = "CalStar Model Property Configuration screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  close_project
  
  If (frmedmod.m_addProp = True) Then
    frmedmod.m_addProp = False
    frmedmod.grid_models.MaxRows = frmedmod.grid_models.MaxRows - 1
  End If
  
  Unload frmMain
End Sub

Private Sub cmd_save_Click()
  Dim lennum As Integer
  Dim onechar As String
  Dim rebuildit As String
  Dim wasf As String
  Dim fieldcounter As Integer
  Dim buildcounter As Integer
  Dim xx, zz As Integer
  
  unity_main.errorstring = "CalStar Model Property Configuration screen Save Changes button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  If (Trim(txt_propname.Text) = "") Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frmMain", "errMsg1", "You must enter a property name"), vbCritical
    frmMain.txt_propname.SetFocus
    Exit Sub
  End If
  
  If (List1.ListIndex = -1) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frmMain", "errMsg2", "You must select a calibration"), vbCritical
    frmMain.List1.SetFocus
    Exit Sub
  End If
  
  For xx = 3 To 13
    If (Trim(frmMain.txt_modvar(xx).Text) = "") Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frmMain", "errMsg3", "You must enter a value in all fields"), vbCritical
      frmMain.txt_modvar(xx).SetFocus
      Exit Sub
    End If
  Next xx

  For fieldcounter = 3 To 13
    wasf = Trim(frmMain.txt_modvar(fieldcounter).Text)
    lennum = Len(wasf)
    rebuildit = ""
    
    For buildcounter = 1 To lennum
      onechar = Mid(wasf, buildcounter, 1)
      
      If (onechar = ",") Then
        onechar = "."
      End If
        
      rebuildit = rebuildit & onechar
    Next buildcounter
    
    frmMain.txt_modvar(fieldcounter).Text = rebuildit
  Next fieldcounter

  frmedmod.grid_models.Row = CInt(frmMain.txtmrow.Text) '4/1/05
  frmedmod.grid_models.Col = 1
  frmedmod.grid_models.Text = frmMain.txt_propname.Text
  frmedmod.grid_models.Col = 2
  frmedmod.grid_models.Text = frmMain.lbl_cpfname.Caption
  frmedmod.grid_models.Col = 3
  frmMain.Text1.Text = "C" & List1.ListIndex & "S" & List2.ListIndex
  frmedmod.grid_models.Text = Trim(frmMain.Text1.Text)

  For zz = 4 To frmedmod.grid_models.MaxCols - 1
    frmedmod.grid_models.Col = zz
    frmedmod.grid_models.Text = frmMain.txt_modvar(zz - 1).Text
  Next zz

  frmedmod.grid_models.Col = frmedmod.grid_models.MaxCols
  frmedmod.grid_models.Text = "NA"

  If (frmedmod.m_addProp = True) Then
    frmedmod.m_addProp = False
    frmedmod.numprops.Text = frmedmod.numprops.Text + 1
  End If
  
  frmedmod.fillproplist
  close_project
  Unload frmMain
End Sub

Private Sub Form_Load()
    
  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
  Label8.Caption = ""
End Sub

Private Sub Picture1_Click()
  
  unity_main.formfrom = 13
  unity_main.varfrom = 1
  frm_kybd.lbl_kybd.Caption = frmMain.Label16.Caption
  frm_kybd.txt_kybd.Text = Trim(frmMain.txt_propname.Text)
  frm_kybd.Show 1
End Sub

Private Sub Picture4_Click()
  
  unity_main.formfrom = 6
  unity_main.varfrom = 3
  frm_numpad.lbl_num.Caption = frmMain.Label1(3).Caption
  frm_numpad.txt_num.Text = txt_modvar(3).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture5_Click()
  
  unity_main.formfrom = 6
  unity_main.varfrom = 4
  frm_numpad.lbl_num.Caption = frmMain.Label1(4).Caption
  frm_numpad.txt_num.Text = txt_modvar(4).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture6_Click()
  
  unity_main.formfrom = 6
  unity_main.varfrom = 5
  frm_numpad.lbl_num.Caption = frmMain.Label1(5).Caption
  frm_numpad.txt_num.Text = txt_modvar(5).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture7_Click()
  
  unity_main.formfrom = 6
  unity_main.varfrom = 6
  frm_numpad.lbl_num.Caption = frmMain.Label1(6).Caption
  frm_numpad.txt_num.Text = txt_modvar(6).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture8_Click()
  
  unity_main.formfrom = 6
  unity_main.varfrom = 7
  frm_numpad.lbl_num.Caption = frmMain.Label1(7).Caption
  frm_numpad.txt_num.Text = txt_modvar(7).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture9_Click()
  
  unity_main.formfrom = 6
  unity_main.varfrom = 8
  frm_numpad.lbl_num.Caption = frmMain.Label1(8).Caption
  frm_numpad.txt_num.Text = txt_modvar(8).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture10_Click()
  
  unity_main.formfrom = 6
  unity_main.varfrom = 9
  frm_numpad.lbl_num.Caption = frmMain.Label1(9).Caption
  frm_numpad.txt_num.Text = txt_modvar(9).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture11_Click()
  
  unity_main.formfrom = 6
  unity_main.varfrom = 10
  frm_numpad.lbl_num.Caption = frmMain.Label1(10).Caption
  frm_numpad.txt_num.Text = txt_modvar(10).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture12_Click()
  
  unity_main.formfrom = 6
  unity_main.varfrom = 11
  frm_numpad.lbl_num.Caption = frmMain.Label1(11).Caption
  frm_numpad.txt_num.Text = txt_modvar(11).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture13_Click()
  
  unity_main.formfrom = 6
  unity_main.varfrom = 12
  frm_numpad.lbl_num.Caption = frmMain.Label1(12).Caption
  frm_numpad.txt_num.Text = txt_modvar(12).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture14_Click()
  
  unity_main.formfrom = 6
  unity_main.varfrom = 13
  frm_numpad.lbl_num.Caption = frmMain.Label1(13).Caption
  frm_numpad.txt_num.Text = txt_modvar(13).Text
  frm_numpad.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  If (Me.WindowState <> vbMinimized) Then
    SaveSetting App.Title, "Settings", "MainLeft", Me.Left
    SaveSetting App.Title, "Settings", "MainTop", Me.Top
    SaveSetting App.Title, "Settings", "MainWidth", Me.Width
    SaveSetting App.Title, "Settings", "MainHeight", Me.Height
  End If
End Sub

Private Sub List1_Click()
  Dim i As Integer
  Dim CalInfo As TCalibInfo
  Dim pWaves(1000) As Single
  
  i = GetCalibInfo(List1.Text, CalInfo, pWaves(0))
  LBL_Calibname.Caption = List1.Text
End Sub

Private Sub List2_Click()
  Dim serInfo As TSeriesInfo
  Dim i As Integer
    
  i = GetSeriesInfo(List2.Text, serInfo)
  
  If (serInfo.NumSpectra > 0) Then
    i = PredSpecNum.Text
    
    If (i < 1) Then
      PredSpecNum.Text = 1
    ElseIf (i > serInfo.NumSpectra) Then
      PredSpecNum.Text = serInfo.NumSpectra
    End If
  Else
    PredSpecNum.Text = 0
  End If
End Sub

Private Sub List3_Click()
  Dim S, S1 As String
  Dim MethInfo As TMethodInfo
  Dim i As Integer
  Dim SpecInfo As TSpecInfo
  Dim myWaves() As Single
  Dim Min, Max, inc As Single
  
  If Label6.Caption = "Available Spectra" Then    'show spectrum information
    ReDim myWaves(2000)
    S = List3.Text
    i = InStrRev(S, "(")    'isolate the serial number
    
    If (i < 2) Then
      Exit Sub
    End If
    
    S = Mid(S, i + 1)
    i = InStrRev(S, ")")
    S = Mid(S, 1, i - 1)
    
    i = GetSpecInfo(Val(S), SpecInfo)
    i = GetWaveData(SpecInfo.iWavelengths, Min, Max, inc, myWaves(0))
    S = "Spectrum Info:" & Chr(13)
    S = S & "Wavelength Min:" & Min & Chr(13)
    S = S & "Wavelength Max:" & Max & Chr(13)
    
    If (inc = 0) Then
      S = S & "Filter spectrum" & Chr(13)
    Else
      S = S & "Wavelength Inc:" & inc & Chr(13)
    End If
    
    Label8.Caption = S
    ReDim myWaves(0)
  End If

  If (Label6.Caption = "Available Methods") Then    'Method information
    If (List3.ListCount > 0) Then
      MethInfo.calibName = "Calib Name" & Chr(0)
      MethInfo.MethName = "Method Name" & Chr(0)
      MethInfo.InstrType = 1
      MethInfo.dummy = 234
      i = GetMethodInfo(List3.Text, MethInfo)
      
      If (i = 0) Then
        S = "Method Info:" & Chr(13)
        S = S & "Method Status:" & MethInfo.MethStatus & Chr(13)
        S1 = MethInfo.calibName
        
        For i = 1 To Len(S1)
          If (Mid(S1, i, 1) = Chr(0)) Then
            Mid(S1, i, 1) = " "
          End If
        Next i
        
        S1 = RTrim(S1)
        S = S & "Calibration Name:" & S1 & Chr(13)
        S = S & "Calibration Status:" & MethInfo.CalibStatus & Chr(13)
        
        If ((MethInfo.CalibStatus And dsLocked) = 0) Then
          S = S & "Calibration is not locked" & Chr(13)
        Else
          S = S & "Calibration is locked" & Chr(13)
        End If
      
        S1 = ""
        S1 = MethInfo.MethName
        
        For i = 1 To Len(S1)
          If (Mid(S1, i, 1) = Chr(0)) Then
            Mid(S1, i, 1) = " "
          End If
        Next i
        
        S1 = RTrim(S1)
      
        If (MethInfo.InstrType = 0) Then  'method is validated if MethInfo.InstrType<>0 !!!
          S = S & "Method '" & S1 & "' is not validated" & Chr(13)
        Else
          S = S & "Method '" & S1 & "' is validated" & Chr(13)
        End If
        
        Label8.Caption = S
      End If
    End If
  End If
End Sub

Private Sub txt_propname_DblClick(Button As Integer)
  
  unity_main.formfrom = 13
  unity_main.varfrom = 1
  frm_kybd.txt_kybd.Text = Trim(frmMain.txt_propname.Text)
  frm_kybd.Show 1
End Sub








