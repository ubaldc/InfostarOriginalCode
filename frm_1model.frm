VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_1model 
   Caption         =   "Model Property Configuration"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
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
   Icon            =   "frm_1model.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniButtonImageXP cmd_browse 
      Height          =   650
      Left            =   8040
      TabIndex        =   3
      Top             =   720
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
      Caption         =   "frm_1model.frx":0442
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_1model.frx":046E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":048E
   End
   Begin HexUniControls.ctlUniListBoxXP lst_modtype 
      Height          =   1020
      Left            =   8040
      TabIndex        =   31
      Top             =   480
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   0
      _ExtentY        =   0
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
      Tip             =   "frm_1model.frx":04AA
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":04CA
      ManualStart     =   0   'False
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10440
      Picture         =   "frm_1model.frx":04E6
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6120
      Width           =   495
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4920
      Picture         =   "frm_1model.frx":0928
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6120
      Width           =   495
   End
   Begin VB.PictureBox Picture12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10440
      Picture         =   "frm_1model.frx":0D6A
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5520
      Width           =   495
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4920
      Picture         =   "frm_1model.frx":11AC
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5520
      Width           =   495
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10440
      Picture         =   "frm_1model.frx":15EE
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4920
      Picture         =   "frm_1model.frx":1A30
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4920
      Width           =   495
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10440
      Picture         =   "frm_1model.frx":1E72
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4920
      Picture         =   "frm_1model.frx":22B4
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4320
      Width           =   495
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4920
      Picture         =   "frm_1model.frx":26F6
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3240
      Width           =   495
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4920
      Picture         =   "frm_1model.frx":2B38
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4920
      Picture         =   "frm_1model.frx":2F7A
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2280
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4920
      Picture         =   "frm_1model.frx":33BC
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1800
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6960
      Picture         =   "frm_1model.frx":37FE
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   840
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      Picture         =   "frm_1model.frx":3C40
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   240
      Width           =   495
   End
   Begin HexUniControls.ctlUniTextBoxXP txtmrow 
      Height          =   495
      Left            =   240
      TabIndex        =   32
      Top             =   8880
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_1model.frx":4082
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
      Tip             =   "frm_1model.frx":40A2
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":40C2
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8400
      Top             =   2520
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9405
      FormDesignWidth =   11250
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   8040
      TabIndex        =   0
      Top             =   7920
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
      Caption         =   "frm_1model.frx":40DE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_1model.frx":410A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":412A
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   650
      Left            =   4920
      TabIndex        =   16
      Top             =   7920
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
      Caption         =   "frm_1model.frx":4146
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_1model.frx":417E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":419E
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   13
      Left            =   9000
      TabIndex        =   15
      Top             =   6120
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_1model.frx":41BA
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
      Tip             =   "frm_1model.frx":41DA
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":41FA
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   12
      Left            =   3480
      TabIndex        =   11
      Top             =   6120
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_1model.frx":4216
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
      Tip             =   "frm_1model.frx":4236
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":4256
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   11
      Left            =   9000
      TabIndex        =   14
      Top             =   5520
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_1model.frx":4272
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
      Tip             =   "frm_1model.frx":4292
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":42B2
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   10
      Left            =   3480
      TabIndex        =   10
      Top             =   5520
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_1model.frx":42CE
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
      Tip             =   "frm_1model.frx":42EE
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":430E
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   9
      Left            =   9000
      TabIndex        =   13
      Top             =   4920
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_1model.frx":432A
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
      Tip             =   "frm_1model.frx":434A
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":436A
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   8
      Left            =   3480
      TabIndex        =   9
      Top             =   4920
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_1model.frx":4386
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
      Tip             =   "frm_1model.frx":43A6
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":43C6
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   7
      Left            =   9000
      TabIndex        =   12
      Top             =   4320
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_1model.frx":43E2
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
      Tip             =   "frm_1model.frx":4402
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":4422
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   6
      Left            =   3480
      TabIndex        =   8
      Top             =   4320
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_1model.frx":443E
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
      Tip             =   "frm_1model.frx":445E
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":447E
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   5
      Left            =   3480
      TabIndex        =   7
      Top             =   3240
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_1model.frx":449A
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
      Tip             =   "frm_1model.frx":44BA
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":44DA
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   4
      Left            =   3480
      TabIndex        =   6
      Top             =   2760
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_1model.frx":44F6
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
      Tip             =   "frm_1model.frx":4516
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":4536
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   3
      Left            =   3480
      TabIndex        =   5
      Top             =   2280
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_1model.frx":4552
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
      Tip             =   "frm_1model.frx":4572
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":4592
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   4
      Top             =   1800
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_1model.frx":45AE
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
      Tip             =   "frm_1model.frx":45CE
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":45EE
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   420
      Index           =   1
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   741
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_1model.frx":460A
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
      Tip             =   "frm_1model.frx":462A
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":464A
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_modvar 
      Height          =   420
      Index           =   0
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   741
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_1model.frx":4666
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
      Tip             =   "frm_1model.frx":4686
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":46A6
   End
   Begin HexUniControls.ctlUniLabel Label3 
      Height          =   375
      Left            =   8040
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
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
      Caption         =   "frm_1model.frx":46C2
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":46F6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":4716
   End
   Begin HexUniControls.ctlUniLabel lbl_note 
      Height          =   615
      Left            =   120
      Top             =   8040
      Width           =   4335
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
      Caption         =   "frm_1model.frx":4732
      BackColor       =   -2147483633
      ForeColor       =   0
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":47A8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":47C8
   End
   Begin HexUniControls.ctlUniLabel Label4 
      Height          =   555
      Left            =   5760
      Top             =   1725
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_1model.frx":47E4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":48B2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":48D2
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   13
      Left            =   5640
      Top             =   6135
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
      Caption         =   "frm_1model.frx":48EE
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":493E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":495E
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   12
      Left            =   165
      Top             =   6135
      Width           =   3150
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
      Caption         =   "frm_1model.frx":497A
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":49CA
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":49EA
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   11
      Left            =   5640
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
      Caption         =   "frm_1model.frx":4A06
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":4A54
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":4A74
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   10
      Left            =   165
      Top             =   5520
      Width           =   3150
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
      Caption         =   "frm_1model.frx":4A90
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":4ADE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":4AFE
   End
   Begin HexUniControls.ctlUniLabel Label2 
      Height          =   375
      Left            =   3480
      Top             =   3840
      Width           =   6960
      _ExtentX        =   12277
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
      Caption         =   "frm_1model.frx":4B1A
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":4B60
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":4B80
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   9
      Left            =   5640
      Top             =   4920
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
      Caption         =   "frm_1model.frx":4B9C
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":4BE2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":4C02
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   8
      Left            =   165
      Top             =   4920
      Width           =   3150
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
      Caption         =   "frm_1model.frx":4C1E
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":4C64
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":4C84
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   7
      Left            =   5640
      Top             =   4320
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
      Caption         =   "frm_1model.frx":4CA0
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":4CDE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":4CFE
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   6
      Left            =   165
      Top             =   4320
      Width           =   3150
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
      Caption         =   "frm_1model.frx":4D1A
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":4D58
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":4D78
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   5
      Left            =   165
      Top             =   3240
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
      Caption         =   "frm_1model.frx":4D94
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":4DDA
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":4DFA
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   4
      Left            =   165
      Top             =   2760
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
      Caption         =   "frm_1model.frx":4E16
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":4E40
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":4E60
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   3
      Left            =   165
      Top             =   2280
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
      Caption         =   "frm_1model.frx":4E7C
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":4EAE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":4ECE
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   2
      Left            =   165
      Top             =   1800
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
      Caption         =   "frm_1model.frx":4EEA
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":4F26
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":4F46
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   1
      Left            =   240
      Top             =   840
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
      Caption         =   "frm_1model.frx":4F62
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":4F96
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":4FB6
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Index           =   0
      Left            =   240
      Top             =   240
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
      Caption         =   "frm_1model.frx":4FD2
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_1model.frx":5002
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_1model.frx":5022
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   10080
      Top             =   2520
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
      Left            =   9000
      Top             =   2520
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_1model.frx":503E
   End
End
Attribute VB_Name = "frm_1model"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public modtypex As Integer  '1=pls,2=mlr,3=Secondary, 4=PRD, 5=CalStar

Private Sub cmd_save_Click()
  Dim lennum As Integer
  Dim onechar As String
  Dim rebuildit As String
  Dim wasf As String
  Dim fieldcounter As Integer
  Dim buildcounter As Integer
  Dim xx As Integer
  Dim filePathName As String
  Dim modExt As String
  Dim uniFile As New clsUniFile
  
  Select Case (frm_1model.modtypex)
    Case 1          'pls model
      unity_main.errorstring = "GRAMS PLSIQ Model Property Configuration screen Save Changes button selected"
    Case 2          'mlr model
      unity_main.errorstring = "MLR Model Property Configuration screen Save Changes button selected"
    Case 3          'secondary model
      unity_main.errorstring = "Secondary Model Property Configuration screen Save Changes button selected"
  End Select
  
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  For xx = 0 To 13
    If (Trim(frm_1model.txt_modvar(xx).Text) = "") Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_1model", "errMsg1", "You must enter a value in all fields"), vbCritical
      frm_1model.txt_modvar(xx).SetFocus
      Exit Sub
   End If
  Next xx
  
  ' Confirm path contains '\' instead of '/'
  filePathName = txt_modvar(1).Text
  check_filepathname_delimiters filePathName
  
  If (InStr(filePathName, "\") = 0) Then
    filePathName = (MODELS_DIR & filePathName)
  End If
  
  ' Check file name extension
  modExt = ("." & LCase(uniFile.st_FileExt(filePathName)))
  
  Select Case (frm_1model.modtypex)
    Case 1          'pls model
      If (modExt <> GRAMS_MODEL_FILE_EXT) Then
        filePathName = uniFile.st_FilePath(filePathName) & "\" & uniFile.st_FileNameNoExt(filePathName) & GRAMS_MODEL_FILE_EXT
      End If
      
    Case 2          'mlr model
      If (modExt <> MLR_MODEL_FILE_EXT) Then
        filePathName = uniFile.st_FilePath(filePathName) & "\" & uniFile.st_FileNameNoExt(filePathName) & MLR_MODEL_FILE_EXT
      End If
      
    Case 3          'secondary model
      If (modExt <> SEC_MODEL_FILE_EXT) Then
        filePathName = uniFile.st_FilePath(filePathName) & "\" & uniFile.st_FileNameNoExt(filePathName) & SEC_MODEL_FILE_EXT
      End If
  End Select
  
  txt_modvar(1).Text = filePathName
  
  For fieldcounter = 3 To 13
    wasf = Trim(frm_1model.txt_modvar(fieldcounter).Text)
    lennum = Len(wasf)
    rebuildit = ""
    
    For buildcounter = 1 To lennum
      onechar = Mid(wasf, buildcounter, 1)
      
      If (onechar = ",") Then
        onechar = "."
      End If
      
      rebuildit = rebuildit & onechar
    Next buildcounter
    
    frm_1model.txt_modvar(fieldcounter).Text = rebuildit
  Next fieldcounter

  frmedmod.grid_models.Row = CInt(frm_1model.txtmrow.Text)

  For xx = 1 To frmedmod.grid_models.MaxCols - 1
    frmedmod.grid_models.Col = xx
    frmedmod.grid_models.Text = frm_1model.txt_modvar(xx - 1).Text
  Next xx

  frmedmod.grid_models.Col = frmedmod.grid_models.MaxCols
  frmedmod.grid_models.Text = "NA"

  If (frmedmod.m_addProp = True) Then
    frmedmod.m_addProp = False
    frmedmod.numprops.Text = frmedmod.numprops.Text + 1
  End If
  
  frmedmod.fillproplist
  Unload frm_1model
End Sub

Private Sub cmd_cancel_Click()
  
  Select Case (frm_1model.modtypex)
    Case 1          'pls model
      unity_main.errorstring = "GRAMS PLSIQ Model Property Configuration screen Cancel button selected"
    Case 2          'mlr model
      unity_main.errorstring = "MLR Model Property Configuration screen Cancel button selected"
    Case 3          'secondary model
      unity_main.errorstring = "Secondary Model Property Configuration screen Cancel button selected"
  End Select
  
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  If (frmedmod.m_addProp = True) Then
    frmedmod.m_addProp = False
    frmedmod.grid_models.MaxRows = frmedmod.grid_models.MaxRows - 1
  End If
  
  Unload frm_1model
End Sub

Private Sub cmd_browse_Click()
  Dim dialog As New clsCommDialogs
  Dim fileDir As String
  Dim sFilter As String
  Dim dlgTitle As String
  Dim fileName As String
  
  Select Case (frm_1model.modtypex)
    Case 1          'pls model
      unity_main.errorstring = "GRAMS PLSIQ Model Property Configuration screen Browse button selected"
    Case 2          'mlr model
      unity_main.errorstring = "MLR Model Property Configuration screen Browse button selected"
    Case 3          'secondary model
      unity_main.errorstring = "Secondary Model Property Configuration screen Browse button selected"
  End Select
  
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  On Error GoTo BAD_FILE
  dialog.InitDialogs
  fileDir = MODELS_DIR

  If (frm_1model.modtypex < 1) Then
    sFilter = ("PLSIQ (*" & GRAMS_MODEL_FILE_EXT & ")" & Chr(0) & "*" & GRAMS_MODEL_FILE_EXT & Chr(0) & "MLR (*" & MLR_MODEL_FILE_EXT & ")" & Chr(0) & "*" & MLR_MODEL_FILE_EXT & Chr(0) & "Secondary (*" & SEC_MODEL_FILE_EXT & ")" & Chr(0) & "*" & SEC_MODEL_FILE_EXT & Chr(0))
  End If
  
  Select Case frm_1model.modtypex
    Case 1
      sFilter = ("PLSIQ (*" & GRAMS_MODEL_FILE_EXT & ")" & Chr(0) + "*" & GRAMS_MODEL_FILE_EXT & Chr(0))
    Case 2
      sFilter = ("MLR (*" & MLR_MODEL_FILE_EXT & ")" & Chr(0) + "*" & MLR_MODEL_FILE_EXT & Chr(0))
    Case 3
      sFilter = ("Secondary (*" & SEC_MODEL_FILE_EXT & ")" & Chr(0) + "*" & SEC_MODEL_FILE_EXT & Chr(0))
  End Select
  
  dlgTitle = MLSupport.GSS("frm_1model", "dlgTitle", "Select Model File")
  fileName = dialog.ShowOpen(Me.hwnd, sFilter, dlgTitle, OFN_FILEMUSTEXIST + OFN_HIDEREADONLY + OFN_NOCHANGEDIR, fileDir)

  If (fileName <> "") Then
    frm_1model.txt_modvar(1).Text = fileName
  End If

  Exit Sub

BAD_FILE:
  CWrap.ShowMessageBoxW MLSupport.GSS("frm_1model", "errMsg2", "Error selecting file, please confirm you selected a valid file"), vbCritical
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub

Private Sub Picture1_Click()
  
  unity_main.formfrom = 3
  unity_main.varfrom = 0
  frm_kybd.lbl_kybd.Caption = frm_1model.Label1(0).Caption
  frm_kybd.txt_kybd.Text = txt_modvar(0).Text
  frm_kybd.Show 1
End Sub

Private Sub Picture2_Click()
  
  unity_main.formfrom = 3
  unity_main.varfrom = 1
  frm_kybd.lbl_kybd.Caption = frm_1model.Label1(1).Caption
  frm_kybd.txt_kybd.Text = txt_modvar(1).Text
  frm_kybd.Show 1
End Sub

Private Sub Picture3_Click()
  
  unity_main.formfrom = 3
  unity_main.varfrom = 2
  frm_numpad.lbl_num.Caption = frm_1model.Label1(2).Caption
  frm_numpad.txt_num.Text = txt_modvar(2).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture4_Click()
  
  unity_main.formfrom = 3
  unity_main.varfrom = 3
  frm_numpad.lbl_num.Caption = frm_1model.Label1(3).Caption
  frm_numpad.txt_num.Text = txt_modvar(3).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture5_Click()
  
  unity_main.formfrom = 3
  unity_main.varfrom = 4
  frm_numpad.lbl_num.Caption = frm_1model.Label1(4).Caption
  frm_numpad.txt_num.Text = txt_modvar(4).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture6_Click()
  
  unity_main.formfrom = 3
  unity_main.varfrom = 5
  frm_numpad.lbl_num.Caption = frm_1model.Label1(5).Caption
  frm_numpad.txt_num.Text = txt_modvar(5).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture7_Click()
  
  unity_main.formfrom = 3
  unity_main.varfrom = 6
  frm_numpad.lbl_num.Caption = frm_1model.Label1(6).Caption
  frm_numpad.txt_num.Text = txt_modvar(6).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture8_Click()
  
  unity_main.formfrom = 3
  unity_main.varfrom = 7
  frm_numpad.lbl_num.Caption = frm_1model.Label1(7).Caption
  frm_numpad.txt_num.Text = txt_modvar(7).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture9_Click()
  
  unity_main.formfrom = 3
  unity_main.varfrom = 8
  frm_numpad.lbl_num.Caption = frm_1model.Label1(8).Caption
  frm_numpad.txt_num.Text = txt_modvar(8).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture10_Click()
  
  unity_main.formfrom = 3
  unity_main.varfrom = 9
  frm_numpad.lbl_num.Caption = frm_1model.Label1(9).Caption
  frm_numpad.txt_num.Text = txt_modvar(9).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture11_Click()
  
  unity_main.formfrom = 3
  unity_main.varfrom = 10
  frm_numpad.lbl_num.Caption = frm_1model.Label1(10).Caption
  frm_numpad.txt_num.Text = txt_modvar(10).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture12_Click()
  
  unity_main.formfrom = 3
  unity_main.varfrom = 11
  frm_numpad.lbl_num.Caption = frm_1model.Label1(11).Caption
  frm_numpad.txt_num.Text = txt_modvar(11).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture13_Click()
  
  unity_main.formfrom = 3
  unity_main.varfrom = 12
  frm_numpad.lbl_num.Caption = frm_1model.Label1(12).Caption
  frm_numpad.txt_num.Text = txt_modvar(12).Text
  frm_numpad.Show 1
End Sub

Private Sub Picture14_Click()
  
  unity_main.formfrom = 3
  unity_main.varfrom = 13
  frm_numpad.lbl_num.Caption = frm_1model.Label1(13).Caption
  frm_numpad.txt_num.Text = txt_modvar(13).Text
  frm_numpad.Show 1
End Sub








