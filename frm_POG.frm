VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frm_POG 
   Caption         =   "LIMS Configuration"
   ClientHeight    =   10560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13380
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
   Icon            =   "frm_POG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10560
   ScaleWidth      =   13380
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm MSComm2 
      Left            =   4320
      Top             =   9240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin HexUniControls.ctlUniComboBoxXP combo_flowctrl 
      Height          =   420
      Left            =   10800
      TabIndex        =   41
      Top             =   4680
      Width           =   2000
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
      Tip             =   "frm_POG.frx":030A
      Sorted          =   0   'False
      HScroll         =   0   'False
      Style           =   3
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
      MouseIcon       =   "frm_POG.frx":032A
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniComboBoxXP combo_stopbits 
      Height          =   420
      Left            =   10800
      TabIndex        =   40
      Top             =   4080
      Width           =   2000
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
      Tip             =   "frm_POG.frx":0346
      Sorted          =   0   'False
      HScroll         =   0   'False
      Style           =   3
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
      MouseIcon       =   "frm_POG.frx":0366
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniComboBoxXP combo_parity 
      Height          =   420
      Left            =   10800
      TabIndex        =   39
      Top             =   3480
      Width           =   2000
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
      Tip             =   "frm_POG.frx":0382
      Sorted          =   0   'False
      HScroll         =   0   'False
      Style           =   3
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
      MouseIcon       =   "frm_POG.frx":03A2
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniComboBoxXP combo_databits 
      Height          =   420
      Left            =   10800
      TabIndex        =   38
      Top             =   2880
      Width           =   2000
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
      Tip             =   "frm_POG.frx":03BE
      Sorted          =   0   'False
      HScroll         =   0   'False
      Style           =   3
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
      MouseIcon       =   "frm_POG.frx":03DE
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniComboBoxXP combo_bps 
      Height          =   420
      Left            =   10800
      TabIndex        =   37
      Top             =   2280
      Width           =   2000
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
      Tip             =   "frm_POG.frx":03FA
      Sorted          =   0   'False
      HScroll         =   0   'False
      Style           =   3
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
      MouseIcon       =   "frm_POG.frx":041A
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniComboBoxXP combo_delim 
      Height          =   420
      Left            =   4080
      TabIndex        =   24
      Top             =   4920
      Width           =   1995
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
      Tip             =   "frm_POG.frx":0436
      Sorted          =   0   'False
      HScroll         =   0   'False
      Style           =   3
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
      MouseIcon       =   "frm_POG.frx":0456
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniComboBoxXP combo_startchar 
      Height          =   420
      Left            =   4080
      TabIndex        =   23
      Top             =   4080
      Width           =   1995
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
      Tip             =   "frm_POG.frx":0472
      Sorted          =   0   'False
      HScroll         =   0   'False
      Style           =   3
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
      MouseIcon       =   "frm_POG.frx":0492
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniCheckXP chk_propOutlier 
      Height          =   300
      Left            =   240
      TabIndex        =   19
      Top             =   8445
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":04AE
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":04EE
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":050E
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlNumIncXP numInc_commPort 
      Height          =   600
      Left            =   10800
      TabIndex        =   36
      Top             =   1560
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
      Text            =   "1"
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
      MouseIcon       =   "frm_POG.frx":052A
      TrapTabKey      =   0   'False
   End
   Begin VB.PictureBox MSComm1 
      Height          =   480
      Left            =   9000
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   50
      Top             =   5880
      Width           =   1200
   End
   Begin HexUniControls.ctlUniCheckXP chk_intercept 
      Height          =   300
      Left            =   240
      TabIndex        =   21
      Top             =   9240
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":0546
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":058A
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":05AA
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniCheckXP chk_sresid 
      Height          =   300
      Left            =   240
      TabIndex        =   18
      Top             =   8040
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":05C6
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":060C
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":062C
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniCheckXP chk_mdist 
      Height          =   300
      Left            =   240
      TabIndex        =   17
      Top             =   7635
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":0648
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":068E
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":06AE
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniCheckXP chk_datetime 
      Height          =   300
      Left            =   240
      TabIndex        =   9
      Top             =   4035
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":06CA
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":0702
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":0722
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniCheckXP chk_appendornew 
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "frm_POG.frx":073E
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":0770
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":0790
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniCheckXP chk_acceptreject 
      Height          =   330
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":07AC
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":0804
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":0824
      ShowFocus       =   -1  'True
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   10320
      Picture         =   "frm_POG.frx":0840
      ScaleHeight     =   420
      ScaleWidth      =   465
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   9600
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   10320
      Picture         =   "frm_POG.frx":0C82
      ScaleHeight     =   420
      ScaleWidth      =   465
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   9000
      Width           =   495
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_limsinpath 
      Height          =   450
      Left            =   5880
      TabIndex        =   35
      Top             =   9600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   794
      BackColor       =   -2147483643
      ForeColor       =   12582912
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_POG.frx":10C4
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
      Tip             =   "frm_POG.frx":10E4
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":1104
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_limsin 
      Height          =   450
      Left            =   5880
      TabIndex        =   34
      Top             =   9000
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   794
      BackColor       =   -2147483643
      ForeColor       =   12582912
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_POG.frx":1120
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
      Tip             =   "frm_POG.frx":1140
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":1160
   End
   Begin HexUniControls.ctlUniCheckXP chk_limsin 
      Height          =   375
      Left            =   5880
      TabIndex        =   33
      Top             =   8400
      Width           =   7095
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
      Caption         =   "frm_POG.frx":117C
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":11EA
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":120A
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniCheckXP chk_sampid 
      Height          =   300
      Left            =   240
      TabIndex        =   12
      Top             =   5640
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":1226
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":1258
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":1278
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniFrameXP Frame1 
      Height          =   2415
      Left            =   4080
      Top             =   5640
      Width           =   4095
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
      Caption         =   "frm_POG.frx":1294
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_POG.frx":12E0
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":1300
      Begin HexUniControls.ctlUniCheckXP chk_input8 
         Height          =   300
         Left            =   2200
         TabIndex        =   32
         Top             =   1900
         Width           =   1800
         _ExtentX        =   3175
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
         Caption         =   "frm_POG.frx":131C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_POG.frx":134A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_POG.frx":136A
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input7 
         Height          =   300
         Left            =   2200
         TabIndex        =   31
         Top             =   1440
         Width           =   1800
         _ExtentX        =   3175
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
         Caption         =   "frm_POG.frx":1386
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_POG.frx":13B4
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_POG.frx":13D4
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input6 
         Height          =   300
         Left            =   2200
         TabIndex        =   30
         Top             =   900
         Width           =   1800
         _ExtentX        =   3175
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
         Caption         =   "frm_POG.frx":13F0
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_POG.frx":141E
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_POG.frx":143E
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input5 
         Height          =   300
         Left            =   2200
         TabIndex        =   29
         Top             =   400
         Width           =   1800
         _ExtentX        =   3175
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
         Caption         =   "frm_POG.frx":145A
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_POG.frx":1488
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_POG.frx":14A8
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input4 
         Height          =   300
         Left            =   300
         TabIndex        =   28
         Top             =   1900
         Width           =   1800
         _ExtentX        =   3175
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
         Caption         =   "frm_POG.frx":14C4
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_POG.frx":14F2
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_POG.frx":1512
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input3 
         Height          =   300
         Left            =   300
         TabIndex        =   27
         Top             =   1400
         Width           =   1800
         _ExtentX        =   3175
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
         Caption         =   "frm_POG.frx":152E
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_POG.frx":155C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_POG.frx":157C
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input2 
         Height          =   300
         Left            =   300
         TabIndex        =   26
         Top             =   900
         Width           =   1800
         _ExtentX        =   3175
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
         Caption         =   "frm_POG.frx":1598
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_POG.frx":15C6
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_POG.frx":15E6
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniCheckXP chk_input1 
         Height          =   300
         Left            =   300
         TabIndex        =   25
         Top             =   400
         Width           =   1800
         _ExtentX        =   3175
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
         Caption         =   "frm_POG.frx":1602
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   0
         Pressed         =   0   'False
         Tip             =   "frm_POG.frx":1630
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_POG.frx":1650
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   5640
      Picture         =   "frm_POG.frx":166C
      ScaleHeight     =   420
      ScaleWidth      =   465
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1200
      Width           =   495
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_pogfilename 
      Height          =   450
      Left            =   2880
      TabIndex        =   4
      Top             =   1200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
      BackColor       =   -2147483643
      ForeColor       =   0
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_POG.frx":1AAE
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
      Tip             =   "frm_POG.frx":1ADE
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":1AFE
   End
   Begin HexUniControls.ctlUniCheckXP chk_comm 
      Height          =   330
      Left            =   4560
      TabIndex        =   2
      Top             =   195
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":1B1A
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":1B62
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":1B82
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniCheckXP chk_file 
      Height          =   330
      Left            =   240
      TabIndex        =   1
      Top             =   195
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":1B9E
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":1BE0
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":1C00
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniCheckXP chk_header 
      Height          =   300
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":1C1C
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":1C50
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":1C70
      ShowFocus       =   -1  'True
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   5640
      Picture         =   "frm_POG.frx":1C8C
      ScaleHeight     =   420
      ScaleWidth      =   465
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   10920
      TabIndex        =   0
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
      Caption         =   "frm_POG.frx":20CE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":20FA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":211A
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   650
      Left            =   8640
      TabIndex        =   42
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
      Caption         =   "frm_POG.frx":2136
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":216E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":218E
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_outpath 
      Height          =   450
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   794
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_POG.frx":21AA
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
      Tip             =   "frm_POG.frx":21E6
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":2206
   End
   Begin HexUniControls.ctlUniCheckXP chk_startchar 
      Height          =   300
      Left            =   240
      TabIndex        =   8
      Top             =   3645
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":2222
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":2260
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":2280
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniCheckXP chk_propvalue 
      Height          =   300
      Left            =   240
      TabIndex        =   16
      Top             =   7245
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":229C
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":22D8
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":22F8
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniCheckXP chk_propname 
      Height          =   300
      Left            =   240
      TabIndex        =   15
      Top             =   6840
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":2314
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":234E
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":236E
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniCheckXP chk_product 
      Height          =   300
      Left            =   240
      TabIndex        =   14
      Top             =   6435
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":238A
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":23B8
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":23D8
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniCheckXP chk_comment 
      Height          =   300
      Left            =   240
      TabIndex        =   13
      Top             =   6045
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":23F4
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":2430
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":2450
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniCheckXP chk_time 
      Height          =   300
      Left            =   240
      TabIndex        =   11
      Top             =   4845
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":246C
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":2494
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":24B4
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniCheckXP chk_date 
      Height          =   300
      Left            =   240
      TabIndex        =   10
      Top             =   4440
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":24D0
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":24F8
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":2518
      ShowFocus       =   -1  'True
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9960
      Top             =   5880
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   10560
      FormDesignWidth =   13380
   End
   Begin HexUniControls.ctlUniCheckXP chk_slope 
      Height          =   300
      Left            =   240
      TabIndex        =   22
      Top             =   9645
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":2534
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":2570
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":2590
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniLabel Label15 
      Height          =   375
      Left            =   4080
      Top             =   3645
      Width           =   2220
      _ExtentX        =   3916
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
      Caption         =   "frm_POG.frx":25AC
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_POG.frx":25EA
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":260A
   End
   Begin HexUniControls.ctlUniLabel Label14 
      Height          =   375
      Left            =   10920
      Top             =   9600
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "frm_POG.frx":2626
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_POG.frx":2658
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":2678
   End
   Begin HexUniControls.ctlUniLabel Label13 
      Height          =   375
      Left            =   10920
      Top             =   9000
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "frm_POG.frx":2694
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_POG.frx":26C6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":26E6
   End
   Begin HexUniControls.ctlUniLabel Label11 
      Height          =   735
      Left            =   8520
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":2702
      BackColor       =   -2147483633
      ForeColor       =   12582912
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_POG.frx":275C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":277C
   End
   Begin HexUniControls.ctlUniLabel Label12 
      Height          =   375
      Left            =   720
      Top             =   1200
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
      Caption         =   "frm_POG.frx":2798
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_POG.frx":27D8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":27F8
   End
   Begin HexUniControls.ctlUniLabel Label10 
      Height          =   420
      Left            =   8640
      Top             =   4680
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":2814
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_POG.frx":284C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":286C
   End
   Begin HexUniControls.ctlUniLabel Label9 
      Height          =   420
      Left            =   8640
      Top             =   4080
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":2888
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_POG.frx":28BA
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":28DA
   End
   Begin HexUniControls.ctlUniLabel Label8 
      Height          =   420
      Left            =   8640
      Top             =   3480
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":28F6
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_POG.frx":2922
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":2942
   End
   Begin HexUniControls.ctlUniLabel Label7 
      Height          =   420
      Left            =   8640
      Top             =   2880
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":295E
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_POG.frx":2990
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":29B0
   End
   Begin HexUniControls.ctlUniLabel Label6 
      Height          =   420
      Left            =   8640
      Top             =   2280
      Width           =   2000
      _ExtentX        =   3519
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":29CC
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_POG.frx":29FE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":2A1E
   End
   Begin HexUniControls.ctlUniLabel Label5 
      Height          =   420
      Left            =   8640
      Top             =   1680
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":2A3A
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_POG.frx":2A6C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":2A8C
   End
   Begin HexUniControls.ctlUniLabel Label4 
      Height          =   405
      Left            =   9000
      Top             =   1000
      Width           =   3555
      _ExtentX        =   6271
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
      Caption         =   "frm_POG.frx":2AA8
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_POG.frx":2AF8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":2B18
   End
   Begin HexUniControls.ctlUniLabel Label2 
      Height          =   375
      Left            =   240
      Top             =   2760
      Width           =   7455
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
      Caption         =   "frm_POG.frx":2B34
      BackColor       =   -2147483633
      ForeColor       =   12582912
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_POG.frx":2BAC
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":2BCC
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Left            =   4080
      Top             =   4560
      Width           =   2220
      _ExtentX        =   3916
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
      Caption         =   "frm_POG.frx":2BE8
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_POG.frx":2C1A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":2C3A
   End
   Begin HexUniControls.ctlUniCheckXP chk_nd 
      Height          =   300
      Left            =   240
      TabIndex        =   20
      Top             =   8835
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":2C56
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":2C8C
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":2CAC
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniCheckXP chk_serialNum 
      Height          =   300
      Left            =   240
      TabIndex        =   47
      Top             =   5235
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":2CC8
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":2D10
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":2D30
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniRadioXP opt_asciiFormat 
      Height          =   300
      Left            =   5400
      TabIndex        =   48
      Top             =   1800
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":2D4C
      Enabled         =   -1  'True
      Align           =   0
      RadioBackColor  =   -2147483643
      RadioForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":2D76
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":2D96
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniRadioXP opt_uniFormat 
      Height          =   300
      Left            =   5400
      TabIndex        =   49
      Top             =   2280
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_POG.frx":2DB2
      Enabled         =   -1  'True
      Align           =   0
      RadioBackColor  =   -2147483643
      RadioForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Pressed         =   0   'False
      Tip             =   "frm_POG.frx":2DE0
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_POG.frx":2E00
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   11160
      Top             =   6000
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
      Left            =   10560
      Top             =   5880
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_POG.frx":2E1C
   End
   Begin VB.Line Line3 
      X1              =   5760
      X2              =   5760
      Y1              =   8160
      Y2              =   10560
   End
   Begin VB.Line Line2 
      X1              =   13440
      X2              =   5760
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line1 
      X1              =   8400
      X2              =   8400
      Y1              =   120
      Y2              =   8160
   End
End
Attribute VB_Name = "frm_POG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Configuration file parameters
Private pog_usefile As Integer
Private pogcomm As Integer
Private pogheader As Integer
Private pogpath As String
Private pogfile As String
Private pogappendornew As Integer
Private pogacceptreject As Integer
Private pogformat As String
Private pogusestart As Integer
Private pogstartchar As String
Private pogdelimtype As String
Private pogdelimchar As String
Private pogdatetime As Integer
Private pogdate As Integer
Private pogtime As Integer
Private pogSerialNum As Integer
Private pogcomment As Integer
Private pogproduct As Integer
Private pogmodelid As Integer
Private pogpropname As Integer
Private pogpropvalue As Integer
Private pogpropoutlier As Integer
Private pogmdist As Integer
Private pogsresid As Integer
Private pognd As Integer
Private pogintercept As Integer
Private pogslope As Integer
Private pogport As Integer
Private pogbps As Long
Private pogdatabits As Integer
Private pogstopbits As Single
Private pogflowctrl As String
Private pogparity As String
Private realdelim As String
Private pogsampid As Integer
Private remoteproduct As Integer
Private limsinpath As String
Private limsinfile As String
Private m_limsUserInputs(1 To MAX_MAN_INPUTS) As Boolean

' Misc variables
Private m_badLIMSIniVal As Boolean
Private passstringd As String
Private writedata As String
Private writeheader As String

Public Sub writepog()
  Dim fullfile As String
  Dim pogexists As Boolean
  Dim headerstring As String 'header if needed
  Dim datastring As String  'delimited string to write out
  Dim nprop As Integer
  Dim fhandle As Integer
  Dim zz As Integer
  Dim tempint, jj As Integer
  Dim numValues As Integer
  Dim valueNames() As String
  Dim Values() As String
  Dim uniMsg As String
  
  uniMsg = MLSupport.GSS("frm_POG", "statMsg1", "Building LIMS output")
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Building LIMS output", uniMsg)

  ' Check if to include start char
  If (unity_main.pogusestart = 1) Then
    headerstring = unity_main.pogstartchar
    datastring = unity_main.pogstartchar
  Else
    headerstring = ""
    datastring = ""
  End If
    
  ' Check if to include Date-Time (military format)
  If (unity_main.pogdatetime = 1) Then
    numValues = numValues + 1
    ReDim Preserve valueNames(1 To numValues)
    ReDim Preserve Values(1 To numValues)
    
    passstringd = MLSupport.GSS("Headers", "dateTime", "Date_Time")
    valueNames(numValues) = passstringd
    
    If (unity_main.pogdelimchar = "ff") Then
      padlen
      headerstring = headerstring & passstringd
    Else
      headerstring = headerstring & passstringd & unity_main.pogdelimchar
    End If
        
    passstringd = unity_main.lbl_miltime.Caption
    Values(numValues) = passstringd
        
    If (unity_main.pogdelimchar = "ff") Then
      padlen
      datastring = datastring & passstringd
    Else
      datastring = datastring & passstringd & unity_main.pogdelimchar
    End If
  End If
    
  ' Check if to include Date
  If (unity_main.pogdate = 1) Then
    numValues = numValues + 1
    ReDim Preserve valueNames(1 To numValues)
    ReDim Preserve Values(1 To numValues)
    
    passstringd = MLSupport.GSS("Headers", "date", "Date")
    valueNames(numValues) = passstringd
        
    If (unity_main.pogdelimchar = "ff") Then
      padlen
      headerstring = headerstring & passstringd
    Else
      headerstring = headerstring & passstringd & unity_main.pogdelimchar
    End If
    
    passstringd = unity_main.lbl_date.Caption
    Values(numValues) = passstringd
    
    If (unity_main.pogdelimchar = "ff") Then
      padlen
      datastring = datastring & passstringd
    Else
      datastring = datastring & passstringd & unity_main.pogdelimchar
    End If
  End If
    
  ' Check if to include Time
  If (unity_main.pogtime = 1) Then
    numValues = numValues + 1
    ReDim Preserve valueNames(1 To numValues)
    ReDim Preserve Values(1 To numValues)
    
    passstringd = MLSupport.GSS("Headers", "time", "Time")
    valueNames(numValues) = passstringd

    If (unity_main.pogdelimchar = "ff") Then
      padlen
      headerstring = headerstring & passstringd
    Else
      headerstring = headerstring & passstringd & unity_main.pogdelimchar
    End If
    
    passstringd = unity_main.lbl_time.Caption
    Values(numValues) = passstringd
    
    If (unity_main.pogdelimchar = "ff") Then
      padlen
      datastring = datastring & passstringd
    Else
      datastring = datastring & passstringd & unity_main.pogdelimchar
    End If
  End If
  
  ' Check if to include System Serial Number
  If (unity_main.pogSerialNum = 1) Then
    numValues = numValues + 1
    ReDim Preserve valueNames(1 To numValues)
    ReDim Preserve Values(1 To numValues)
    
    passstringd = MLSupport.GSS("Headers", "serNum", "Serial No.")
    valueNames(numValues) = passstringd

    If (unity_main.pogdelimchar = "ff") Then
      padlen
      headerstring = headerstring & passstringd
    Else
      headerstring = headerstring & passstringd & unity_main.pogdelimchar
    End If
    
    passstringd = unity_main.m_sysSerialNum
    Values(numValues) = passstringd
    
    If (unity_main.pogdelimchar = "ff") Then
      padlen
      datastring = datastring & passstringd
    Else
      datastring = datastring & passstringd & unity_main.pogdelimchar
    End If
  End If
  
  ' Check if to include Sample ID
  If (unity_main.pogsampid = 1) Then
    numValues = numValues + 1
    ReDim Preserve valueNames(1 To numValues)
    ReDim Preserve Values(1 To numValues)
    
    passstringd = MLSupport.GSS("Headers", "sampleID", "Sample ID")
    valueNames(numValues) = passstringd
    
    If (unity_main.pogdelimchar = "ff") Then
      padlen
      headerstring = headerstring & passstringd
    Else
      headerstring = headerstring & passstringd & unity_main.pogdelimchar
    End If
    
    passstringd = unity_main.txtsamplename.Text
    Values(numValues) = passstringd
    
    If (unity_main.pogdelimchar = "ff") Then
      padlen
      datastring = datastring & passstringd
    Else
      datastring = datastring & passstringd & unity_main.pogdelimchar
    End If
  End If

  ' Check if to include sample comment
  If (unity_main.pogcomment = 1) Then
    numValues = numValues + 1
    ReDim Preserve valueNames(1 To numValues)
    ReDim Preserve Values(1 To numValues)
    
    passstringd = MLSupport.GSS("Headers", "comment", "Comment")
    valueNames(numValues) = passstringd

    If (unity_main.pogdelimchar = "ff") Then
      padlen
      headerstring = headerstring & passstringd
    Else
      headerstring = headerstring & passstringd & unity_main.pogdelimchar
    End If
        
    passstringd = unity_main.txtsampcomment.Text
    Values(numValues) = passstringd
    
    If (unity_main.pogdelimchar = "ff") Then
      padlen
      datastring = datastring & passstringd
    Else
      datastring = datastring & passstringd & unity_main.pogdelimchar
    End If
  End If

  ' Check if to include product name
  If (unity_main.pogproduct = 1) Then
    numValues = numValues + 1
    ReDim Preserve valueNames(1 To numValues)
    ReDim Preserve Values(1 To numValues)
    
    passstringd = MLSupport.GSS("Headers", "product", "Product")
    valueNames(numValues) = passstringd
        
    If (unity_main.pogdelimchar = "ff") Then
      padlen
      headerstring = headerstring & passstringd
    Else
      headerstring = headerstring & passstringd & unity_main.pogdelimchar
    End If
    
    passstringd = unity_main.lblProd1.Caption
    Values(numValues) = passstringd
    
    If (unity_main.pogdelimchar = "ff") Then
      padlen
      datastring = datastring & passstringd
    Else
      datastring = datastring & passstringd & unity_main.pogdelimchar
    End If
  End If
    
  ' Check if to include inputs 1 - 8
  For zz = 1 To MAX_MAN_INPUTS
    If (LIMSUserInputs(zz) = True) Then
      numValues = numValues + 1
      ReDim Preserve valueNames(1 To numValues)
      ReDim Preserve Values(1 To numValues)
    
      passstringd = Trim(frm_scanname.lbl(zz).Caption)
    
      If Trim(passstringd) = "" Then
        passstringd = MLSupport.GSS("Headers", "input", "Input") & " " & CStr(zz)
      End If
      
      valueNames(numValues) = passstringd
      
      If (unity_main.pogdelimchar = "ff") Then
        padlen
        headerstring = headerstring & passstringd
      Else
        headerstring = headerstring & passstringd & unity_main.pogdelimchar
      End If
            
      ' Setup for input enable field
      frm_buttoncfg.ss_buttonconfig.Col = zz
      frm_buttoncfg.ss_buttonconfig.Row = 1
  
      ' Check if input enabled
      If (unity_main.m_useMIV = True) And (frm_buttoncfg.ss_buttonconfig.Value = 1) Then
        ' Setup for text entry/list box selection field
        frm_buttoncfg.ss_buttonconfig.Col = zz
        frm_buttoncfg.ss_buttonconfig.Row = 2
  
        ' Check if using text entry
        If (frm_buttoncfg.ss_buttonconfig.Value = 0) Then
          passstringd = Trim(frm_scanname.txtbx(zz).Text)
        Else    ' Using list
          passstringd = Trim(frm_scanname.combo(zz).Text)
        End If
      Else
        passstringd = MLSupport.GSS("Headers", "na", "NA")
      End If
    
      Values(numValues) = passstringd
    
      If (unity_main.pogdelimchar = "ff") Then
        padlen
        datastring = datastring & passstringd
      Else
        datastring = datastring & passstringd & unity_main.pogdelimchar
      End If
    End If
  Next zz
    
  'properties start at col 14.  row 1 = property, row2=value, col 15 = prop2 in row 1, prop2 value in row 2.
  'row1 = propname
  'row2 = value (prop)
  tempint = frmedmod.numprops.Text
    
  ' Check if to include the various propery parameters
  For jj = 1 To tempint
    If (unity_main.pogpropname = 1) Then                                   'prop name
      numValues = numValues + 1
      ReDim Preserve valueNames(1 To numValues)
      ReDim Preserve Values(1 To numValues)
    
      passstringd = MLSupport.GSS("Headers", "property", "Property") & " " & CStr(jj)
      valueNames(numValues) = passstringd
      
      If (unity_main.pogdelimchar = "ff") Then
        padlen
        headerstring = headerstring & passstringd
      Else
        headerstring = headerstring & passstringd & unity_main.pogdelimchar
      End If
      
      unity_main.fpspread_pred.Row = jj
      unity_main.fpspread_pred.Col = 1
      passstringd = unity_main.fpspread_pred.Text
      Values(numValues) = passstringd
      
      If (unity_main.pogdelimchar = "ff") Then
        padlen
        datastring = datastring & passstringd
      Else
        datastring = datastring & passstringd & unity_main.pogdelimchar
      End If
    End If
    
    If (unity_main.pogpropvalue = 1) Then                                 'Property value
      numValues = numValues + 1
      ReDim Preserve valueNames(1 To numValues)
      ReDim Preserve Values(1 To numValues)
    
      passstringd = MLSupport.GSS("Headers", "value", "Value") & " " & CStr(jj)
      valueNames(numValues) = passstringd
      
      If (unity_main.pogdelimchar = "ff") Then
        padlen
        headerstring = headerstring & passstringd
      Else
        headerstring = headerstring & passstringd & unity_main.pogdelimchar
      End If
      
      unity_main.fpspread_pred.Row = jj
      unity_main.fpspread_pred.Col = 2
      passstringd = unity_main.fpspread_pred.Text
      Values(numValues) = passstringd
      
      If (unity_main.pogdelimchar = "ff") Then
        padlen
        datastring = datastring & passstringd
      Else
        datastring = datastring & passstringd & unity_main.pogdelimchar
      End If
    End If
        
    nprop = frmedmod.numprops.Text
        
    If (unity_main.pogmdist = 1) Then                                   'm-dist
      numValues = numValues + 1
      ReDim Preserve valueNames(1 To numValues)
      ReDim Preserve Values(1 To numValues)
    
      passstringd = MLSupport.GSS("Headers", "mDist", "M-Dist") & " " & CStr(jj)
      valueNames(numValues) = passstringd
      
      If (unity_main.pogdelimchar = "ff") Then
        padlen
        headerstring = headerstring & passstringd
      Else
        headerstring = headerstring & passstringd & unity_main.pogdelimchar
      End If
      
      ' Check if no prediction made
      If (unity_main.lstmd.ListCount = 0) Then
        passstringd = unity_main.m_noOLVal
      Else
        passstringd = unity_main.lstmd.List(jj - 1)
      End If
      
      Values(numValues) = passstringd
            
      If (unity_main.pogdelimchar = "ff") Then
        padlen
        datastring = datastring & passstringd
      Else
        datastring = datastring & passstringd & unity_main.pogdelimchar
      End If
    End If
    
    If (unity_main.pogsresid = 1) Then                                   'resid
      numValues = numValues + 1
      ReDim Preserve valueNames(1 To numValues)
      ReDim Preserve Values(1 To numValues)
    
      passstringd = MLSupport.GSS("Headers", "sResid", "S-Resid") & " " & CStr(jj)
      valueNames(numValues) = passstringd
      
      If (unity_main.pogdelimchar = "ff") Then
        padlen
        headerstring = headerstring & passstringd
      Else
        headerstring = headerstring & passstringd & unity_main.pogdelimchar
      End If
      
      ' Check if no prediction made
      If (unity_main.lstresrat.ListCount = 0) Then
        passstringd = unity_main.m_noOLVal
      Else
        passstringd = unity_main.lstresrat.List(jj - 1)
      End If
      
      Values(numValues) = passstringd
 
      If (unity_main.pogdelimchar = "ff") Then
        padlen
        datastring = datastring & passstringd
      Else
        datastring = datastring & passstringd & unity_main.pogdelimchar
      End If
    End If
        
    If (unity_main.pogpropoutlier = 1) Then                        'Property outlier
      numValues = numValues + 1
      ReDim Preserve valueNames(1 To numValues)
      ReDim Preserve Values(1 To numValues)
    
      passstringd = MLSupport.GSS("Header", "outlier", "Outlier") & " " & CStr(jj)
      valueNames(numValues) = passstringd

      If (unity_main.pogdelimchar = "ff") Then
        padlen
        headerstring = headerstring & passstringd
      Else
        headerstring = headerstring & passstringd & unity_main.pogdelimchar
      End If
            
      ' Check if no prediction made
      If (unity_main.lst_qual.ListCount = 0) Then
        passstringd = unity_main.m_noOLVal
      Else
        passstringd = unity_main.lst_qual.List(jj - 1)
      End If
      
      Values(numValues) = passstringd
            
      If (unity_main.pogdelimchar = "ff") Then
        padlen
        datastring = datastring & passstringd
      Else
        datastring = datastring & passstringd & unity_main.pogdelimchar
      End If
    End If
        
    If (unity_main.pognd = 1) Then                                   ' neighborhood distance
      numValues = numValues + 1
      ReDim Preserve valueNames(1 To numValues)
      ReDim Preserve Values(1 To numValues)
    
      passstringd = MLSupport.GSS("Headers", "nD", "ND") & " " & CStr(jj)
      valueNames(numValues) = passstringd
      
      If (unity_main.pogdelimchar = "ff") Then
        padlen
        headerstring = headerstring & passstringd
      Else
        headerstring = headerstring & passstringd & unity_main.pogdelimchar
      End If
      
      ' Check if no prediction made
      If (unity_main.lst_nd.ListCount = 0) Then
        passstringd = unity_main.m_noOLVal
      Else
        passstringd = unity_main.lst_nd.List(jj - 1)
      End If
      
      Values(numValues) = passstringd
            
      If (unity_main.pogdelimchar = "ff") Then
        padlen
        datastring = datastring & passstringd
      Else
        datastring = datastring & passstringd & unity_main.pogdelimchar
      End If
    End If
    
    If (unity_main.pogintercept = 1) Then                 'intercept bias
      numValues = numValues + 1
      ReDim Preserve valueNames(1 To numValues)
      ReDim Preserve Values(1 To numValues)
    
      passstringd = MLSupport.GSS("Headers", "intercept", "Intercept") & " " & CStr(jj)
      valueNames(numValues) = passstringd
      
      If (unity_main.pogdelimchar = "ff") Then
        padlen
        headerstring = headerstring & passstringd
      Else
        headerstring = headerstring & passstringd & unity_main.pogdelimchar
      End If
      
      ' Check if no prediction made
      If (unity_main.lstint.ListCount = 0) Then
        passstringd = unity_main.m_noOLVal
      Else
        passstringd = unity_main.lstint.List(jj - 1)
      End If
      
      Values(numValues) = passstringd
      
      If (unity_main.pogdelimchar = "ff") Then
        padlen
        datastring = datastring & passstringd
      Else
        datastring = datastring & passstringd & unity_main.pogdelimchar
      End If
    End If

    If (unity_main.pogslope = 1) Then                                   'slope
      numValues = numValues + 1
      ReDim Preserve valueNames(1 To numValues)
      ReDim Preserve Values(1 To numValues)
    
      passstringd = MLSupport.GSS("Headers", "slope", "Slope") & " " & CStr(jj)
      valueNames(numValues) = passstringd
      
      If (unity_main.pogdelimchar = "ff") Then
        padlen
        headerstring = headerstring & passstringd
      Else
        headerstring = headerstring & passstringd & unity_main.pogdelimchar
      End If
      
      ' Check if no prediction made
      If (unity_main.lstslope.ListCount = 0) Then
        passstringd = unity_main.m_noOLVal
      Else
        passstringd = unity_main.lstslope.List(jj - 1)
      End If
      
      Values(numValues) = passstringd
      
      If (unity_main.pogdelimchar = "ff") Then
        padlen
        datastring = datastring & passstringd
      Else
        datastring = datastring & passstringd & unity_main.pogdelimchar
      End If
    End If
  Next jj

  ' Send LIMS data to any client application
  unity_main.IPCServer1.NewLIMSData 0, numValues, valueNames, Values
  
  writeheader = headerstring
  writedata = datastring
End Sub

Public Sub writeitout()
  Dim fileName As String
  Dim pogexists As Boolean
  Dim headerstring As String 'header if needed
  Dim datastring As String  'delimited string to write out
  Dim nprop As Integer
  Dim errMsg As String
  Dim uniFile As New clsUniFile
  Dim uniMsg As String
  
  headerstring = writeheader
  datastring = writedata
  
  ' Check if to write data to file
  If (pog_usefile = 1) Then
    On Error GoTo FILE_ERROR
    fileName = Trim(unity_main.pogpath) & Trim(unity_main.pogfile)
    
    uniMsg = MLSupport.GGS_Params("frm_POG.statMsg1", "Writing sample results to LIMS file: %1", fileName)
    Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Writing sample results to LIMS file: " & fileName), uniMsg)
    
    CreatePath unity_main.pogpath
    pogexists = CFile.st_FileExist(fileName)

    ' Check if file already exists and to append data
    If (pogexists = True) And (unity_main.pogappendornew = 0) Then
      If (uniFile.OpenFileAppend(fileName) = False) Then GoTo FILE_ERROR
    Else    'create new file
      If (uniFile.OpenFileWrite(fileName) = False) Then GoTo FILE_ERROR
        
      ' Check if to write out text as ASCII
      If (unity_main.pogformat = "ASCII") Then
        uniFile.WriteBOM fe_ANSI
      Else    ' Unicode
        uniFile.WriteBOM fe_UTF16LE
      End If
    End If
       
    ' Check if to write out text as ASCII
    If (unity_main.pogformat = "ASCII") Then
      ' Check if to write header to file
      If (unity_main.pogheader = 1) Then
        uniFile.WriteAnsiLine headerstring
      End If
    
      uniFile.WriteAnsiLine datastring
    Else    ' Unicode
      ' Check if to write header to file
      If (unity_main.pogheader = 1) Then
        uniFile.WriteUnicodeLine headerstring
      End If
    
      uniFile.WriteUnicodeLine datastring
    End If
    
    uniFile.Flush
    uniFile.CloseFile
  End If
  
  ' Check if transmit through serial port
  If (pogcomm = 1) Then
    Dim portOkay As Boolean
      
    ' Check if port opened already
    If (MSComm2.PortOpen = False) Then
      portOkay = initcomm(False)
    Else
      portOkay = True
    End If

    ' Check if port okay to use
    If (portOkay = True) Then
      Dim xmtBuff() As Byte
        
      uniMsg = MLSupport.GGS_Params("frm_POG.statMsg2", "Transmitting sample results to LIMS through serial port %1", CStr(unity_main.pogport))
      Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Transmitting sample results to LIMS through serial port " & unity_main.pogport), uniMsg)
      On Error GoTo CommWriteError
         
      ' Check if to transmit header to file
      If (unity_main.pogheader = 1) Then
        headerstring = headerstring & vbCr & vbLf
        
        ' Check if to send out text as ASCII
        If (unity_main.pogformat = "ASCII") Then
          xmtBuff = StrConv(headerstring, vbFromUnicode)
        Else    ' Unicode
          xmtBuff = headerstring
        End If
        
        MSComm2.Output = xmtBuff
        Wait (1)
      End If
      
      datastring = datastring & vbCr & vbLf
        
      ' Check if to send out text as ASCII
      If (unity_main.pogformat = "ASCII") Then
        xmtBuff = StrConv(datastring, vbFromUnicode)
      Else    ' Unicode
        xmtBuff = datastring
      End If
      
      MSComm2.Output = xmtBuff
    End If
  End If
  Exit Sub
    
FILE_ERROR:
  uniFile.CloseFile

  errMsg = (fileName & " file write error. " & Error$)
  unity_main.errorstring = errMsg
  unity_main.write_error
  uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", fileName, Error$)
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  Exit Sub
    
CommWriteError:
  errMsg = ("LIMS serial port " & unity_main.pogport & " transmit error. " & Error$)
  unity_main.errorstring = errMsg
  unity_main.write_error
  uniMsg = MLSupport.GGS_Params("frm_POG.errMsg3", "LIMS serial port %1 transmit error. %2", CStr(unity_main.pogport), Error$)
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub

Public Sub checknewprod()
  Dim fpath As String
  Dim fname As String
  Dim fullname As String
  Dim tempin As String
  Dim origname As String
  Dim tmpName As String
  Dim newprod As String
  Dim xx As Integer
  Dim yy As Integer
  Dim tmpFile As String
  Dim errMsg As String
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim rc As Boolean
  Dim uniMsg As String

  fpath = unity_main.limsinpath
  fname = unity_main.limsinfile
  fullname = fpath & fname
  
  ' Check if remote product file exists
  If (uniFile.st_FileExist(fullname) = True) Then
    On Error GoTo FILE_ERROR
    tmpFile = (CFG_DIR & TMP_REMOTE_PROD_FILE)
    uniFile.st_CopyFile fullname, tmpFile

    ' Open file and read only first line
    If (uniFile.OpenFileRead(tmpFile) = True) Then
      fEncoding = uniFile.ReadBOM
      
      If (fEncoding = fe_ANSI) Then
        rc = uniFile.ReadAnsiLine(tempin)
      Else
        rc = uniFile.ReadUnicodeLine(tempin)
      End If
    
      If (rc = False) Then GoTo FILE_ERROR
    
      xx = InStr(1, tempin, "=")
      yy = Len(tempin)
      newprod = Trim(Mid(tempin, xx + 1))   'new product name as string
      Call remote_product_selection(newprod, "Remote file product selection")
      uniFile.st_RmFile fullname
    Else
FILE_ERROR:
      errMsg = (fullname & " file read error. " & Error$)
      unity_main.errorstring = errMsg
      unity_main.write_error
      uniMsg = MLSupport.GGS_Params("fileErrMsg4", "%1 file read error. %2", fullname, Error$)
      uniMsg = MLSupport.GGS_Params("frm_POG.errMsg2", "%1. Continuing to use prior product", uniMsg)
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    End If
    
    uniFile.CloseFile
  
    If (tmpFile <> "") And (uniFile.st_FileExist(tmpFile) = True) Then
      uniFile.st_RmFile tmpFile
    End If
  End If
End Sub

Public Function remote_product_selection(prodName As String, remoteFunc As String) As Boolean
  Dim nprods As Integer
  Dim xx As Integer
  Dim tmpName As String

  remote_product_selection = True
  nprods = FRM_SEL_PRODUCT.LSTPRODUCTS.ListCount
    
  For xx = 0 To (nprods - 1)
    tmpName = Trim(FRM_SEL_PRODUCT.LSTPRODUCTS.List(xx))
      
    If (tmpName = prodName) Then
      unity_main.errorstring = (remoteFunc & ": " & prodName)
      unity_main.write_error (LOG_DBG_LEVEL1)
      
      If (unity_main.load_prod_file(Trim(FRM_SEL_PRODUCT.LST_INIFILE.List(xx)), True) = True) Then
        Call unity_main.save_last_product(Trim(FRM_SEL_PRODUCT.LST_INIFILE.List(xx)))
        frmedmod.fixthesize
      End If
      
      Exit Function
    End If
  Next xx

  unity_main.errorstring = (remoteFunc & " could not find product" & ": " & prodName)
  unity_main.write_error (LOG_DBG_LEVEL1)
  remote_product_selection = False
End Function

Public Sub setup_lims_lists()

  frm_POG.combo_startchar.AddItem (">")
  frm_POG.combo_startchar.AddItem (">>")
  frm_POG.combo_startchar.AddItem ("<")
  frm_POG.combo_startchar.AddItem ("/")
  frm_POG.combo_startchar.AddItem ("//")
  frm_POG.combo_startchar.AddItem ("\")
  frm_POG.combo_startchar.AddItem ("\\")
  frm_POG.combo_startchar.AddItem ("#")

  frm_POG.combo_delim.AddItem ("Tab")
  frm_POG.combo_delim.AddItem ("Comma")
  frm_POG.combo_delim.AddItem ("Fixed Field")
  frm_POG.combo_delim.AddItem ("Colon")
  frm_POG.combo_delim.AddItem ("Semicolon")
  frm_POG.combo_delim.AddItem ("Backslash")
  frm_POG.combo_delim.AddItem ("Forwardslash")

  frm_POG.combo_bps.AddItem ("1200")
  frm_POG.combo_bps.AddItem ("2400")
  frm_POG.combo_bps.AddItem ("4800")
  frm_POG.combo_bps.AddItem ("9600")
  frm_POG.combo_bps.AddItem ("19200")
  frm_POG.combo_bps.AddItem ("38400")

  frm_POG.combo_databits.AddItem ("5")
  frm_POG.combo_databits.AddItem ("6")
  frm_POG.combo_databits.AddItem ("7")
  frm_POG.combo_databits.AddItem ("8")

  frm_POG.combo_parity.AddItem ("Even")
  frm_POG.combo_parity.AddItem ("Odd")
  frm_POG.combo_parity.AddItem ("None")
  frm_POG.combo_parity.AddItem ("Mark")
  frm_POG.combo_parity.AddItem ("Space")

  frm_POG.combo_stopbits.AddItem ("1")
  frm_POG.combo_stopbits.AddItem ("1.5")
  frm_POG.combo_stopbits.AddItem ("2")

  frm_POG.combo_flowctrl.AddItem ("Xon / Xoff")
  frm_POG.combo_flowctrl.AddItem ("Hardware")
  frm_POG.combo_flowctrl.AddItem ("None")
End Sub

Public Sub load_lims(mustBeCfg As Boolean)
  Dim sourceFile As String
  Dim fileExist As Boolean
  Dim filePathName As String
  Dim uniFile As New clsUniFile
  Dim uniMsg As String
  Dim ii As Integer
  
  m_badLIMSIniVal = False
  sourceFile = (CFG_DIR & LIMS_CFG_FILE)   ' Define source file name.
  
  fileExist = uniFile.st_FileExist(sourceFile)
  
  If (fileExist = False) Then
    ' Check if file should be configured
    If (mustBeCfg = True) Then
      uniMsg = MLSupport.GGS_Params("fileErrMsg6", "%1 file not found", sourceFile)
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_POG.errMsg1", "%1. Loaded product configured for LIMS output; please configure the LIMS settings and save them to file.", uniMsg), vbExclamation
      Exit Sub
    End If
  End If
  
  ' Setup default values
  unity_main.m_fileVersion = INFOSTAR_VER
  pog_usefile = 0
  pogcomm = 0
  pogpath = LIMS_DIR
  pogfile = "lims.txt"
  pogappendornew = 0
  pogacceptreject = 0
  pogformat = "ASCII"
  pogheader = 0
  pogusestart = 0
  pogdatetime = 0
  pogdate = 0
  pogtime = 0
  pogSerialNum = 0
  pogsampid = 0
  pogcomment = 0
  pogproduct = 0
  pogpropname = 0
  pogpropvalue = 0
  pogmdist = 0
  pogsresid = 0
  pogpropoutlier = 0
  pognd = 0
  pogintercept = 0
  pogslope = 0
  pogstartchar = ""
  pogdelimtype = "Tab"
  
  For ii = 1 To MAX_MAN_INPUTS
    m_limsUserInputs(ii) = False
  Next ii
  
  pogport = 1
  pogbps = 9600
  pogdatabits = 8
  pogparity = "None"
  pogstopbits = 1
  pogflowctrl = "None"
  remoteproduct = 0
  limsinpath = ""
  limsinfile = ""
  
  If (fileExist = True) Then
    Call load_lims_file_vals(sourceFile)
  End If
  
  ' Check for invalid file version
  If (unity_main.m_fileVersion <> INFOSTAR_VER) Then
    unity_main.errorstring = (sourceFile & " had incompatible value. Version was " & unity_main.m_fileVersion & "; updated to " & INFOSTAR_VER)
    unity_main.write_error
    unity_main.m_fileVersion = INFOSTAR_VER
    m_badLIMSIniVal = True
  End If
  
  If (pog_usefile <> 0) Then
    frm_POG.chk_file.Value = 1
  Else
    frm_POG.chk_file.Value = 0
  End If
  
  If (pogcomm <> 0) Then
    frm_POG.chk_comm.Value = 1
  Else
    frm_POG.chk_comm.Value = 0
  End If
  
  ' Confirm path contains '\' instead of '/'
  filePathName = pogpath
  check_filepathname_delimiters filePathName
  pogpath = filePathName
  
  If (Right(pogpath, 1) <> "\") Then
    pogpath = pogpath & "\"
  End If
    
  frm_POG.txt_outpath.Text = pogpath
    
  frm_POG.txt_pogfilename.Text = pogfile
  
  If (pogappendornew <> 0) Then
    frm_POG.chk_appendornew.Value = 1
  Else
    frm_POG.chk_appendornew.Value = 0
  End If
  
  If (pogacceptreject <> 0) Then
    frm_POG.chk_acceptreject.Value = 1
  Else
    frm_POG.chk_acceptreject.Value = 0
  End If
  
  Select Case (pogformat)
    Case "ASCII"
      frm_POG.opt_asciiFormat.Value = True
    Case "Unicode"
      frm_POG.opt_uniFormat.Value = True
    Case Else
      unity_main.errorstring = (sourceFile & " had incompatible value. LIMS_Format was " & pogformat & "; updated to ASCII")
      unity_main.write_error
      pogformat = "ASCII"
      frm_POG.opt_asciiFormat.Value = True
      m_badLIMSIniVal = True
  End Select
  
  If (pogheader <> 0) Then
    frm_POG.chk_header.Value = 1
  Else
    frm_POG.chk_header.Value = 0
  End If
  
  If (pogusestart <> 0) Then
    frm_POG.chk_startchar.Value = 1
  Else
    frm_POG.chk_startchar.Value = 0
  End If
  
  If (pogdatetime <> 0) Then
    frm_POG.chk_datetime.Value = 1
  Else
    frm_POG.chk_datetime.Value = 0
  End If
  
  If (pogdate <> 0) Then
    frm_POG.chk_date.Value = 1
  Else
    frm_POG.chk_date.Value = 0
  End If
  
  If (pogtime <> 0) Then
    frm_POG.chk_time.Value = 1
  Else
    frm_POG.chk_time.Value = 0
  End If
  
  If (pogSerialNum <> 0) Then
    frm_POG.chk_serialNum.Value = 1
  Else
    frm_POG.chk_serialNum.Value = 0
  End If
  
  If (pogcomment <> 0) Then
    frm_POG.chk_comment.Value = 1
  Else
    frm_POG.chk_comment.Value = 0
  End If
  
  If (pogsampid <> 0) Then
    frm_POG.chk_sampid.Value = 1
  Else
    frm_POG.chk_sampid.Value = 0
  End If
  
  If (pogproduct <> 0) Then
    frm_POG.chk_product.Value = 1
  Else
    frm_POG.chk_product.Value = 0
  End If
  
  If (pogpropname <> 0) Then
    frm_POG.chk_propname.Value = 1
  Else
    frm_POG.chk_propname.Value = 0
  End If
  
  If (pogpropvalue <> 0) Then
    frm_POG.chk_propvalue.Value = 1
  Else
    frm_POG.chk_propvalue.Value = 0
  End If
  
  If (pogmdist <> 0) Then
    frm_POG.chk_mdist.Value = 1
  Else
    frm_POG.chk_mdist.Value = 0
  End If
  
  If (pogsresid <> 0) Then
    frm_POG.chk_sresid.Value = 1
  Else
    frm_POG.chk_sresid.Value = 0
  End If
  
  If (pogpropoutlier <> 0) Then
    frm_POG.chk_propOutlier.Value = 1
  Else
    frm_POG.chk_propOutlier.Value = 0
  End If
  
  If (pognd <> 0) Then
    frm_POG.chk_nd.Value = 1
  Else
    frm_POG.chk_nd.Value = 0
  End If
  
  If (pogintercept <> 0) Then
    frm_POG.chk_intercept.Value = 1
  Else
    frm_POG.chk_intercept.Value = 0
  End If
  
  If (pogslope <> 0) Then
    frm_POG.chk_slope.Value = 1
  Else
    frm_POG.chk_slope.Value = 0
  End If
  
  Select Case (pogstartchar)
    Case ">"
      frm_POG.combo_startchar.Text = pogstartchar
    Case ">>"
      frm_POG.combo_startchar.Text = pogstartchar
    Case "<"
      frm_POG.combo_startchar.Text = pogstartchar
    Case "/"
      frm_POG.combo_startchar.Text = pogstartchar
    Case "//"
      frm_POG.combo_startchar.Text = pogstartchar
    Case "\"
      frm_POG.combo_startchar.Text = pogstartchar
    Case "\\"
      frm_POG.combo_startchar.Text = pogstartchar
    Case "#"
      frm_POG.combo_startchar.Text = pogstartchar
    Case Else
      If (pogusestart = 1) Then
        unity_main.errorstring = (sourceFile & " had incompatible value. Start_Char was " & pogstartchar & "; updated to >")
        unity_main.write_error
        pogstartchar = ">"
        frm_POG.combo_startchar.Text = pogstartchar
        m_badLIMSIniVal = True
      End If
  End Select

  Select Case (pogdelimtype)
    Case "Tab"
      frm_POG.combo_delim.Text = pogdelimtype
      pogdelimchar = Chr$(vbKeyTab)
    Case "Comma"
      frm_POG.combo_delim.Text = pogdelimtype
      pogdelimchar = ","
    Case "Fixed Field"
      frm_POG.combo_delim.Text = pogdelimtype
      pogdelimchar = "ff"
    Case "Colon"
      frm_POG.combo_delim.Text = pogdelimtype
      pogdelimchar = ":"
    Case "Semicolon"
      frm_POG.combo_delim.Text = pogdelimtype
      pogdelimchar = ";"
    Case "Backslash"
      frm_POG.combo_delim.Text = pogdelimtype
      pogdelimchar = "/"
    Case "Forwardslash"
      frm_POG.combo_delim.Text = pogdelimtype
      pogdelimchar = "\"
    Case Else
      unity_main.errorstring = (sourceFile & " had incompatible value. Delimeter was " & pogdelimtype & "; updated to Tab")
      unity_main.write_error
      frm_POG.combo_delim.Text = "Tab"
      pogdelimtype = "Tab"
      pogdelimchar = Chr$(vbKeyTab)
      m_badLIMSIniVal = True
  End Select
  
  If (m_limsUserInputs(1) = True) Then
    frm_POG.chk_input1.Value = 1
  Else
    frm_POG.chk_input1.Value = 0
  End If
  
  If (m_limsUserInputs(2) = True) Then
    frm_POG.chk_input2.Value = 1
  Else
    frm_POG.chk_input2.Value = 0
  End If
  
  If (m_limsUserInputs(3) = True) Then
    frm_POG.chk_input3.Value = 1
  Else
    frm_POG.chk_input3.Value = 0
  End If
  
  If (m_limsUserInputs(4) = True) Then
    frm_POG.chk_input4.Value = 1
  Else
    frm_POG.chk_input4.Value = 0
  End If
  
  If (m_limsUserInputs(5) = True) Then
    frm_POG.chk_input5.Value = 1
  Else
    frm_POG.chk_input5.Value = 0
  End If
  
  If (m_limsUserInputs(6) = True) Then
    frm_POG.chk_input6.Value = 1
  Else
    frm_POG.chk_input6.Value = 0
  End If
  
  If (m_limsUserInputs(7) = True) Then
    frm_POG.chk_input7.Value = 1
  Else
    frm_POG.chk_input7.Value = 0
  End If
  
  If (m_limsUserInputs(8) = True) Then
    frm_POG.chk_input8.Value = 1
  Else
    frm_POG.chk_input8.Value = 0
  End If
  
  If (pogport < numInc_commPort.Min) Or (pogport > numInc_commPort.Max) Then
    unity_main.errorstring = (sourceFile & " had incompatible value. Comm_Port was " & pogport & "; updated to " & numInc_commPort.Min)
    unity_main.write_error
    pogport = numInc_commPort.Min
    m_badLIMSIniVal = True
  End If

  frm_POG.numInc_commPort.Text = pogport
  
  Select Case (pogbps)
    Case 1200
      frm_POG.combo_bps.Text = pogbps
    Case 2400
      frm_POG.combo_bps.Text = pogbps
    Case 4800
      frm_POG.combo_bps.Text = pogbps
    Case 9600
      frm_POG.combo_bps.Text = pogbps
    Case 19200
      frm_POG.combo_bps.Text = pogbps
    Case 38400
      frm_POG.combo_bps.Text = pogbps
    Case Else
      unity_main.errorstring = (sourceFile & " had incompatible value. Baud_Rate was " & pogbps & "; updated to 9600")
      unity_main.write_error
      pogbps = 9600
      frm_POG.combo_bps.Text = pogbps
      m_badLIMSIniVal = True
  End Select

  Select Case (pogdatabits)
    Case 5
      frm_POG.combo_databits.Text = pogdatabits
    Case 6
      frm_POG.combo_databits.Text = pogdatabits
    Case 7
      frm_POG.combo_databits.Text = pogdatabits
    Case 8
      frm_POG.combo_databits.Text = pogdatabits
    Case Else
      unity_main.errorstring = (sourceFile & " had incompatible value. Data_Bits was " & pogdatabits & "; updated to 8")
      unity_main.write_error
      pogdatabits = 8
      frm_POG.combo_databits.Text = pogdatabits
      m_badLIMSIniVal = True
  End Select
  
  Select Case (pogparity)
    Case "Even"
      frm_POG.combo_parity.Text = pogparity
    Case "Odd"
      frm_POG.combo_parity.Text = pogparity
    Case "None"
      frm_POG.combo_parity.Text = pogparity
    Case "Mark"
      frm_POG.combo_parity.Text = pogparity
    Case "Space"
      frm_POG.combo_parity.Text = pogparity
    Case Else
      unity_main.errorstring = (sourceFile & " had incompatible value. Parity was " & pogparity & "; updated to None")
      unity_main.write_error
      pogparity = "None"
      frm_POG.combo_parity.Text = pogparity
      m_badLIMSIniVal = True
  End Select

  Select Case (pogstopbits)
    Case "1"
      frm_POG.combo_stopbits.Text = pogstopbits
    Case "1.5"
      frm_POG.combo_stopbits.Text = pogstopbits
    Case "2"
      frm_POG.combo_stopbits.Text = pogstopbits
    Case Else
      unity_main.errorstring = (sourceFile & " had incompatible value. Stop_Bits was " & pogstopbits & "; updated to 1")
      unity_main.write_error
      pogstopbits = "1"
      frm_POG.combo_stopbits.Text = pogstopbits
      m_badLIMSIniVal = True
  End Select

  Select Case (pogflowctrl)
    Case "Xon / Xoff"
      frm_POG.combo_flowctrl.Text = pogflowctrl
    Case "Hardware"
      frm_POG.combo_flowctrl.Text = pogflowctrl
    Case "None"
      frm_POG.combo_flowctrl.Text = pogflowctrl
    Case Else
      unity_main.errorstring = (sourceFile & " had incompatible value. Flow_Control was " & pogflowctrl & "; updated to None")
      unity_main.write_error
      pogflowctrl = "None"
      frm_POG.combo_flowctrl.Text = pogflowctrl
      m_badLIMSIniVal = True
  End Select

  If (remoteproduct <> 0) Then
    frm_POG.chk_limsin.Value = 1
  Else
    frm_POG.chk_limsin.Value = 0
  End If
  
  If (limsinpath = "Not Used") Then
    limsinpath = ""
  End If
  
  If (limsinpath <> "") Then
    ' Confirm remote product file path contains '\' instead of '/'
    filePathName = limsinpath
    check_filepathname_delimiters filePathName
    limsinpath = filePathName
  
    If (Right(limsinpath, 1) <> "\") Then
      limsinpath = limsinpath & "\"
    End If
  End If
  
  If (chk_limsin.Value = 1) And (limsinpath = "") Then
    unity_main.errorstring = (sourceFile & " had incompatible value. Remote_File_Path was " & limsinpath & "; updated to " & LIMS_DIR)
    unity_main.write_error
    limsinpath = LIMS_DIR
    m_badLIMSIniVal = True
  End If
  
  frm_POG.txt_limsinpath.Text = limsinpath
  
  If (limsinfile = "Not Used") Then
    limsinfile = ""
  End If
  
  If (chk_limsin.Value = 1) And (limsinfile = "") Then
    unity_main.errorstring = (sourceFile & " had incompatible value. Remote_File_Name was " & limsinfile & "; updated to RemoteProduct.txt")
    unity_main.write_error
    limsinfile = "RemoteProduct.txt"
    m_badLIMSIniVal = True
  End If
  
  frm_POG.txt_limsin.Text = limsinfile
  
  ' Check if ini file had bad value
  If (m_badLIMSIniVal = True) Then
    unity_main.errorstring = (sourceFile & " had incompatible value(s). Updated file with proper/default value(s).")
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg7", "%1 had incompatible value(s). Updated file with proper/default value(s)", sourceFile)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg3", "%1. Refer to system log for more details", uniMsg), vbOKOnly
    Call save_lims(mustBeCfg)
  End If
  
  If (mustBeCfg = True) Then
    ' Copy loaded LIMS config into system operational variables
    unity_main.pog_usefile = pog_usefile
    unity_main.pogcomm = pogcomm
    unity_main.pogheader = pogheader
    unity_main.pogpath = pogpath
    unity_main.pogfile = pogfile
    unity_main.pogappendornew = pogappendornew
    unity_main.pogacceptreject = pogacceptreject
    unity_main.pogformat = pogformat
    unity_main.pogusestart = pogusestart
    unity_main.pogstartchar = pogstartchar
    unity_main.pogdelimtype = pogdelimtype
    unity_main.pogdelimchar = pogdelimchar
    unity_main.pogdatetime = pogdatetime
    unity_main.pogdate = pogdate
    unity_main.pogtime = pogtime
    unity_main.pogSerialNum = pogSerialNum
    unity_main.pogcomment = pogcomment
    unity_main.pogproduct = pogproduct
    unity_main.pogmodelid = pogmodelid
    unity_main.pogpropname = pogpropname
    unity_main.pogpropvalue = pogpropvalue
    unity_main.pogpropoutlier = pogpropoutlier
    unity_main.pogmdist = pogmdist
    unity_main.pogsresid = pogsresid
    unity_main.pognd = pognd
    unity_main.pogintercept = pogintercept
    unity_main.pogslope = pogslope
    unity_main.pogport = pogport
    unity_main.pogbps = pogbps
    unity_main.pogdatabits = pogdatabits
    unity_main.pogstopbits = pogstopbits
    unity_main.pogflowctrl = pogflowctrl
    unity_main.pogparity = pogparity
    unity_main.pogsampid = pogsampid
    unity_main.remoteproduct = remoteproduct
    unity_main.limsinpath = limsinpath
    unity_main.limsinfile = limsinfile
    
    For ii = 1 To MAX_MAN_INPUTS
      LIMSUserInputs(ii) = m_limsUserInputs(ii)
    Next ii
    
    If (pogcomm = 1) Then
      Call initcomm(False)
    End If
  End If
End Sub

Private Function load_lims_file_vals(srcFile As String) As Boolean
  Dim tmpFile As String
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

  On Error GoTo FILE_ERROR
  tmpFile = (CFG_DIR & TMP_LIMS_CFG_FILE)   ' Define target file name.
  uniFile.st_CopyFile srcFile, tmpFile      ' Copy source to target.
  
  If (uniFile.OpenFileRead(tmpFile) = True) Then
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
          Case "file_based"
            pog_usefile = CInt(varVal)
          Case "comm_port_based"
            pogcomm = CInt(varVal)
          Case "lims_file_path"
            pogpath = varVal
          Case "lims_filename"
            pogfile = varVal
          Case "lims_file_overwrite"
            pogappendornew = CInt(varVal)
          Case "lims_accept_reject"
            pogacceptreject = CInt(varVal)
          Case "lims_format"
            pogformat = varVal
          Case "include_header"
            pogheader = CInt(varVal)
          Case "include_start_character"
            pogusestart = CInt(varVal)
          Case "include_date_time"
            pogdatetime = CInt(varVal)
          Case "include_date"
            pogdate = CInt(varVal)
          Case "include_time"
            pogtime = CInt(varVal)
          Case "include_serialnum"
            pogSerialNum = CInt(varVal)
          Case "sample_id"            ' older .ini file version
            pogsampid = CInt(varVal)
          Case "include_sample_id"    ' newer .ini file versi0n
            pogsampid = CInt(varVal)
          Case "include_comment"
            pogcomment = CInt(varVal)
          Case "include_product"
            pogproduct = CInt(varVal)
          Case "include_property_name"
            pogpropname = CInt(varVal)
          Case "include_property_value"
            pogpropvalue = CInt(varVal)
          Case "include_property_md"
            pogmdist = CInt(varVal)
          Case "include_property_s-resid"
            pogsresid = CInt(varVal)
          Case "include_property_outlier"
            pogpropoutlier = CInt(varVal)
          Case "include_property_nd"
            pognd = CInt(varVal)
          Case "include_property_intercept"
            pogintercept = CInt(varVal)
          Case "include_property_slope"
            pogslope = CInt(varVal)
          Case "start_char"
            pogstartchar = varVal
          Case "delimeter"                ' older .ini file version
            pogdelimtype = varVal
          Case "delimiter"                ' newer .ini file version
            pogdelimtype = varVal
          Case "button_1_info"            ' older .ini file version
            m_limsUserInputs(1) = CBool(varVal)
          Case "input_1_info"             ' newer .ini file version
            m_limsUserInputs(1) = CBool(varVal)
          Case "button_2_info"            ' older .ini file version
            m_limsUserInputs(2) = CBool(varVal)
          Case "input_2_info"             ' newer .ini file version
            m_limsUserInputs(2) = CBool(varVal)
          Case "button_3_info"            ' older .ini file version
            m_limsUserInputs(3) = CBool(varVal)
          Case "input_3_info"             ' newer .ini file version
            m_limsUserInputs(3) = CBool(varVal)
          Case "button_4_info"            ' older .ini file version
            m_limsUserInputs(4) = CBool(varVal)
          Case "input_4_info"             ' newer .ini file version
            m_limsUserInputs(4) = CBool(varVal)
          Case "button_5_info"            ' older .ini file version
            m_limsUserInputs(5) = CBool(varVal)
          Case "input_5_info"             ' newer .ini file version
            m_limsUserInputs(5) = CBool(varVal)
          Case "button_6_info"            ' older .ini file version
            m_limsUserInputs(6) = CBool(varVal)
          Case "input_6_info"             ' newer .ini file version
            m_limsUserInputs(6) = CBool(varVal)
          Case "button_7_info"            ' older .ini file version
            m_limsUserInputs(7) = CBool(varVal)
          Case "input_7_info"             ' newer .ini file version
            m_limsUserInputs(7) = CBool(varVal)
          Case "button_8_info"            ' older .ini file version
            m_limsUserInputs(8) = CBool(varVal)
          Case "input_8_info"             ' newer .ini file version
            m_limsUserInputs(8) = CBool(varVal)
          Case "comm_port"
            If (varVal <> "") Then
              pogport = CInt(varVal)
            Else
              pogport = 0
            End If
          Case "bits_per_second"          ' older .ini file version
            If (varVal <> "") Then
              pogbps = CLng(varVal)
            Else
              pogbps = 0
            End If
          Case "baud_rate"                ' newer .ini file version
            If (varVal <> "") Then
              pogbps = CLng(varVal)
            Else
              pogbps = 0
            End If
          Case "data_bits"
            If (varVal <> "") Then
              pogdatabits = CInt(varVal)
            Else
              pogdatabits = 0
            End If
          Case "parity"
            pogparity = varVal
          Case "stop_bits"
            If (varVal <> "") Then
              pogstopbits = CSng(varVal)
            Else
              pogstopbits = 0
            End If
          Case "flow_control"
            pogflowctrl = varVal
          Case "enable lims control"            ' older .ini file version
            remoteproduct = CInt(varVal)
          Case "enable_remote_control"          ' newer .ini file version
            remoteproduct = CInt(varVal)
          Case "lims in file path"              ' older .ini file version
            limsinpath = varVal
          Case "remote_file_path"               ' newer .ini file version
            limsinpath = varVal
          Case "lims control file name"         ' older .ini file version
            limsinfile = varVal
          Case "remote_file_name"               ' newer .ini file version
            limsinfile = varVal
        End Select
      End If
    Wend
    
    load_lims_file_vals = True
  Else
FILE_ERROR:
    errMsg = (srcFile & " file read error. " & Error$)
    unity_main.errorstring = errMsg
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg4", "%1 file read error. %2", srcFile, Error$)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    load_lims_file_vals = False
  End If
 
  uniFile.CloseFile
  
  If (tmpFile <> "") And (uniFile.st_FileExist(tmpFile) = True) Then
    uniFile.st_RmFile tmpFile
  End If
  
  Exit Function
  
BAD_INI_VALUE:
  unity_main.errorstring = (CFG_DIR & LIMS_CFG_FILE & " line " & CStr(lineCnt) & " had incompatible value. " & Trim(tmpStrg) & " - " & varVal & "; will use default value")
  unity_main.write_error
  m_badLIMSIniVal = True
  Resume Next
End Function

Private Function check_cfg_settings() As Boolean
  
  check_cfg_settings = False
  
  ' Check parameters if to output LIMS to file
  If (chk_file.Value = 1) Then
    If (Trim(frm_POG.txt_outpath.Text) = "") Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_POG", "errMsg1", "Please enter a file path for LIMS output file"), vbExclamation
      Exit Function
    End If
  
    If (Trim(frm_POG.txt_pogfilename.Text) = "") Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_POG", "errMsg2", "Please enter a file name for LIMS output file"), vbExclamation
      Exit Function
    End If
  End If

  If (frm_POG.chk_startchar.Value = 1) And (frm_POG.combo_startchar.Text = "") Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_POG", "errMsg3", "Please select a Start Character for LIMS output"), vbExclamation
    Exit Function
  End If

  If (frm_POG.combo_delim.Text = "") Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_POG", "errMsg4", "Please select a Delimiter for LIMS output"), vbExclamation
    Exit Function
  End If
  
  If (frm_POG.chk_comm.Value = 1) Then
    ' NOTE: Replaced NI CWSerial w/ MS Comm Control MSComm 05/01/08 M.Spivey
    Dim portOkay As Boolean

    If (Trim(frm_POG.combo_bps.Text) = "") Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_POG", "errMsg5", "Please select a Baud Rate for LIMS serial output"), vbExclamation
      Exit Function
    End If
    
    If (Trim(frm_POG.combo_databits.Text) = "") Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_POG", "errMsg6", "Please select a Data Bits for LIMS serial output"), vbExclamation
      Exit Function
    End If
    
    If (Trim(frm_POG.combo_parity.Text) = "") Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_POG", "errMsg7", "Please select a Parity for LIMS serial output"), vbExclamation
      Exit Function
    End If
    
    If (Trim(frm_POG.combo_stopbits.Text) = "") Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_POG", "errMsg8", "Please select a Stop Bits for LIMS serial output"), vbExclamation
      Exit Function
    End If
    
    If (Trim(frm_POG.combo_flowctrl.Text) = "") Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_POG", "errMsg9", "Please select a Flow Control for LIMS serial output"), vbExclamation
      Exit Function
    End If
    
    portOkay = initcomm(True)
    
    If (portOkay = False) Then
      Exit Function
    End If
  End If
  
  ' Check parameters if product selection via remote control
  If (chk_limsin.Value = 1) Then
    If (Trim(frm_POG.txt_limsin.Text) = "") Or (Trim(frm_POG.txt_limsin.Text) = "Not Used") Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_POG", "errMsg10", "Please enter a file name for LIMS remote control product selection"), vbExclamation
      Exit Function
    End If
  
    If (Trim(frm_POG.txt_limsinpath.Text) = "") Or (Trim(frm_POG.txt_limsinpath.Text) = "Not Used") Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_POG", "errMsg11", "Please enter a file path for LIMS remote control product selection"), vbExclamation
      Exit Function
    End If
  End If
  
  check_cfg_settings = True
End Function

Private Function initcomm(newCfgFlg As Boolean) As Boolean
  Dim port As Integer
  Dim baud As Integer
  Dim parity As String
  Dim dataBits As Integer
  Dim stopBits As Integer
  Dim flowCtrl As String

  On Error GoTo ClosePortError
  initcomm = False

  If (MSComm2.PortOpen = True) Then
    MSComm2.PortOpen = False     ' close port
  End If
    
  On Error GoTo OpenPortError
  
  If (newCfgFlg = True) Then
    port = numInc_commPort.Text
    baud = combo_bps.Text
    dataBits = combo_databits.Text
    parity = combo_parity.Text
    stopBits = combo_stopbits.Text
    flowCtrl = combo_flowctrl.Text
  Else
    port = unity_main.pogport
    baud = unity_main.pogbps
    dataBits = unity_main.pogdatabits
    parity = unity_main.pogparity
    stopBits = unity_main.pogstopbits
    flowCtrl = unity_main.pogflowctrl
  End If
  
  MSComm2.CommPort = port   ' setup comm port #

  ' Setup port parity
  Select Case parity
    Case "Even"
      parity = "E"
    Case "Odd"
      parity = "O"
    Case "None"
      parity = "N"
    Case "Mark"
      parity = "M"
    Case "Space"
      parity = "S"
  End Select

  ' Comm parameter settings = "baud,parity,databit,stopbit"
  MSComm2.Settings = baud & "," & parity & "," & dataBits & "," & stopBits
  
  ' Setup port flow control
  Select Case flowCtrl
    Case "Xon / Xoff"
      flowCtrl = comXOnXoff
    Case "None"
      flowCtrl = comNone
    Case "Hardware"
      flowCtrl = comRTS
  End Select

  MSComm2.Handshaking = flowCtrl
  
  ' Open comm port
  MSComm2.PortOpen = True
  initcomm = True
  Exit Function
 
OpenPortError:
  CWrap.ShowMessageBoxW MLSupport.GSS("frm_POG", "errMsg12", "Problem with configuring selected LIMS com port, please confirm you've selected a valid port and it is functional"), vbCritical
  Exit Function
  
ClosePortError:
  CWrap.ShowMessageBoxW MLSupport.GSS("frm_POG", "errMsg13", "Problem closing previous configured LIMS com port"), vbCritical
End Function

Private Sub save_lims(setupCfg As Boolean)
  Dim tempstring As String
  Dim fileExt As String
  Dim filePathName As String
  Dim ii As Integer
  Dim uniFile As New clsUniFile
  Dim uniMsg As String
  Dim errMsg As String
  
  ' Update settings
  pog_usefile = frm_POG.chk_file.Value
  pogcomm = frm_POG.chk_comm.Value
  
  ' Confirm path contains '\' instead of '/'
  filePathName = Trim(frm_POG.txt_outpath.Text)
  check_filepathname_delimiters filePathName
  frm_POG.txt_outpath.Text = filePathName
  
  If (Right(frm_POG.txt_outpath.Text, 1) <> "\") Then
    frm_POG.txt_outpath.Text = frm_POG.txt_outpath.Text & "\"
  End If
  
  ' Save LIMS file path
  pogpath = frm_POG.txt_outpath.Text

  pogfile = Trim(frm_POG.txt_pogfilename.Text)
  
  pogappendornew = frm_POG.chk_appendornew.Value
  pogacceptreject = frm_POG.chk_acceptreject.Value
  
  If (frm_POG.opt_asciiFormat.Value = True) Then
    pogformat = "ASCII"
  Else
    pogformat = "Unicode"
  End If
  
  pogheader = frm_POG.chk_header.Value
  pogusestart = frm_POG.chk_startchar.Value
  pogdatetime = frm_POG.chk_datetime.Value
  pogdate = frm_POG.chk_date.Value
  pogtime = frm_POG.chk_time.Value
  pogSerialNum = frm_POG.chk_serialNum.Value
  pogsampid = frm_POG.chk_sampid.Value
  pogcomment = frm_POG.chk_comment.Value
  pogproduct = frm_POG.chk_product.Value
  pogpropname = frm_POG.chk_propname.Value
  pogpropvalue = frm_POG.chk_propvalue.Value
  pogmdist = frm_POG.chk_mdist.Value
  pogsresid = frm_POG.chk_sresid.Value
  pogpropoutlier = frm_POG.chk_propOutlier.Value
  pognd = frm_POG.chk_nd.Value
  pogintercept = frm_POG.chk_intercept.Value
  pogslope = frm_POG.chk_slope.Value
  
  pogstartchar = frm_POG.combo_startchar.Text
  
  Select Case (frm_POG.combo_delim.Text)
    Case "Tab"
      pogdelimchar = Chr$(vbKeyTab)
    Case "Comma"
      pogdelimchar = ","
    Case "Fixed Field"
      pogdelimchar = "ff"
    Case "Colon"
      pogdelimchar = ":"
    Case "Semicolon"
      pogdelimchar = ";"
    Case "Backslash"
      pogdelimchar = "/"
    Case "Forwardslash"
      pogdelimchar = "\"
  End Select
  
  pogdelimtype = frm_POG.combo_delim.Text
  
  m_limsUserInputs(1) = frm_POG.chk_input1.Value
  m_limsUserInputs(2) = frm_POG.chk_input2.Value
  m_limsUserInputs(3) = frm_POG.chk_input3.Value
  m_limsUserInputs(4) = frm_POG.chk_input4.Value
  m_limsUserInputs(5) = frm_POG.chk_input5.Value
  m_limsUserInputs(6) = frm_POG.chk_input6.Value
  m_limsUserInputs(7) = frm_POG.chk_input7.Value
  m_limsUserInputs(8) = frm_POG.chk_input8.Value
  
  If (frm_POG.numInc_commPort.Text = "") Then
    pogport = 0
  Else
    pogport = CInt(frm_POG.numInc_commPort.Text)
  End If
  
  If (frm_POG.combo_bps.Text = "") Then
    pogbps = 0
  Else
    pogbps = CLng(frm_POG.combo_bps.Text)
  End If
    
  If (frm_POG.combo_databits.Text = "") Then
    pogdatabits = 0
  Else
    pogdatabits = CInt(frm_POG.combo_databits.Text)
  End If
  
  pogparity = frm_POG.combo_parity.Text
  
  If (frm_POG.combo_stopbits.Text = "") Then
    pogstopbits = 0
  Else
    pogstopbits = CSng(frm_POG.combo_stopbits.Text)
  End If
  
  pogflowctrl = frm_POG.combo_flowctrl.Text
  
  remoteproduct = frm_POG.chk_limsin.Value
    
  ' Get remote product file name
  limsinfile = Trim(txt_limsin.Text)
  
  If (Trim(frm_POG.txt_limsinpath.Text) <> "") Then
    ' Confirm remote product file path contains '\' instead of '/'
    filePathName = Trim(frm_POG.txt_limsinpath.Text)
    check_filepathname_delimiters filePathName
    frm_POG.txt_limsinpath.Text = filePathName
    
    If Right(frm_POG.txt_limsinpath.Text, 1) <> "\" Then
      frm_POG.txt_limsinpath.Text = frm_POG.txt_limsinpath.Text & "\"
    End If
  End If
  
  ' Save remote product file path
  limsinpath = frm_POG.txt_limsinpath.Text
  
  On Error GoTo FILE_ERROR
  
  If (uniFile.OpenFileWrite(CFG_DIR & LIMS_CFG_FILE) = True) Then
    uniFile.WriteBOM fe_UTF16LE
    uniFile.WriteUnicodeLine ("Version=" & INFOSTAR_VER)
    uniFile.WriteUnicodeLine ("File_Based=" & pog_usefile)
    uniFile.WriteUnicodeLine ("Comm_Port_Based=" & pogcomm)
    uniFile.WriteUnicodeLine ("LIMS_File_Path=" & pogpath)
    uniFile.WriteUnicodeLine ("LIMS_Filename=" & pogfile)
    uniFile.WriteUnicodeLine ("LIMS_File_Overwrite=" & pogappendornew)
    uniFile.WriteUnicodeLine ("LIMS_Accept_Reject=" & pogacceptreject)
    uniFile.WriteUnicodeLine ("LIMS_Format=" & pogformat)
    uniFile.WriteUnicodeLine ("Include_Header=" & pogheader)
    uniFile.WriteUnicodeLine ("Include_Start_Character=" & pogusestart)
    uniFile.WriteUnicodeLine ("Include_Date_Time=" & pogdatetime)
    uniFile.WriteUnicodeLine ("Include_Date=" & pogdate)
    uniFile.WriteUnicodeLine ("Include_Time=" & pogtime)
    uniFile.WriteUnicodeLine ("Include_SerialNum=" & pogSerialNum)
    uniFile.WriteUnicodeLine ("Include_Sample_ID=" & pogsampid)
    uniFile.WriteUnicodeLine ("Include_Comment=" & pogcomment)
    uniFile.WriteUnicodeLine ("Include_Product=" & pogproduct)
    uniFile.WriteUnicodeLine ("Include_Property_Name=" & pogpropname)
    uniFile.WriteUnicodeLine ("Include_Property_Value=" & pogpropvalue)
    uniFile.WriteUnicodeLine ("Include_Property_MD=" & pogmdist)
    uniFile.WriteUnicodeLine ("Include_Property_S-Resid=" & pogsresid)
    uniFile.WriteUnicodeLine ("Include_Property_Outlier=" & pogpropoutlier)
    uniFile.WriteUnicodeLine ("Include_Property_ND=" & pognd)
    uniFile.WriteUnicodeLine ("Include_Property_Intercept=" & pogintercept)
    uniFile.WriteUnicodeLine ("Include_Property_Slope=" & pogslope)
    uniFile.WriteUnicodeLine ("Start_Char=" & pogstartchar)
    uniFile.WriteUnicodeLine ("Delimiter=" & frm_POG.combo_delim.Text)
    uniFile.WriteUnicodeLine ("Input_1_Info=" & m_limsUserInputs(1))
    uniFile.WriteUnicodeLine ("Input_2_Info=" & m_limsUserInputs(2))
    uniFile.WriteUnicodeLine ("Input_3_Info=" & m_limsUserInputs(3))
    uniFile.WriteUnicodeLine ("Input_4_Info=" & m_limsUserInputs(4))
    uniFile.WriteUnicodeLine ("Input_5_Info=" & m_limsUserInputs(5))
    uniFile.WriteUnicodeLine ("Input_6_Info=" & m_limsUserInputs(6))
    uniFile.WriteUnicodeLine ("Input_7_Info=" & m_limsUserInputs(7))
    uniFile.WriteUnicodeLine ("Input_8_Info=" & m_limsUserInputs(8))
    uniFile.WriteUnicodeLine ("Comm_Port=" & frm_POG.numInc_commPort.Text)
    uniFile.WriteUnicodeLine ("Baud_Rate=" & frm_POG.combo_bps.Text)
    uniFile.WriteUnicodeLine ("Data_Bits=" & frm_POG.combo_databits.Text)
    uniFile.WriteUnicodeLine ("Parity=" & frm_POG.combo_parity.Text)
    uniFile.WriteUnicodeLine ("Stop_Bits=" & frm_POG.combo_stopbits.Text)
    uniFile.WriteUnicodeLine ("Flow_Control=" & frm_POG.combo_flowctrl.Text)
    uniFile.WriteUnicodeLine ("Enable_Remote_Control=" & remoteproduct)
    uniFile.WriteUnicodeLine ("Remote_File_Path=" & limsinpath)
    uniFile.WriteUnicodeLine ("Remote_File_Name=" & limsinfile)
    uniFile.Flush
    
    If (setupCfg = True) Then
      ' Copy saved LIMS config into system operational variables
      unity_main.pog_usefile = pog_usefile
      unity_main.pogcomm = pogcomm
      unity_main.pogheader = pogheader
      unity_main.pogpath = pogpath
      unity_main.pogfile = pogfile
      unity_main.pogappendornew = pogappendornew
      unity_main.pogacceptreject = pogacceptreject
      unity_main.pogformat = pogformat
      unity_main.pogusestart = pogusestart
      unity_main.pogstartchar = pogstartchar
      unity_main.pogdelimtype = pogdelimtype
      unity_main.pogdelimchar = pogdelimchar
      unity_main.pogdatetime = pogdatetime
      unity_main.pogdate = pogdate
      unity_main.pogtime = pogtime
      unity_main.pogSerialNum = pogSerialNum
      unity_main.pogcomment = pogcomment
      unity_main.pogproduct = pogproduct
      unity_main.pogmodelid = pogmodelid
      unity_main.pogpropname = pogpropname
      unity_main.pogpropvalue = pogpropvalue
      unity_main.pogpropoutlier = pogpropoutlier
      unity_main.pogmdist = pogmdist
      unity_main.pogsresid = pogsresid
      unity_main.pognd = pognd
      unity_main.pogintercept = pogintercept
      unity_main.pogslope = pogslope
      unity_main.pogport = pogport
      unity_main.pogbps = pogbps
      unity_main.pogdatabits = pogdatabits
      unity_main.pogstopbits = pogstopbits
      unity_main.pogflowctrl = pogflowctrl
      unity_main.pogparity = pogparity
      unity_main.pogsampid = pogsampid
      unity_main.remoteproduct = remoteproduct
      unity_main.limsinpath = limsinpath
      unity_main.limsinfile = limsinfile
      
      For ii = 1 To MAX_MAN_INPUTS
        LIMSUserInputs(ii) = m_limsUserInputs(ii)
      Next ii
      
      If (pogcomm = 1) Then
        Call initcomm(False)
      End If
    End If
  Else
FILE_ERROR:
    errMsg = (CFG_DIR & LIMS_CFG_FILE & " file write error. " & Error$)
    unity_main.errorstring = errMsg
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", CFG_DIR & LIMS_CFG_FILE, Error$)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  End If
  
  uniFile.CloseFile
End Sub

Private Sub padlen() 'used to pad for fixed field outputs
  Dim strlen, zz As Integer

  strlen = Len(passstringd)
  
  For zz = strlen To 29
    passstringd = passstringd & Chr$(vbKeySpace)
  Next zz
End Sub

Private Sub cmd_cancel_Click()
  
  unity_main.errorstring = "LIMS Configuration screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Unload frm_POG
End Sub

Private Sub cmd_save_Click()
  
  unity_main.errorstring = "LIMS Configuration screen Save Changes button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  If (check_cfg_settings = True) Then
    Call save_lims(True)
    Unload frm_POG

    unity_main.errorstring = ("User saved new settings for configuration file: " & (CFG_DIR & LIMS_CFG_FILE))
    unity_main.write_error (LOG_DBG_LEVEL1)
  End If
End Sub

Private Sub Form_Load()
  
  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
  ' Setup lists selections
  setup_lims_lists
End Sub

Private Sub Picture1_Click()
  
  unity_main.formfrom = 12
  unity_main.varfrom = 4
  frm_kybd.lbl_kybd.Caption = frm_POG.Label13.Caption
  frm_kybd.txt_kybd.Text = Trim(frm_POG.txt_limsin.Text)
  frm_kybd.Show 1
End Sub

Private Sub Picture2_Click()
  
  unity_main.formfrom = 12
  unity_main.varfrom = 3
  frm_kybd.lbl_kybd.Caption = frm_POG.Label12.Caption
  frm_kybd.txt_kybd.Text = Trim(frm_POG.txt_pogfilename.Text)
  frm_kybd.Show 1
End Sub

Private Sub Picture3_Click()
  
  unity_main.formfrom = 12
  unity_main.varfrom = 1
  frm_kybd.lbl_kybd.Caption = frm_POG.Label14.Caption
  frm_kybd.txt_kybd.Text = Trim(frm_POG.txt_outpath.Text)
  frm_kybd.Show 1
End Sub

Private Sub Picture4_Click()
  
  unity_main.formfrom = 12
  unity_main.varfrom = 5
  frm_kybd.lbl_kybd.Caption = frm_POG.Label14.Caption
  frm_kybd.txt_kybd.Text = Trim(frm_POG.txt_limsinpath.Text)
  frm_kybd.Show 1
End Sub








