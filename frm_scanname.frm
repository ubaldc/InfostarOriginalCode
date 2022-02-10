VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{488D84E9-A9A9-4C74-949B-F7752803ADEE}#99.2#0"; "EMDVirtualKeyBoard.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_scanname 
   Caption         =   "User Inputs"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12150
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
   Icon            =   "frm_scanname.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   12150
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniComboBoxXP combo 
      Height          =   360
      Index           =   8
      Left            =   8880
      TabIndex        =   17
      Top             =   2595
      Width           =   2505
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
      Tip             =   "frm_scanname.frx":0442
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
      MouseIcon       =   "frm_scanname.frx":0462
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniComboBoxXP combo 
      Height          =   360
      Index           =   7
      Left            =   6180
      TabIndex        =   15
      Top             =   2595
      Width           =   2500
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
      Tip             =   "frm_scanname.frx":047E
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
      MouseIcon       =   "frm_scanname.frx":049E
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniComboBoxXP combo 
      Height          =   360
      Index           =   6
      Left            =   3480
      TabIndex        =   13
      Top             =   2595
      Width           =   2500
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
      Tip             =   "frm_scanname.frx":04BA
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
      MouseIcon       =   "frm_scanname.frx":04DA
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniComboBoxXP combo 
      Height          =   360
      Index           =   5
      Left            =   720
      TabIndex        =   11
      Top             =   2595
      Width           =   2500
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
      Tip             =   "frm_scanname.frx":04F6
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
      MouseIcon       =   "frm_scanname.frx":0516
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniComboBoxXP combo 
      Height          =   360
      Index           =   4
      Left            =   8880
      TabIndex        =   9
      Top             =   1530
      Width           =   2500
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
      Tip             =   "frm_scanname.frx":0532
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
      MouseIcon       =   "frm_scanname.frx":0552
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniComboBoxXP combo 
      Height          =   360
      Index           =   3
      Left            =   6180
      TabIndex        =   7
      Top             =   1530
      Width           =   2500
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
      Tip             =   "frm_scanname.frx":056E
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
      MouseIcon       =   "frm_scanname.frx":058E
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniComboBoxXP combo 
      Height          =   360
      Index           =   2
      Left            =   3480
      TabIndex        =   5
      Top             =   1530
      Width           =   2500
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
      Tip             =   "frm_scanname.frx":05AA
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
      MouseIcon       =   "frm_scanname.frx":05CA
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniComboBoxXP combo 
      Height          =   360
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   1530
      Width           =   2500
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
      Tip             =   "frm_scanname.frx":05E6
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
      MouseIcon       =   "frm_scanname.frx":0606
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniLabelNoFlick lbl_prod 
      Height          =   495
      Left            =   8280
      Top             =   480
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scanname.frx":0622
      BackColor       =   -2147483633
      ForeColor       =   16711680
      BorderColor     =   0
      Alignment       =   2
      VAlignment      =   1
      BackStyle       =   0
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scanname.frx":0642
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0662
   End
   Begin HexUniControls.ctlUniTextBoxXP txtbx 
      Height          =   360
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   1530
      Visible         =   0   'False
      Width           =   2500
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_scanname.frx":067E
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
      Tip             =   "frm_scanname.frx":069E
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":06BE
   End
   Begin HexUniControls.ctlUniTextBoxXP txtbx 
      Height          =   360
      Index           =   2
      Left            =   3480
      TabIndex        =   4
      Top             =   1530
      Visible         =   0   'False
      Width           =   2500
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_scanname.frx":06DA
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
      Tip             =   "frm_scanname.frx":06FA
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":071A
   End
   Begin HexUniControls.ctlUniTextBoxXP txtbx 
      Height          =   360
      Index           =   3
      Left            =   6180
      TabIndex        =   6
      Top             =   1530
      Visible         =   0   'False
      Width           =   2500
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_scanname.frx":0736
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
      Tip             =   "frm_scanname.frx":0756
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0776
   End
   Begin HexUniControls.ctlUniTextBoxXP txtbx 
      Height          =   360
      Index           =   4
      Left            =   8880
      TabIndex        =   8
      Top             =   1530
      Visible         =   0   'False
      Width           =   2500
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_scanname.frx":0792
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
      Tip             =   "frm_scanname.frx":07B2
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":07D2
   End
   Begin HexUniControls.ctlUniTextBoxXP txtbx 
      Height          =   360
      Index           =   5
      Left            =   720
      TabIndex        =   10
      Top             =   2595
      Visible         =   0   'False
      Width           =   2500
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_scanname.frx":07EE
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
      Tip             =   "frm_scanname.frx":080E
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":082E
   End
   Begin HexUniControls.ctlUniTextBoxXP txtbx 
      Height          =   360
      Index           =   6
      Left            =   3480
      TabIndex        =   12
      Top             =   2595
      Visible         =   0   'False
      Width           =   2500
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_scanname.frx":084A
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
      Tip             =   "frm_scanname.frx":086A
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":088A
   End
   Begin HexUniControls.ctlUniTextBoxXP txtbx 
      Height          =   360
      Index           =   7
      Left            =   6180
      TabIndex        =   14
      Top             =   2595
      Visible         =   0   'False
      Width           =   2500
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_scanname.frx":08A6
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
      Tip             =   "frm_scanname.frx":08C6
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":08E6
   End
   Begin HexUniControls.ctlUniTextBoxXP txtbx 
      Height          =   360
      Index           =   8
      Left            =   8880
      TabIndex        =   16
      Top             =   2595
      Visible         =   0   'False
      Width           =   2500
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_scanname.frx":0902
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
      Tip             =   "frm_scanname.frx":0922
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0942
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   10680
      Top             =   3840
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8190
      FormDesignWidth =   12150
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   700
      Left            =   9600
      TabIndex        =   19
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
      Caption         =   "frm_scanname.frx":095E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scanname.frx":098A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":09AA
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_run 
      Height          =   700
      Left            =   6720
      TabIndex        =   18
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
      Caption         =   "frm_scanname.frx":09C6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scanname.frx":0A04
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0A24
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_sampinfo 
      Height          =   450
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Width           =   4695
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_scanname.frx":0A40
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
      Tip             =   "frm_scanname.frx":0A60
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0A80
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_fname 
      Height          =   450
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3135
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_scanname.frx":0A9C
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
      Tip             =   "frm_scanname.frx":0ABC
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0ADC
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_exit 
      Height          =   700
      Left            =   9600
      TabIndex        =   20
      Top             =   7080
      Visible         =   0   'False
      Width           =   1750
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
      Caption         =   "frm_scanname.frx":0AF8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scanname.frx":0B20
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0B40
   End
   Begin VBoard_EMD.KeySet_STD KeySet_STD1 
      Height          =   3180
      Left            =   120
      Top             =   3240
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   5609
      LargeFontName   =   "Arial Unicode MS"
      LargeFontSize   =   6.75
      SmallFontName   =   "Arial Unicode MS"
      ButtonWidth     =   40
      ButtonHeight    =   40
      SendKeysEnabled =   -1  'True
      RegAtDT         =   -1  'True
   End
   Begin HexUniControls.ctlUniLabel lbl_prodName 
      Height          =   375
      Left            =   8400
      Top             =   120
      Width           =   3615
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
      Caption         =   "frm_scanname.frx":0B5C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scanname.frx":0B9A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0BBA
   End
   Begin HexUniControls.ctlUniLabel lbl 
      Height          =   345
      Index           =   1
      Left            =   720
      Top             =   1200
      Width           =   2505
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
      Caption         =   "frm_scanname.frx":0BD6
      BackColor       =   -2147483633
      ForeColor       =   32768
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scanname.frx":0BF6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0C16
   End
   Begin HexUniControls.ctlUniLabel lbl 
      Height          =   345
      Index           =   2
      Left            =   3480
      Top             =   1200
      Width           =   2505
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
      Caption         =   "frm_scanname.frx":0C32
      BackColor       =   -2147483633
      ForeColor       =   32768
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scanname.frx":0C52
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0C72
   End
   Begin HexUniControls.ctlUniLabel lbl 
      Height          =   345
      Index           =   3
      Left            =   6180
      Top             =   1200
      Width           =   2505
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
      Caption         =   "frm_scanname.frx":0C8E
      BackColor       =   -2147483633
      ForeColor       =   32768
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scanname.frx":0CAE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0CCE
   End
   Begin HexUniControls.ctlUniLabel lbl 
      Height          =   345
      Index           =   4
      Left            =   8880
      Top             =   1200
      Width           =   2505
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
      Caption         =   "frm_scanname.frx":0CEA
      BackColor       =   -2147483633
      ForeColor       =   32768
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scanname.frx":0D0A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0D2A
   End
   Begin HexUniControls.ctlUniLabel lbl 
      Height          =   345
      Index           =   5
      Left            =   720
      Top             =   2190
      Width           =   2505
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
      Caption         =   "frm_scanname.frx":0D46
      BackColor       =   -2147483633
      ForeColor       =   32768
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scanname.frx":0D66
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0D86
   End
   Begin HexUniControls.ctlUniLabel lbl 
      Height          =   345
      Index           =   6
      Left            =   3480
      Top             =   2190
      Width           =   2505
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
      Caption         =   "frm_scanname.frx":0DA2
      BackColor       =   -2147483633
      ForeColor       =   32768
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scanname.frx":0DC2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0DE2
   End
   Begin HexUniControls.ctlUniLabel lbl 
      Height          =   345
      Index           =   7
      Left            =   6180
      Top             =   2190
      Width           =   2505
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
      Caption         =   "frm_scanname.frx":0DFE
      BackColor       =   -2147483633
      ForeColor       =   32768
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scanname.frx":0E1E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0E3E
   End
   Begin HexUniControls.ctlUniLabel lbl 
      Height          =   345
      Index           =   8
      Left            =   8880
      Top             =   2190
      Width           =   2505
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
      Caption         =   "frm_scanname.frx":0E5A
      BackColor       =   -2147483633
      ForeColor       =   32768
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scanname.frx":0E7A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0E9A
   End
   Begin HexUniControls.ctlUniLabel lbl_sampinfo 
      Height          =   375
      Left            =   3480
      Top             =   120
      Width           =   4695
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
      Caption         =   "frm_scanname.frx":0EB6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scanname.frx":0EF2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0F12
   End
   Begin HexUniControls.ctlUniLabel lbl_fname 
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   3135
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
      Caption         =   "frm_scanname.frx":0F2E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scanname.frx":0F64
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scanname.frx":0F84
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   10800
      Top             =   4800
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
      Top             =   4320
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_scanname.frx":0FA0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   4
      X1              =   720
      X2              =   11400
      Y1              =   2040
      Y2              =   2040
   End
End
Attribute VB_Name = "frm_scanname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public entername As Boolean
Public from_button As Integer
Public miv_cat As String  'string with all inputs
Public gotrepname As Boolean

Sub gotthename()
  Dim ii As Integer
  
  stripcommas
  unity_main.txtsamplename.Text = Trim(frm_scanname.txt_fname.Text)
  unity_main.txtsampcomment.Text = Trim(frm_scanname.txt_sampinfo.Text)
  
  ' Save user inputs values
  For ii = 1 To MAX_MAN_INPUTS
    If (frm_scanname.combo(ii).Visible = True) Then
      UserInputs(ii) = frm_scanname.combo(ii).Text
    Else
      If (frm_scanname.txtbx(ii).Visible = True) Then
        UserInputs(ii) = frm_scanname.txtbx(ii).Text
      End If
    End If
  Next ii
  
  frm_scanname.Visible = False
  unity_main.gotscanname = True
End Sub

Sub runwithbuttons()
  Dim fileName As String
  Dim optVal As Integer
  Dim tempint As Integer
  Dim ii, nn As Integer
  Dim fileExt As String
  Dim nchars As Integer
  Dim uniMsg As String

  tempint = Len(Trim(frm_scanname.txt_fname.Text))
  
  If (tempint = 0) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_scanname", "errMsg1", "Please enter sample name"), vbExclamation
    frm_scanname.txt_fname.SetFocus
    Exit Sub
  End If
  
  For ii = 1 To MAX_MAN_INPUTS
    If (frm_scanname.combo(ii).Visible = True) And (frm_scanname.combo(ii).Text <> "") Then
      For nn = 1 To frm_scanname.combo(ii).ListCount
        If (frm_scanname.combo(ii).Text = frm_scanname.combo(ii).List(nn - 1)) Then GoTo NEXT_COMBO
      Next nn
     
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_scanname", "errMsg2", "Please select entry from list or delete entry"), vbExclamation
      frm_scanname.combo(ii).SetFocus
      Exit Sub
    End If
NEXT_COMBO:
  Next ii

  valfilename
  
  If (LCase(unity_main.m_saveIt) = "save") Then
    fileName = unity_main.m_saveDir & frm_scanname.txt_fname.Text & SPC_FILE_EXT

    If (CFile.st_FileExist(fileName) = True) Then
      uniMsg = MLSupport.GSS("frm_scanname", "statMsg1", "A sample spectrum with this name already exists, do you want to overwrite it?")
      optVal = CWrap.ShowMessageBoxW(uniMsg, vbYesNo)
    
      If (optVal = vbNo) Then
        Exit Sub
      End If
    End If
  End If
  
  frm_scanname.gotthename
End Sub

Sub stripcommas()
  Dim str1 As String
  Dim str2 As String
  Dim zz As Integer
  Dim onechar As String
  Dim lenstr As Integer
  
  str1 = Trim(frm_scanname.txt_sampinfo.Text)
  str2 = ""

  For zz = 1 To Len(str1)
    onechar = Mid(str1, zz, 1)
    
    If (onechar = ",") Then
      onechar = ";"
    End If
    
    str2 = str2 & onechar
  Next zz
  
  frm_scanname.txt_sampinfo.Text = str2
End Sub

Sub valfilename()
  Dim i As Integer
  Dim S As String
  Dim fileName As String
      
  fileName = txt_fname.Text
  i = 1
    
  Do While i <= Len(fileName)
    S = Mid(fileName, i, 1)
        
    ' Remove any illegal chars
    If S = "\" Or S = "/" Or S = ":" Or S = "*" Or S = "?" Or S = "<" Or S = ">" Or S = "|" Or S = Chr(34) Then
      If i > 1 Then
        fileName = Left(fileName, i - 1) & Right(fileName, Len(fileName) - i)
      Else
        fileName = Right(fileName, Len(fileName) - 1)
      End If
    End If
    
    i = i + 1
  Loop
    
  txt_fname.Text = fileName
End Sub

Private Sub cmd_cancel_Click()
  
  Call unity_main.kill_loop(LOG_DBG_LEVEL1, True, "User Inputs screen Cancel button selected")
  unity_main.txtsamplename.Text = ""
  unity_main.txtsampcomment.Text = ""
  frm_scanname.Visible = False
End Sub

Private Sub cmd_exit_Click()
  
  unity_main.errorstring = "User Inputs screen Exit button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  frm_scanname.Visible = False
End Sub

Private Sub cmd_run_Click()

  unity_main.errorstring = "User Inputs screen Run Sample button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Call frm_scanname.runwithbuttons
End Sub

Private Sub Form_Activate()
  Dim ii As Integer

  ' Check if to clear user input entries
  If (unity_main.m_clrUserInputs = True) Then
    For ii = 1 To MAX_MAN_INPUTS
      frm_scanname.combo(ii).Text = ""
      frm_scanname.txtbx(ii).Text = ""
    Next ii
  End If

  ' Check if using global sample naming convention
  If (unity_main.m_enableGlobalName = True) Then
    frm_scanname.txt_fname.Text = unity_main.txtsamplename.Text
    frm_scanname.txt_sampinfo.SetFocus
  Else
    If (unity_main.m_sNameMode <> 1) Then
      frm_scanname.txt_sampinfo.SetFocus
    Else
      ' Check if to clear manual entries
      If (unity_main.m_clrManualName = True) Then
        frm_scanname.txt_fname.Text = ""
        frm_scanname.txt_sampinfo.Text = ""
      End If
  
      frm_scanname.txt_fname.SetFocus
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  ' Check if F1 key depressed
  If (KeyCode = vbKeyF1) Then
    cmd_run_Click
  End If
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub








