VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_lmkybd 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Infostar Keyboard"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12450
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
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
   ScaleHeight     =   7095
   ScaleWidth      =   12450
   StartUpPosition =   3  'Windows Default
   Begin ResizeLibCtl.ReSize ReSize2 
      Left            =   240
      Top             =   5040
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7095
      FormDesignWidth =   12450
   End
   Begin HexUniControls.ctlUniButtonImageXP Command67 
      Height          =   615
      Left            =   10080
      TabIndex        =   67
      Top             =   1080
      Width           =   1335
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
      Caption         =   "frm_lmkybd.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0028
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0048
   End
   Begin HexUniControls.ctlUniButtonImageXP Command56 
      Height          =   615
      Left            =   8760
      TabIndex        =   66
      Top             =   1080
      Width           =   1095
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
      Caption         =   "frm_lmkybd.frx":0064
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0098
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":00B8
   End
   Begin HexUniControls.ctlUniButtonImageXP Command55 
      Height          =   600
      Left            =   2760
      TabIndex        =   65
      Top             =   5640
      Width           =   5895
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":00D4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":00FE
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":011E
   End
   Begin HexUniControls.ctlUniButtonImageXP Command66 
      Height          =   600
      Left            =   2040
      TabIndex        =   64
      Top             =   4920
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":013A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":015C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":017C
   End
   Begin HexUniControls.ctlUniButtonImageXP Command65 
      Height          =   600
      Left            =   2745
      TabIndex        =   63
      Top             =   4920
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0198
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":01BA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":01DA
   End
   Begin HexUniControls.ctlUniButtonImageXP Command64 
      Height          =   600
      Left            =   3435
      TabIndex        =   62
      Top             =   4920
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":01F6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0218
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0238
   End
   Begin HexUniControls.ctlUniButtonImageXP Command63 
      Height          =   600
      Left            =   4140
      TabIndex        =   61
      Top             =   4920
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0254
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0276
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0296
   End
   Begin HexUniControls.ctlUniButtonImageXP Command62 
      Height          =   600
      Left            =   4845
      TabIndex        =   60
      Top             =   4920
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":02B2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":02D4
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":02F4
   End
   Begin HexUniControls.ctlUniButtonImageXP Command61 
      Height          =   600
      Left            =   5535
      TabIndex        =   59
      Top             =   4920
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0310
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0332
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0352
   End
   Begin HexUniControls.ctlUniButtonImageXP Command60 
      Height          =   600
      Left            =   6240
      TabIndex        =   58
      Top             =   4920
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":036E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0390
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":03B0
   End
   Begin HexUniControls.ctlUniButtonImageXP Command59 
      Height          =   600
      Left            =   6945
      TabIndex        =   57
      Top             =   4920
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":03CC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":03EE
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":040E
   End
   Begin HexUniControls.ctlUniButtonImageXP Command58 
      Height          =   600
      Left            =   7635
      TabIndex        =   56
      Top             =   4920
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":042A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":044C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":046C
   End
   Begin HexUniControls.ctlUniButtonImageXP Command57 
      Height          =   600
      Left            =   8385
      TabIndex        =   55
      Top             =   4920
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0488
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":04AA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":04CA
   End
   Begin HexUniControls.ctlUniButtonImageXP Command54 
      Height          =   600
      Left            =   960
      TabIndex        =   54
      Top             =   4200
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":04E6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":050E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":052E
   End
   Begin HexUniControls.ctlUniButtonImageXP Command53 
      Height          =   600
      Left            =   2040
      TabIndex        =   53
      Top             =   4200
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":054A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":056C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":058C
   End
   Begin HexUniControls.ctlUniButtonImageXP Command52 
      Height          =   600
      Left            =   2745
      TabIndex        =   52
      Top             =   4200
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":05A8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":05CA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":05EA
   End
   Begin HexUniControls.ctlUniButtonImageXP Command51 
      Height          =   600
      Left            =   3435
      TabIndex        =   51
      Top             =   4200
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0606
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0628
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0648
   End
   Begin HexUniControls.ctlUniButtonImageXP Command50 
      Height          =   600
      Left            =   4140
      TabIndex        =   50
      Top             =   4200
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0664
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0686
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":06A6
   End
   Begin HexUniControls.ctlUniButtonImageXP Command49 
      Height          =   600
      Left            =   4845
      TabIndex        =   49
      Top             =   4200
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":06C2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":06E4
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0704
   End
   Begin HexUniControls.ctlUniButtonImageXP Command48 
      Height          =   600
      Left            =   5535
      TabIndex        =   48
      Top             =   4200
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0720
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0742
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0762
   End
   Begin HexUniControls.ctlUniButtonImageXP Command47 
      Height          =   600
      Left            =   6240
      TabIndex        =   47
      Top             =   4200
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":077E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":07A0
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":07C0
   End
   Begin HexUniControls.ctlUniButtonImageXP Command46 
      Height          =   600
      Left            =   6945
      TabIndex        =   46
      Top             =   4200
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":07DC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":07FE
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":081E
   End
   Begin HexUniControls.ctlUniButtonImageXP Command45 
      Height          =   600
      Left            =   7635
      TabIndex        =   45
      Top             =   4200
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":083A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":085C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":087C
   End
   Begin HexUniControls.ctlUniButtonImageXP Command44 
      Height          =   600
      Left            =   8385
      TabIndex        =   44
      Top             =   4200
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0898
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":08BA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":08DA
   End
   Begin HexUniControls.ctlUniButtonImageXP Command43 
      Height          =   600
      Left            =   9135
      TabIndex        =   43
      Top             =   4200
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":08F6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0918
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0938
   End
   Begin HexUniControls.ctlUniButtonImageXP Command42 
      Height          =   600
      Left            =   9885
      TabIndex        =   42
      Top             =   4200
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0954
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":097E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":099E
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_kybd 
      Height          =   735
      Left            =   2760
      TabIndex        =   41
      Top             =   960
      Width           =   5175
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_lmkybd.frx":09BA
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
      Tip             =   "frm_lmkybd.frx":09DA
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":09FA
   End
   Begin HexUniControls.ctlUniButtonImageXP Command41 
      Height          =   600
      Left            =   10680
      TabIndex        =   40
      Top             =   2760
      Width           =   960
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0A16
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0A3A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0A5A
   End
   Begin HexUniControls.ctlUniButtonImageXP Command40 
      Height          =   600
      Left            =   10600
      TabIndex        =   39
      Top             =   3480
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0A76
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0A98
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0AB8
   End
   Begin HexUniControls.ctlUniButtonImageXP Command39 
      Height          =   600
      Left            =   960
      TabIndex        =   38
      Top             =   3480
      Width           =   960
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0AD4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0AFA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0B1A
   End
   Begin HexUniControls.ctlUniButtonImageXP Command38 
      Height          =   600
      Left            =   2040
      TabIndex        =   37
      Top             =   3480
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0B36
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0B58
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0B78
   End
   Begin HexUniControls.ctlUniButtonImageXP Command37 
      Height          =   600
      Left            =   2745
      TabIndex        =   36
      Top             =   3480
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0B94
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0BB6
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0BD6
   End
   Begin HexUniControls.ctlUniButtonImageXP Command36 
      Height          =   600
      Left            =   3435
      TabIndex        =   35
      Top             =   3480
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0BF2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0C14
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0C34
   End
   Begin HexUniControls.ctlUniButtonImageXP Command35 
      Height          =   600
      Left            =   4140
      TabIndex        =   34
      Top             =   3480
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0C50
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0C72
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0C92
   End
   Begin HexUniControls.ctlUniButtonImageXP Command34 
      Height          =   600
      Left            =   4845
      TabIndex        =   33
      Top             =   3480
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0CAE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0CD0
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0CF0
   End
   Begin HexUniControls.ctlUniButtonImageXP Command33 
      Height          =   600
      Left            =   5535
      TabIndex        =   32
      Top             =   3480
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0D0C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0D2E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0D4E
   End
   Begin HexUniControls.ctlUniButtonImageXP Command32 
      Height          =   600
      Left            =   6240
      TabIndex        =   31
      Top             =   3480
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0D6A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0D8C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0DAC
   End
   Begin HexUniControls.ctlUniButtonImageXP Command31 
      Height          =   600
      Left            =   6945
      TabIndex        =   30
      Top             =   3480
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0DC8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0DEA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0E0A
   End
   Begin HexUniControls.ctlUniButtonImageXP Command30 
      Height          =   600
      Left            =   7635
      TabIndex        =   29
      Top             =   3480
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0E26
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0E48
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0E68
   End
   Begin HexUniControls.ctlUniButtonImageXP Command29 
      Height          =   600
      Left            =   8385
      TabIndex        =   28
      Top             =   3480
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0E84
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0EA6
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0EC6
   End
   Begin HexUniControls.ctlUniButtonImageXP Command28 
      Height          =   600
      Left            =   9135
      TabIndex        =   27
      Top             =   3480
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0EE2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0F04
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0F24
   End
   Begin HexUniControls.ctlUniButtonImageXP Command27 
      Height          =   600
      Left            =   9885
      TabIndex        =   26
      Top             =   3480
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0F40
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0F62
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0F82
   End
   Begin HexUniControls.ctlUniButtonImageXP Command26 
      Height          =   600
      Left            =   1320
      TabIndex        =   25
      Top             =   2040
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0F9E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":0FC0
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":0FE0
   End
   Begin HexUniControls.ctlUniButtonImageXP Command25 
      Height          =   600
      Left            =   2040
      TabIndex        =   24
      Top             =   2040
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":0FFC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":101E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":103E
   End
   Begin HexUniControls.ctlUniButtonImageXP Command24 
      Height          =   600
      Left            =   2745
      TabIndex        =   23
      Top             =   2040
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":105A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":107C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":109C
   End
   Begin HexUniControls.ctlUniButtonImageXP Command23 
      Height          =   600
      Left            =   3435
      TabIndex        =   22
      Top             =   2040
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":10B8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":10DA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":10FA
   End
   Begin HexUniControls.ctlUniButtonImageXP Command22 
      Height          =   600
      Left            =   4140
      TabIndex        =   21
      Top             =   2040
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":1116
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":1138
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":1158
   End
   Begin HexUniControls.ctlUniButtonImageXP Command21 
      Height          =   600
      Left            =   4845
      TabIndex        =   20
      Top             =   2040
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":1174
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":1196
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":11B6
   End
   Begin HexUniControls.ctlUniButtonImageXP Command20 
      Height          =   600
      Left            =   5535
      TabIndex        =   19
      Top             =   2040
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":11D2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":11F4
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":1214
   End
   Begin HexUniControls.ctlUniButtonImageXP Command19 
      Height          =   600
      Left            =   6240
      TabIndex        =   18
      Top             =   2040
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":1230
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":1252
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":1272
   End
   Begin HexUniControls.ctlUniButtonImageXP Command18 
      Height          =   600
      Left            =   6945
      TabIndex        =   17
      Top             =   2040
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":128E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":12B0
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":12D0
   End
   Begin HexUniControls.ctlUniButtonImageXP Command17 
      Height          =   600
      Left            =   7635
      TabIndex        =   16
      Top             =   2040
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":12EC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":130E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":132E
   End
   Begin HexUniControls.ctlUniButtonImageXP Command16 
      Height          =   600
      Left            =   8385
      TabIndex        =   15
      Top             =   2040
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":134A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":136C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":138C
   End
   Begin HexUniControls.ctlUniButtonImageXP Command15 
      Height          =   600
      Left            =   9135
      TabIndex        =   14
      Top             =   2040
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":13A8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":13CA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":13EA
   End
   Begin HexUniControls.ctlUniButtonImageXP Command14 
      Height          =   600
      Left            =   9885
      TabIndex        =   13
      Top             =   2040
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":1406
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":1428
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":1448
   End
   Begin HexUniControls.ctlUniButtonImageXP Command1 
      Height          =   600
      Left            =   1320
      TabIndex        =   12
      Top             =   2760
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":1464
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":1486
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":14A6
   End
   Begin HexUniControls.ctlUniButtonImageXP Command2 
      Height          =   600
      Left            =   2040
      TabIndex        =   11
      Top             =   2760
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":14C2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":14E4
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":1504
   End
   Begin HexUniControls.ctlUniButtonImageXP Command3 
      Height          =   600
      Left            =   2745
      TabIndex        =   10
      Top             =   2760
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":1520
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":1542
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":1562
   End
   Begin HexUniControls.ctlUniButtonImageXP Command4 
      Height          =   600
      Left            =   3435
      TabIndex        =   9
      Top             =   2760
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":157E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":15A0
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":15C0
   End
   Begin HexUniControls.ctlUniButtonImageXP Command5 
      Height          =   600
      Left            =   4140
      TabIndex        =   8
      Top             =   2760
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":15DC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":15FE
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":161E
   End
   Begin HexUniControls.ctlUniButtonImageXP Command6 
      Height          =   600
      Left            =   4845
      TabIndex        =   7
      Top             =   2760
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":163A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":165C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":167C
   End
   Begin HexUniControls.ctlUniButtonImageXP Command7 
      Height          =   600
      Left            =   5535
      TabIndex        =   6
      Top             =   2760
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":1698
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":16BA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":16DA
   End
   Begin HexUniControls.ctlUniButtonImageXP Command8 
      Height          =   600
      Left            =   6240
      TabIndex        =   5
      Top             =   2760
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":16F6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":1718
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":1738
   End
   Begin HexUniControls.ctlUniButtonImageXP Command9 
      Height          =   600
      Left            =   6945
      TabIndex        =   4
      Top             =   2760
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":1754
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":1776
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":1796
   End
   Begin HexUniControls.ctlUniButtonImageXP Command10 
      Height          =   600
      Left            =   7635
      TabIndex        =   3
      Top             =   2760
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":17B2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":17D4
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":17F4
   End
   Begin HexUniControls.ctlUniButtonImageXP Command11 
      Height          =   600
      Left            =   8385
      TabIndex        =   2
      Top             =   2760
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":1810
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":1832
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":1852
   End
   Begin HexUniControls.ctlUniButtonImageXP Command12 
      Height          =   600
      Left            =   9135
      TabIndex        =   1
      Top             =   2760
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":186E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":1890
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":18B0
   End
   Begin HexUniControls.ctlUniButtonImageXP Command13 
      Height          =   600
      Left            =   9885
      TabIndex        =   0
      Top             =   2760
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_lmkybd.frx":18CC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_lmkybd.frx":18EE
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":190E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   240
      Top             =   7800
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7095
      FormDesignWidth =   12450
   End
   Begin HexUniControls.ctlUniImage ctlUniImage1 
      Height          =   900
      Left            =   240
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1588
      Picture         =   "frm_lmkybd.frx":192A
      Tip             =   "frm_lmkybd.frx":6831
      Enabled         =   -1  'True
      Border          =   0   'False
      BackColor       =   14737632
      BorderColor     =   -1
      RoundedBorders  =   -1  'True
      Stretch         =   -1  'True
      XTransp         =   0
      YTransp         =   0
      MousePointer    =   0
      MouseIcon       =   "frm_lmkybd.frx":6851
   End
   Begin HexUniControls.ctlUniFormCaption uniTitle 
      Left            =   120
      Top             =   4320
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_lmkybd.frx":686D
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   840
      X2              =   11880
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   840
      X2              =   11880
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   11900
      X2              =   11900
      Y1              =   1920
      Y2              =   6600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   840
      X2              =   840
      Y1              =   1920
      Y2              =   6600
   End
End
Attribute VB_Name = "frm_lmkybd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
txt_kybd.Text = txt_kybd.Text & Command1.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command10_Click()
txt_kybd.Text = txt_kybd.Text & Command10.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command11_Click()
txt_kybd.Text = txt_kybd.Text & Command11.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command12_Click()
txt_kybd.Text = txt_kybd.Text & Command12.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command13_Click()
txt_kybd.Text = txt_kybd.Text & Command13.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command14_Click()
txt_kybd.Text = txt_kybd.Text & Command14.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command15_Click()
txt_kybd.Text = txt_kybd.Text & Command15.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command16_Click()
txt_kybd.Text = txt_kybd.Text & Command16.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command17_Click()
txt_kybd.Text = txt_kybd.Text & Command17.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command18_Click()
txt_kybd.Text = txt_kybd.Text & Command18.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command2_Click()
txt_kybd.Text = txt_kybd.Text & Command2.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command20_Click()
txt_kybd.Text = txt_kybd.Text & Command20.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command21_Click()
txt_kybd.Text = txt_kybd.Text & Command21.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command22_Click()
txt_kybd.Text = txt_kybd.Text & Command22.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command23_Click()
txt_kybd.Text = txt_kybd.Text & Command23.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command24_Click()
txt_kybd.Text = txt_kybd.Text & Command24.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command25_Click()
txt_kybd.Text = txt_kybd.Text & Command25.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command26_Click()
txt_kybd.Text = txt_kybd.Text & Command26.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command27_Click()
txt_kybd.Text = txt_kybd.Text & Command27.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command28_Click()
txt_kybd.Text = txt_kybd.Text & Command28.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command29_Click()
txt_kybd.Text = txt_kybd.Text & Command29.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command3_Click()
txt_kybd.Text = txt_kybd.Text & Command3.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command30_Click()
txt_kybd.Text = txt_kybd.Text & Command30.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command31_Click()
txt_kybd.Text = txt_kybd.Text & Command31.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command32_Click()
txt_kybd.Text = txt_kybd.Text & Command32.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command33_Click()
txt_kybd.Text = txt_kybd.Text & Command33.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command34_Click()
txt_kybd.Text = txt_kybd.Text & Command34.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command35_Click()
txt_kybd.Text = txt_kybd.Text & Command35.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command36_Click()
txt_kybd.Text = txt_kybd.Text & Command36.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command37_Click()
txt_kybd.Text = txt_kybd.Text & Command37.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command38_Click()
txt_kybd.Text = txt_kybd.Text & Command38.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command39_Click()
txt_kybd.Text = txt_kybd.Text & vbTab
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command4_Click()
txt_kybd.Text = txt_kybd.Text & Command4.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command40_Click()
txt_kybd.Text = txt_kybd.Text & Command40.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command41_Click()
Dim tempString As String
Dim lenstring As Integer
Dim tempstring2 As String
tempString = txt_kybd.Text
lenstring = Len(tempString)
If lenstring = 0 Then
    Exit Sub
End If
tempstring2 = Mid(tempString, 1, (lenstring - 1))
txt_kybd.Text = tempstring2
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command42_Click()
txt_kybd.Text = txt_kybd.Text & vbEnter
Clipboard.Clear
Clipboard.SetText txt_kybd.Text

End Sub

Private Sub Command43_Click()
txt_kybd.Text = txt_kybd.Text & Command43.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command44_Click()
txt_kybd.Text = txt_kybd.Text & Command44.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command45_Click()
txt_kybd.Text = txt_kybd.Text & Command45.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command46_Click()
txt_kybd.Text = txt_kybd.Text & Command46.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command47_Click()
txt_kybd.Text = txt_kybd.Text & Command47.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command48_Click()
txt_kybd.Text = txt_kybd.Text & Command48.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command49_Click()
txt_kybd.Text = txt_kybd.Text & Command49.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command5_Click()
txt_kybd.Text = txt_kybd.Text & Command5.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command50_Click()
txt_kybd.Text = txt_kybd.Text & Command50.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command51_Click()
txt_kybd.Text = txt_kybd.Text & Command51.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command52_Click()
txt_kybd.Text = txt_kybd.Text & Command52.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command53_Click()
txt_kybd.Text = txt_kybd.Text & Command53.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command54_Click()
txt_kybd.Text = txt_kybd.Text & Command54.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command55_Click()
txt_kybd.Text = txt_kybd.Text & " "
Clipboard.Clear
Clipboard.SetText txt_kybd.Text

End Sub

Private Sub Command56_Click()
txt_kybd.Text = ""
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command57_Click()
txt_kybd.Text = txt_kybd.Text & Command57.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command58_Click()
txt_kybd.Text = txt_kybd.Text & Command58.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command59_Click()
txt_kybd.Text = txt_kybd.Text & Command59.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command6_Click()
txt_kybd.Text = txt_kybd.Text & Command6.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command60_Click()
txt_kybd.Text = txt_kybd.Text & Command60.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command61_Click()
txt_kybd.Text = txt_kybd.Text & Command61.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command62_Click()
txt_kybd.Text = txt_kybd.Text & Command62.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command63_Click()
txt_kybd.Text = txt_kybd.Text & Command63.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command64_Click()
txt_kybd.Text = txt_kybd.Text & Command64.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command65_Click()
txt_kybd.Text = txt_kybd.Text & Command65.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command66_Click()
txt_kybd.Text = txt_kybd.Text & Command66.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command7_Click()
txt_kybd.Text = txt_kybd.Text & Command7.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command8_Click()
txt_kybd.Text = txt_kybd.Text & Command8.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub

Private Sub Command9_Click()
txt_kybd.Text = txt_kybd.Text & Command9.Caption
Clipboard.Clear
Clipboard.SetText txt_kybd.Text
End Sub


Private Sub Command67_Click()
frm_guilevel.txt_pw.SelText = LCase(Clipboard.GetText())
frm_guilevel.tmr2.Enabled = True
frm_guilevel.lbl_wrongpw.Visible = False
frm_guilevel.tmr_pw.Enabled = False
frm_guilevel.pwpassed = False
Call frm_guilevel.checkpw2
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub








