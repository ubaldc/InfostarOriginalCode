VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{E6AC3E35-BC5B-44AC-B1A0-251A8A08AD90}#17.0#0"; "XYPlot.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_scan 
   Caption         =   "Sample Spectra"
   ClientHeight    =   10815
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12315
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
   Icon            =   "frm_scan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10815
   ScaleWidth      =   12315
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniCheckXP chk_adjust 
      Height          =   360
      Left            =   3120
      TabIndex        =   25
      Top             =   9360
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scan.frx":058A
      Enabled         =   -1  'True
      Align           =   0
      CheckBackColor  =   -2147483643
      CheckForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":060C
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":062C
      ShowFocus       =   -1  'True
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_gridOff 
      Height          =   650
      Left            =   2136
      TabIndex        =   20
      Top             =   9840
      Width           =   1900
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_scan.frx":0648
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0678
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0698
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cursorOff 
      Height          =   650
      Left            =   120
      TabIndex        =   18
      Top             =   9840
      Width           =   1900
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_scan.frx":06B4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":06E8
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0708
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_unzoom 
      Height          =   650
      Left            =   4152
      TabIndex        =   21
      Top             =   9840
      Width           =   1900
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_scan.frx":0724
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0750
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0770
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_exit 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   10200
      TabIndex        =   0
      Top             =   9840
      Width           =   1900
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_scan.frx":078C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":07B4
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":07D4
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_print 
      Height          =   650
      Left            =   8184
      TabIndex        =   23
      Top             =   9840
      Width           =   1900
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_scan.frx":07F0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":081A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":083A
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_export 
      Height          =   650
      Left            =   6168
      TabIndex        =   22
      Top             =   9840
      Width           =   1900
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_scan.frx":0856
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":088A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":08AA
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cursorOn 
      Height          =   650
      Left            =   120
      TabIndex        =   17
      Top             =   9840
      Width           =   1900
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_scan.frx":08C6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":08F8
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0918
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_gridOn 
      Height          =   650
      Left            =   2136
      TabIndex        =   19
      Top             =   9840
      Width           =   1900
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_scan.frx":0934
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0962
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0982
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cursorLeft 
      Height          =   360
      Left            =   720
      TabIndex        =   14
      Top             =   9360
      Visible         =   0   'False
      Width           =   500
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_scan.frx":099E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":09C0
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":09E0
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cursorRight 
      Height          =   360
      Left            =   1440
      TabIndex        =   15
      Top             =   9360
      Visible         =   0   'False
      Width           =   500
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_scan.frx":09FC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0A1E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0A3E
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cursorPageLeft 
      Height          =   360
      Left            =   120
      TabIndex        =   13
      Top             =   9360
      Visible         =   0   'False
      Width           =   500
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_scan.frx":0A5A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0A7E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0A9E
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cursorPageRight 
      Height          =   360
      Left            =   2040
      TabIndex        =   16
      Top             =   9360
      Visible         =   0   'False
      Width           =   500
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_scan.frx":0ABA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0ADE
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0AFE
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_clear 
      Height          =   300
      Index           =   4
      Left            =   1440
      TabIndex        =   10
      Top             =   8430
      Width           =   500
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scan.frx":0B1A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0B3C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0B5C
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_clear 
      Height          =   300
      Index           =   3
      Left            =   1440
      TabIndex        =   8
      Top             =   8070
      Width           =   500
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scan.frx":0B78
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0B9A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0BBA
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_clear 
      Height          =   300
      Index           =   2
      Left            =   1440
      TabIndex        =   6
      Top             =   7710
      Width           =   500
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scan.frx":0BD6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0BF8
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0C18
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_clear 
      Height          =   300
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   7350
      Width           =   500
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scan.frx":0C34
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0C56
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0C76
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_clear 
      Height          =   300
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   6990
      Width           =   500
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scan.frx":0C92
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0CB4
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0CD4
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_load 
      Height          =   300
      Index           =   4
      Left            =   720
      TabIndex        =   9
      Top             =   8430
      Width           =   500
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scan.frx":0CF0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0D12
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0D32
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_load 
      Height          =   300
      Index           =   3
      Left            =   720
      TabIndex        =   7
      Top             =   8070
      Width           =   500
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scan.frx":0D4E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0D70
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0D90
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_load 
      Height          =   300
      Index           =   2
      Left            =   720
      TabIndex        =   5
      Top             =   7710
      Width           =   500
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scan.frx":0DAC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0DCE
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0DEE
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_load 
      Height          =   300
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   7350
      Width           =   500
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scan.frx":0E0A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0E2C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0E4C
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_load 
      Height          =   300
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   6990
      Width           =   500
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scan.frx":0E68
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0E8A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0EAA
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_load 
      Height          =   300
      Index           =   5
      Left            =   720
      TabIndex        =   11
      Top             =   8790
      Width           =   500
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scan.frx":0EC6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0EE8
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0F08
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_clear 
      Height          =   300
      Index           =   5
      Left            =   1440
      TabIndex        =   12
      Top             =   8790
      Width           =   500
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scan.frx":0F24
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_scan.frx":0F46
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0F66
   End
   Begin XYPlotGraph.XYPlot XYPlot1 
      Height          =   6135
      Left            =   120
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   10821
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   11640
      Top             =   9360
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   10815
      FormDesignWidth =   12315
   End
   Begin HexUniControls.ctlUniLabel lbl_scanDateTime 
      Height          =   325
      Index           =   4
      Left            =   2160
      Top             =   8430
      Width           =   2700
      _ExtentX        =   4763
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
      Caption         =   "frm_scan.frx":0F82
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":0FA2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":0FC2
   End
   Begin HexUniControls.ctlUniLabel lbl_scanDateTime 
      Height          =   325
      Index           =   3
      Left            =   2160
      Top             =   8070
      Width           =   2700
      _ExtentX        =   4763
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
      Caption         =   "frm_scan.frx":0FDE
      BackColor       =   -2147483633
      ForeColor       =   0
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":0FFE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":101E
   End
   Begin HexUniControls.ctlUniLabel lbl_scanDateTime 
      Height          =   325
      Index           =   2
      Left            =   2160
      Top             =   7710
      Width           =   2700
      _ExtentX        =   4763
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
      Caption         =   "frm_scan.frx":103A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":105A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":107A
   End
   Begin HexUniControls.ctlUniLabel lbl_scanDateTime 
      Height          =   325
      Index           =   1
      Left            =   2160
      Top             =   7350
      Width           =   2700
      _ExtentX        =   4763
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
      Caption         =   "frm_scan.frx":1096
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":10B6
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":10D6
   End
   Begin HexUniControls.ctlUniLabel lbl_cursorXValue 
      Height          =   375
      Left            =   9600
      Top             =   6240
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "frm_scan.frx":10F2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":1112
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":1132
   End
   Begin HexUniControls.ctlUniLabel Label4 
      Height          =   255
      Left            =   240
      Top             =   6240
      Width           =   1575
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scan.frx":114E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":1188
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":11A8
   End
   Begin HexUniControls.ctlUniLabel lbl_scanDateTime 
      Height          =   325
      Index           =   0
      Left            =   2160
      Top             =   6990
      Width           =   2700
      _ExtentX        =   4763
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
      Caption         =   "frm_scan.frx":11C4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":11E4
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":1204
   End
   Begin HexUniControls.ctlUniLabel lbl_cursorYValue 
      Height          =   375
      Index           =   4
      Left            =   6840
      Top             =   6240
      Width           =   1095
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
      Caption         =   "frm_scan.frx":1220
      BackColor       =   -2147483633
      ForeColor       =   0
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":1240
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":1260
   End
   Begin HexUniControls.ctlUniLabel lbl_cursorYValue 
      Height          =   375
      Index           =   3
      Left            =   5640
      Top             =   6240
      Width           =   1095
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
      Caption         =   "frm_scan.frx":127C
      BackColor       =   -2147483633
      ForeColor       =   0
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":129C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":12BC
   End
   Begin HexUniControls.ctlUniLabel lbl_cursorYValue 
      Height          =   375
      Index           =   2
      Left            =   4440
      Top             =   6240
      Width           =   1095
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
      Caption         =   "frm_scan.frx":12D8
      BackColor       =   -2147483633
      ForeColor       =   0
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":12F8
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":1318
   End
   Begin HexUniControls.ctlUniLabel lbl_cursorYValue 
      Height          =   375
      Index           =   1
      Left            =   3240
      Top             =   6240
      Width           =   1095
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
      Caption         =   "frm_scan.frx":1334
      BackColor       =   -2147483633
      ForeColor       =   0
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":1354
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":1374
   End
   Begin HexUniControls.ctlUniLabel lbl_cursorYValue 
      Height          =   375
      Index           =   0
      Left            =   2040
      Top             =   6240
      Width           =   1095
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
      Caption         =   "frm_scan.frx":1390
      BackColor       =   -2147483633
      ForeColor       =   0
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":13B0
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":13D0
   End
   Begin HexUniControls.ctlUniLabel lbl_comment 
      Height          =   330
      Index           =   4
      Left            =   5040
      Top             =   8430
      Width           =   7000
      _ExtentX        =   12356
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
      Caption         =   "frm_scan.frx":13EC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":140C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":142C
   End
   Begin HexUniControls.ctlUniLabel lbl_comment 
      Height          =   330
      Index           =   3
      Left            =   5040
      Top             =   8070
      Width           =   7000
      _ExtentX        =   12356
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
      Caption         =   "frm_scan.frx":1448
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":1468
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":1488
   End
   Begin HexUniControls.ctlUniLabel lbl_comment 
      Height          =   330
      Index           =   2
      Left            =   5040
      Top             =   7710
      Width           =   7000
      _ExtentX        =   12356
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
      Caption         =   "frm_scan.frx":14A4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":14C4
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":14E4
   End
   Begin HexUniControls.ctlUniLabel lbl_comment 
      Height          =   330
      Index           =   1
      Left            =   5040
      Top             =   7350
      Width           =   7000
      _ExtentX        =   12356
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
      Caption         =   "frm_scan.frx":1500
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":1520
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":1540
   End
   Begin HexUniControls.ctlUniLabel lbl_comment 
      Height          =   330
      Index           =   0
      Left            =   5040
      Top             =   6990
      Width           =   7000
      _ExtentX        =   12356
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
      Caption         =   "frm_scan.frx":155C
      BackColor       =   -2147483633
      ForeColor       =   0
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":157C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":159C
   End
   Begin HexUniControls.ctlUniLabel Label2 
      Height          =   255
      Left            =   1425
      Top             =   6630
      Width           =   495
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scan.frx":15B8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":15E2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":1602
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   255
      Left            =   705
      Top             =   6630
      Width           =   495
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scan.frx":161E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":1646
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":1666
   End
   Begin HexUniControls.ctlUniLabel lbl_comment 
      Height          =   330
      Index           =   5
      Left            =   5040
      Top             =   8790
      Width           =   7000
      _ExtentX        =   12356
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
      Caption         =   "frm_scan.frx":1682
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":16A2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":16C2
   End
   Begin HexUniControls.ctlUniLabel lbl_cursorYValue 
      Height          =   375
      Index           =   5
      Left            =   8040
      Top             =   6240
      Width           =   1095
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
      Caption         =   "frm_scan.frx":16DE
      BackColor       =   -2147483633
      ForeColor       =   0
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   1
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":16FE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":171E
   End
   Begin HexUniControls.ctlUniLabel lbl_scanDateTime 
      Height          =   330
      Index           =   5
      Left            =   2160
      Top             =   8790
      Width           =   2700
      _ExtentX        =   4763
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
      Caption         =   "frm_scan.frx":173A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":175A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":177A
   End
   Begin HexUniControls.ctlUniLabel Label3 
      Height          =   375
      Left            =   10800
      Top             =   6240
      Width           =   615
      _ExtentX        =   1085
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
      Caption         =   "frm_scan.frx":1796
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":17BA
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":17DA
   End
   Begin HexUniControls.ctlUniLabel Label5 
      Height          =   375
      Left            =   9240
      Top             =   6240
      Width           =   255
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_scan.frx":17F6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_scan.frx":1818
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_scan.frx":1838
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   11040
      Top             =   9000
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
      Left            =   11040
      Top             =   9360
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_scan.frx":1854
   End
End
Attribute VB_Name = "frm_scan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public spctoopen As String
Public nspc As Integer
Public ytypespc As String
Public clearplot As Boolean
Public jpg_toexport As String
Public nptsx As Integer

Private m_cursorLastState As Boolean
Private m_gridLastState As Boolean
Private m_plotIndx As Integer
Private m_spcFilename As String
Private m_lastSpcDir As String

Private Sub cmd_clear_Click(index As Integer)
  Dim nn As Integer
  Dim errStrg As String
  Dim rc As Boolean
  
  unity_main.errorstring = "Sample Spectra screen Clear button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  If (XYPlot1.SubsetLabel(index) <> "") Then
    lbl_cursorYValue(index).Caption = ""
   
    For nn = 0 To XYPlot1.MaxSubsets - 1
      If (Len(lbl_cursorYValue(nn)) <> 0) Then GoTo NO_CLEAR_X
    Next nn
    
    lbl_cursorXValue.Caption = ""
        
NO_CLEAR_X:
    rc = XYPlot1.ClearSpectrum(index, errStrg)
    
    If (rc = True) Then
      XYPlot1.SubsetLabel(index) = ""
      lbl_comment(index).Caption = ""
      lbl_scanDateTime(index).Caption = ""
    Else
      unity_main.errorstring = errStrg
      unity_main.write_error
    End If
  End If
End Sub

Private Sub cmd_cursorLeft_Click()
  
  unity_main.errorstring = "Sample Spectra screen Cursor Left button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  XYPlot1.CursorLeft
End Sub

Private Sub cmd_cursorOff_Click()
  
  unity_main.errorstring = "Sample Spectra screen Cursor Off button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Call XYPlot1.DisplayCursor(False)
End Sub

Private Sub cmd_cursorOn_Click()
  
  unity_main.errorstring = "Sample Spectra screen Cursor On button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Call XYPlot1.DisplayCursor(True)
End Sub

Private Sub cmd_cursorPageLeft_Click()
  
  unity_main.errorstring = "Sample Spectra screen Cursor Page Left button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  XYPlot1.CursorPageLeft
End Sub

Private Sub cmd_cursorPageRight_Click()
  
  unity_main.errorstring = "Sample Spectra screen Cursor Page Right button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  XYPlot1.CursorPageRight
End Sub

Private Sub cmd_CursorRight_Click()
  
  unity_main.errorstring = "Sample Spectra screen Cursor Right button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Call XYPlot1.CursorRight
End Sub

Private Sub cmd_exit_Click()
  
  Call unity_main.restart_loop(LOG_DBG_LEVEL3, "Sample Spectra screen Exit button selected")
  m_lastSpcDir = ""
  Unload frm_scan
End Sub

Private Sub cmd_Export_Click()
  Dim rc As Boolean
  
  unity_main.errorstring = "Sample Spectra screen Save Graph button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  unity_main.formfrom = 10
  jpg_toexport = ""
  frm_kybd.lbl_kybd.Caption = MLSupport.GSS("frm_scan", "lbl_kybd", "File Name for saved JPEG file")
  frm_kybd.Show 1
  
  If (jpg_toexport <> "") Then
    rc = XYPlot1.SaveJPGImage(JPGS_DIR & jpg_toexport)

    If (rc <> True) Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_scan", "errMsg1", "Error saving JPEG file, please confirm you entered a valid file name"), vbCritical
    End If
  End If
End Sub

Private Sub cmd_GridOff_Click()
  
  unity_main.errorstring = "Sample Spectra screen Grid Off button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Call XYPlot1.DisplayGrids(False)
End Sub

Private Sub cmd_GridOn_Click()

  unity_main.errorstring = "Sample Spectra screen Grid On button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Call XYPlot1.DisplayGrids(True)
End Sub

Private Sub cmd_load_Click(index As Integer)

  unity_main.errorstring = "Sample Spectra screen Load button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  m_plotIndx = index
  Call LoadSPC
End Sub

Sub LoadSPC()
  Dim dialog As New clsCommDialogs
  Dim fileDir As String
  Dim sFilter As String
  Dim dlgTitle As String
  Dim fileName As String
  Dim strlen As Integer
  Dim indx As Integer

  On Error GoTo BAD_FILE
  dialog.InitDialogs
  
  If (m_lastSpcDir = "") Then
    If (unity_main.m_saveDir <> "") Then
      fileDir = unity_main.m_saveDir
    Else
      fileDir = SPECTRA_DIR
    End If
  Else
    fileDir = m_lastSpcDir
  End If

  sFilter = ("SPC (*.spc)" & Chr(0) & "*.spc" & Chr(0))
  dlgTitle = MLSupport.GSS("frm_scan", "dlgTitle", "Select Spectrum File to Plot")
  fileName = dialog.ShowOpen(Me.hwnd, sFilter, dlgTitle, OFN_FILEMUSTEXIST + OFN_HIDEREADONLY + OFN_NOCHANGEDIR, fileDir)
 
  If (fileName <> "") Then
    strlen = Len(fileName)
    indx = InStrRev(fileName, "\")
    m_lastSpcDir = Left(fileName, indx)
    m_spcFilename = fileName
    Call plotunity
  End If
  
  Exit Sub
  
BAD_FILE:
  CWrap.ShowMessageBoxW MLSupport.GSS("frm_scan", "errMsg2", "Error selecting file, please confirm you selected a valid file"), vbCritical
End Sub

Private Sub cmd_Print_Click()
  
  unity_main.errorstring = "Sample Spectra screen Print button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  On Error GoTo goofedprint
  frm_scan.PrintForm
  GoTo printok
  
goofedprint:
  CWrap.ShowMessageBoxW MLSupport.GSS("frm_scan", "errMsg3", "Problem with spectrum printout, please check your printer settings and confirm proper printer selected"), vbCritical
printok:
End Sub

Sub plotunity()
  Dim spcIO As GSpcIOLib.GSPCio
  Dim gradXpt As Double
  Dim varXVals As Variant
  Dim varYVals As Variant
  Dim subFileIndx As Long
  Dim numSubfiles As Long
  Dim temptext, errStrg As String
  Dim i As Integer
  Dim errMsg As String
  Dim uniMsg As String
  Dim infFileDirName As String
  Dim dataComment As String
  Dim instrum As String
  Dim origin As String
  Dim dataTimestamp As String
  Dim userName As String
  Dim adjustXAxis As Boolean
  
  On Error Resume Next
  temptext = m_spcFilename
     
  For i = Len(temptext) To 1 Step -1
    If Mid(temptext, i, 1) = "\" Then
      Exit For
    End If
  Next i

  ' Open spectrum file
  If (LoadSpcFile(m_spcFilename, spcIO, numSubfiles, errStrg) = True) Then
    subFileIndx = 0
       
    ' Get spectrum data
    If (GetSpcFileSpectrumData(spcIO, subFileIndx, varXVals, varYVals, errStrg) = True) Then
      If (chk_adjust = 1) Then
        adjustXAxis = True
      Else
        adjustXAxis = False
      End If
      
      gradXpt = (spcIO.LastPoint - spcIO.FirstPoint + 1) / spcIO.NumPoints
      
      ' Plot spectrum data
      If (XYPlot1.PlotSpectrum2(m_plotIndx, spcIO.FirstPoint, spcIO.LastPoint, gradXpt, SubsetDataTypes.SS_VARIANT, vbNull, vbNull, varXVals, varYVals, adjustXAxis, errStrg) = True) Then
        XYPlot1.SubsetLabel(m_plotIndx) = Mid(temptext, i + 1)
        infFileDirName = Replace(m_spcFilename, SPC_FILE_EXT, SPC_INFO_FILE_EXT)
          
        ' Get any related Unicode info
        If (GetSpcFileUnicodeData(infFileDirName, dataComment, instrum, origin, dataTimestamp, userName) = True) Then
          lbl_comment(m_plotIndx).Caption = dataComment
          lbl_scanDateTime(m_plotIndx).Caption = dataTimestamp
        Else
          lbl_comment(m_plotIndx).Caption = spcIO.Comment
          lbl_scanDateTime(m_plotIndx).Caption = spcIO.Date
        End If
        
        Set spcIO = Nothing
        CloseSpcFile
        Exit Sub
      Else
        errMsg = "Error plotting spectrum file: " & m_spcFilename
        uniMsg = MLSupport.GGS_Params("errMsg6", "Error plotting spectrum file: %1", m_spcFilename)
      End If
    Else
      errMsg = "Error reading spectrum file: " & m_spcFilename
      uniMsg = MLSupport.GGS_Params("errMsg5", "Error reading from spectrum file: %1", m_spcFilename)
    End If
  Else
    errMsg = "Error opening spectrum file: " & m_spcFilename
    uniMsg = MLSupport.GGS_Params("errMsg4", "Error opening spectrum file: %1", m_spcFilename)
  End If
    
  Set spcIO = Nothing
  CloseSpcFile
  unity_main.errorstring = errMsg & ". " & errStrg
  unity_main.write_error
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub
Private Sub cmd_Unzoom_Click()
  
  unity_main.errorstring = "Sample Spectra screen Unzoom button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Call XYPlot1.UnZoom
End Sub

Private Sub Form_Load()
  Dim nn As Integer
  Dim pad As Double
  
  ' Apply language file to form
  MLSupport.ApplyToForm Me

#If ABBFT Then
  Label3.Caption = "1/cm"
  chk_adjust.Visible = False
#End If

  XYPlot1.AxisMinMaxPad = 3
  XYPlot1.FontSize = FontSizes.SMALL
  XYPlot1.MainTitle = ""
  XYPlot1.SubTitle = ""
  
#If ABBFT Then
  XYPlot1.XAxisLabel = MLSupport.GSS("XYPlot", "XAxisLabel2", "Wavenumber (1/cm)")
#Else
  XYPlot1.XAxisLabel = MLSupport.GSS("XYPlot", "XAxisLabel", "Wavelength (nm)")
#End If

  XYPlot1.YAxisLabel = MLSupport.GSS("XYPlot", "YAxisLabel3", "Absorbance")

#If ABBFT Then
  pad = (unity_main.m_mb3000.m_endWavenum - unity_main.m_mb3000.m_startWavenum) / 50
  Call XYPlot1.Initialize(unity_main.m_mb3000.m_startWavenum - 25, unity_main.m_mb3000.m_endWavenum + 25, unity_main.m_mb3000.m_waveNumIncr, _
                          True, False, True, True, LegendStyles.TWO_LINE, True)
  
#Else
  Call XYPlot1.Initialize(unity_main.m_minWvln, unity_main.m_maxWvln, MS11CfgData.wvlnIncr, _
                          True, False, True, True, LegendStyles.ONE_LINE, True)
#End If

  Call XYPlot1.SetupManualXAxis(5, 5)

  For nn = 0 To XYPlot1.MaxSubsets - 1
    lbl_cursorYValue(nn).ForeColor = XYPlot1.SubsetColor(nn)
    lbl_comment(nn).ForeColor = XYPlot1.SubsetColor(nn)
    lbl_scanDateTime(nn).ForeColor = XYPlot1.SubsetColor(nn)
    cmd_load(nn).ForeColor = XYPlot1.SubsetColor(nn)
    cmd_clear(nn).ForeColor = XYPlot1.SubsetColor(nn)
  Next nn
  
  cmd_gridOff.Visible = False
  cmd_gridOn.Visible = True
  cmd_cursorOff.Visible = False
  cmd_cursorOn.Visible = True
    
  
  If (Trim(unity_main.txtsamplename.Text) <> "") Then
    If (LCase(unity_main.m_saveIt) = "save") Then
      If (unity_main.m_repsAvg > 1) Then
        If (unity_main.m_saveReps = True) And (unity_main.repcounter = 0) Then
          m_spcFilename = unity_main.m_saveDir & unity_main.txtsamplename.Text & SPC_FILE_EXT
        Else
          ' Display last replicate scan
          m_spcFilename = unity_main.m_saveDir & LAST_ABSORB_SCAN_FILE & SPC_FILE_EXT
        End If
      Else
        m_spcFilename = unity_main.m_saveDir & unity_main.txtsamplename.Text & SPC_FILE_EXT
      End If
    Else
      ' Display last replicate scan
      m_spcFilename = unity_main.m_saveDir & LAST_ABSORB_SCAN_FILE & SPC_FILE_EXT
    End If
    
    m_plotIndx = 0
    plotunity
  End If
End Sub

Private Sub XYPlot1_CursorDisplayed(ByVal state As Boolean)
  Dim nn As Integer

  m_cursorLastState = state

  If (state = True) Then
    cmd_cursorOn.Visible = False
    cmd_cursorOff.Visible = True
    cmd_cursorLeft.Visible = True
    cmd_cursorPageLeft.Visible = True
    cmd_cursorPageRight.Visible = True
    cmd_cursorRight.Visible = True
  Else
    lbl_cursorXValue.Caption = ""
    
    For nn = 0 To XYPlot1.MaxSubsets - 1
      lbl_cursorYValue(nn).Caption = ""
    Next nn
  
    cmd_cursorOn.Visible = True
    cmd_cursorOff.Visible = False
    cmd_cursorLeft.Visible = False
    cmd_cursorPageLeft.Visible = False
    cmd_cursorPageRight.Visible = False
    cmd_cursorRight.Visible = False
  End If
End Sub

Private Sub XYPlot1_CursorMoved(ByVal cursorXVal As Double, cursorYVals As Variant)
  Dim formatStr As String
  Dim nn, cnt As Integer
  
  lbl_cursorXValue.Caption = FormatNumber(cursorXVal, 1, vbTrue, vbFalse, vbFalse)
  
  For nn = 0 To XYPlot1.MaxSubsets - 1
    If (cursorYVals(nn) = XYPlot1.NullYData) Then
      lbl_cursorYValue(nn).Caption = ""
    Else
      formatStr = FormatNumber(cursorYVals(nn), 5, vbTrue, vbFalse, vbFalse)
  
      If (cursorYVals(nn) < 0) Then
        cnt = 8
      Else
        cnt = 7
      End If
    
      lbl_cursorYValue(nn).Caption = Left(formatStr, cnt)
    End If
  Next nn
End Sub

Private Sub XYPlot1_GridsDisplayed(ByVal state As Boolean)

  m_gridLastState = state

  If (state = True) Then
    cmd_gridOff.Visible = True
    cmd_gridOn.Visible = False
  Else
    cmd_gridOff.Visible = False
    cmd_gridOn.Visible = True
  End If
End Sub








