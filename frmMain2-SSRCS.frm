VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{E6AC3E35-BC5B-44AC-B1A0-251A8A08AD90}#17.0#0"; "XYPlot.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Object = "{60589A73-E999-4E5A-945B-C64AF481B816}#14.1#0"; "InfoStarIPCServer.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{A96315EF-5AAC-4626-8EA0-0452C5C93C09}#1.0#0"; "SSRCSClient.ocx"
Begin VB.Form unity_main 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "InfoStar Main"
   ClientHeight    =   10560
   ClientLeft      =   195
   ClientTop       =   105
   ClientWidth     =   13035
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000040&
   Icon            =   "frmMain2-SSRCS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmMain2-SSRCS.frx":030A
   ScaleHeight     =   10560
   ScaleWidth      =   13035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin HexUniControls.ctlUniButtonImageXP cmd_verifyRef 
      Height          =   705
      Left            =   2000
      TabIndex        =   31
      Top             =   3170
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1244
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
      Caption         =   "frmMain2-SSRCS.frx":0BD4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":0C14
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":0C34
   End
   Begin HexUniControls.ctlUniImage img_ssrcsConnect 
      Height          =   705
      Left            =   9960
      Top             =   255
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1244
      Picture         =   "frmMain2-SSRCS.frx":0C50
      Tip             =   "frmMain2-SSRCS.frx":10A2
      Enabled         =   -1  'True
      Border          =   -1  'True
      BackColor       =   12632064
      BorderColor     =   -1
      RoundedBorders  =   -1  'True
      Stretch         =   -1  'True
      XTransp         =   0
      YTransp         =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":10C2
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_runBatch 
      Height          =   700
      Left            =   2000
      TabIndex        =   1
      Top             =   1630
      Width           =   1820
      _ExtentX        =   3201
      _ExtentY        =   1244
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
      Caption         =   "frmMain2-SSRCS.frx":10DE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":1110
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":1130
   End
   Begin VB.Timer tmr_autoSmplrCommRsp 
      Enabled         =   0   'False
      Left            =   5880
      Top             =   10440
   End
   Begin MSCommLib.MSComm msComm_autoSmplr 
      Left            =   5160
      Top             =   10440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      ParitySetting   =   2
   End
   Begin VB.Timer tmr_mb3000State 
      Enabled         =   0   'False
      Left            =   3480
      Top             =   10440
   End
   Begin MSWinsockLib.Winsock winsock_clientMB3000 
      Left            =   4560
      Top             =   10440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmr_mb3000CommRsp 
      Enabled         =   0   'False
      Left            =   3960
      Top             =   10440
   End
   Begin HexUniControls.ctlUniListBoxXP lst_last50MaxCols 
      Height          =   135
      Left            =   3000
      TabIndex        =   29
      Top             =   10440
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   238
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TrapTab         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":114C
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":116C
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   585
      ScaleWidth      =   3675
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5520
      Width           =   3700
      Begin VB.Image img_batchRpt 
         Height          =   480
         Left            =   3050
         Picture         =   "frmMain2-SSRCS.frx":1188
         Stretch         =   -1  'True
         Top             =   45
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image img_report 
         Height          =   480
         Left            =   50
         Picture         =   "frmMain2-SSRCS.frx":15CA
         Stretch         =   -1  'True
         Top             =   45
         Width           =   480
      End
      Begin VB.Image img_csv 
         Height          =   480
         Left            =   1250
         Picture         =   "frmMain2-SSRCS.frx":1C34
         Stretch         =   -1  'True
         Top             =   45
         Width           =   480
      End
      Begin VB.Image img_ticket 
         Height          =   480
         Left            =   2450
         Picture         =   "frmMain2-SSRCS.frx":2076
         Stretch         =   -1  'True
         Top             =   45
         Width           =   480
      End
      Begin VB.Image img_binocs 
         Height          =   480
         Left            =   650
         Picture         =   "frmMain2-SSRCS.frx":24B8
         Stretch         =   -1  'True
         Top             =   45
         Width           =   480
      End
      Begin VB.Image img_help 
         Height          =   480
         Left            =   1850
         Picture         =   "frmMain2-SSRCS.frx":28FA
         Stretch         =   -1  'True
         Top             =   45
         Width           =   480
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_ssrcsConnect 
      Height          =   700
      Left            =   1380
      TabIndex        =   10
      Top             =   4710
      Width           =   1200
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
      Caption         =   "frmMain2-SSRCS.frx":2D3C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":2D76
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":2D96
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_clientAppl 
      Height          =   700
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   4710
      Visible         =   0   'False
      Width           =   1200
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
      Caption         =   "frmMain2-SSRCS.frx":2DB2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":2DD2
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":2DF2
   End
   Begin HexUniControls.ctlUniLabel lbl_sampleName 
      Height          =   495
      Left            =   0
      Top             =   6120
      Width           =   1225
      _ExtentX        =   2170
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmMain2-SSRCS.frx":2E0E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      BorderColor     =   0
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":2E3C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":2E5C
   End
   Begin HexUniControls.ctlUniTextBoxXP txtsamplename 
      Height          =   450
      Left            =   1260
      TabIndex        =   12
      Top             =   6180
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   794
      BorderColor     =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmMain2-SSRCS.frx":2E78
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
      Tip             =   "frmMain2-SSRCS.frx":2E98
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":2EB8
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_ref 
      Height          =   700
      Left            =   2640
      TabIndex        =   8
      Top             =   3940
      Width           =   1200
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
      Caption         =   "frmMain2-SSRCS.frx":2ED4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":2EFA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":2F1A
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_bias 
      Height          =   700
      Left            =   1380
      TabIndex        =   5
      Top             =   3940
      Visible         =   0   'False
      Width           =   1200
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
      Caption         =   "frmMain2-SSRCS.frx":2F36
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":2F5E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":2F7E
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_exit 
      Height          =   700
      Left            =   120
      TabIndex        =   4
      Top             =   3940
      Width           =   1200
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
      Caption         =   "frmMain2-SSRCS.frx":2F9A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":2FC2
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":2FE2
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_options 
      Height          =   700
      Left            =   120
      TabIndex        =   3
      Top             =   3170
      Width           =   3700
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
      Caption         =   "frmMain2-SSRCS.frx":2FFE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":302C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":304C
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_sample 
      Height          =   700
      Left            =   120
      TabIndex        =   0
      Top             =   1630
      Width           =   3700
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
      Caption         =   "frmMain2-SSRCS.frx":3068
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":309E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":30BE
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_select 
      Height          =   700
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   3700
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
      Caption         =   "frmMain2-SSRCS.frx":30DA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":3116
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":3136
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_repacks 
      Height          =   700
      Left            =   2640
      TabIndex        =   11
      Top             =   4710
      Width           =   1200
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
      Caption         =   "frmMain2-SSRCS.frx":3152
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":3180
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":31A0
   End
   Begin VB.Timer tmr_sec1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   5040
   End
   Begin HexUniControls.ctlUniLabel lbl_cursorValue 
      Height          =   300
      Left            =   20
      Top             =   7190
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmMain2-SSRCS.frx":31BC
      BackColor       =   16777215
      ForeColor       =   -2147483630
      BorderColor     =   16777215
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":31DC
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":31FC
   End
   Begin HexUniControls.ctlUniLabel lbl_scanDateTime 
      Height          =   300
      Left            =   2280
      Top             =   7190
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmMain2-SSRCS.frx":3218
      BackColor       =   16777215
      ForeColor       =   -2147483630
      BorderColor     =   16777215
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":3238
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":3258
   End
   Begin XYPlotGraph.XYPlot XYPlot1 
      Height          =   3195
      Left            =   0
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7200
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   5636
   End
   Begin HexUniControls.ctlUniLabel lbl_opStatus 
      Height          =   765
      Left            =   4080
      Top             =   6360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1349
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmMain2-SSRCS.frx":3274
      BackColor       =   -2147483633
      ForeColor       =   16711680
      BorderColor     =   0
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":3294
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":32B4
   End
   Begin FPUSpreadADO.fpSpread ss_last50 
      Height          =   3195
      Left            =   5040
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7200
      Width           =   7905
      _Version        =   458752
      _ExtentX        =   13944
      _ExtentY        =   5636
      _StockProps     =   64
      ColHeaderDisplay=   0
      DAutoHeadings   =   0   'False
      DAutoSizeCols   =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   -2147483633
      MaxCols         =   133
      MaxRows         =   50
      OperationMode   =   1
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmMain2-SSRCS.frx":32D0
      UnitType        =   2
      UserResize      =   1
   End
   Begin HexUniControls.ctlUniLabel lbl_time 
      Height          =   255
      Left            =   11400
      Top             =   10080
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmMain2-SSRCS.frx":35BE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      BorderColor     =   0
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":35DE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":35FE
   End
   Begin HexUniControls.ctlUniLabel lbl_date 
      Height          =   255
      Left            =   8040
      Top             =   10110
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmMain2-SSRCS.frx":361A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      BorderColor     =   0
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":363A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":365A
   End
   Begin HexUniControls.ctlUniLabel lbl_backdate 
      Height          =   735
      Left            =   8520
      Top             =   6360
      Width           =   4455
      _ExtentX        =   7858
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
      Caption         =   "frmMain2-SSRCS.frx":3676
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      BorderColor     =   0
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":3696
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":36B6
   End
   Begin HexUniControls.ctlUniLabel lbl_sampleComment 
      Height          =   450
      Left            =   0
      Top             =   6720
      Width           =   1225
      _ExtentX        =   2170
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
      Caption         =   "frmMain2-SSRCS.frx":36D2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      BorderColor     =   0
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":3700
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":3720
   End
   Begin HexUniControls.ctlUniListBoxXP lst_qual 
      Height          =   540
      Left            =   1440
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   10920
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
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
      TrapTab         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":373C
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":375C
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin VB.Timer tmr_sample 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3840
      Top             =   4080
   End
   Begin VB.Timer tmr_ref 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3720
      Top             =   3000
   End
   Begin VB.Timer tmr_all 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3720
      Top             =   2280
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   1200
      Top             =   5040
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   10560
      FormDesignWidth =   13035
   End
   Begin HexUniControls.ctlUniListBoxXP lst_modtype 
      Height          =   645
      Left            =   4920
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   11280
      Visible         =   0   'False
      Width           =   1215
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
      Tip             =   "frmMain2-SSRCS.frx":3778
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":3798
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin HexUniControls.ctlUniListBoxXP lstresrat 
      Height          =   645
      Left            =   12240
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   11280
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
      Tip             =   "frmMain2-SSRCS.frx":37B4
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":37D4
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin HexUniControls.ctlUniListBoxXP lstrr2 
      Height          =   645
      Left            =   10920
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   11280
      Visible         =   0   'False
      Width           =   1095
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
      Tip             =   "frmMain2-SSRCS.frx":37F0
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":3810
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin HexUniControls.ctlUniListBoxXP lstrr 
      Height          =   645
      Left            =   9720
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   11280
      Visible         =   0   'False
      Width           =   1095
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
      Tip             =   "frmMain2-SSRCS.frx":382C
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":384C
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin HexUniControls.ctlUniListBoxXP lstmd 
      Height          =   645
      Left            =   8400
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   11280
      Visible         =   0   'False
      Width           =   1215
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
      Tip             =   "frmMain2-SSRCS.frx":3868
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":3888
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin HexUniControls.ctlUniListBoxXP preds 
      Height          =   645
      Left            =   6840
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   11280
      Visible         =   0   'False
      Width           =   1215
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
      Tip             =   "frmMain2-SSRCS.frx":38A4
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":38C4
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin HexUniControls.ctlUniFrameXP Frame1 
      Height          =   1575
      Left            =   0
      Top             =   0
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   2778
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmMain2-SSRCS.frx":38E0
      Enabled         =   -1  'True
      BackColor       =   12632064
      ForeColor       =   -2147483640
      Tip             =   "frmMain2-SSRCS.frx":3900
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":3920
      Begin HexUniControls.ctlUniImage ctlUniImage1 
         Height          =   1065
         Left            =   10920
         Top             =   15
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1879
         Picture         =   "frmMain2-SSRCS.frx":393C
         Tip             =   "frmMain2-SSRCS.frx":78CE0
         Enabled         =   -1  'True
         Border          =   0   'False
         BackColor       =   12632064
         BorderColor     =   -1
         RoundedBorders  =   -1  'True
         Stretch         =   -1  'True
         XTransp         =   0
         YTransp         =   0
         MousePointer    =   0
         MouseIcon       =   "frmMain2-SSRCS.frx":78D00
      End
      Begin HexUniControls.ctlUniLabel lblSysDateTime 
         Height          =   360
         Left            =   9240
         Top             =   1095
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmMain2-SSRCS.frx":78D1C
         BackColor       =   12632064
         ForeColor       =   -2147483630
         Alignment       =   2
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMain2-SSRCS.frx":78D62
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMain2-SSRCS.frx":78D82
      End
      Begin HexUniControls.ctlUniLabel lbl_movement 
         Height          =   425
         Left            =   3795
         Top             =   1080
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   741
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmMain2-SSRCS.frx":78D9E
         BackColor       =   12632064
         ForeColor       =   192
         BorderColor     =   12632064
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMain2-SSRCS.frx":78DBE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMain2-SSRCS.frx":78DDE
      End
      Begin HexUniControls.ctlUniLabel lblMovementTitle 
         Height          =   425
         Left            =   360
         Top             =   1080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   741
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmMain2-SSRCS.frx":78DFA
         BackColor       =   12632064
         ForeColor       =   16711680
         BorderColor     =   12632064
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMain2-SSRCS.frx":78E3A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMain2-SSRCS.frx":78E5A
      End
      Begin HexUniControls.ctlUniLabel lblSampleModeTitle 
         Height          =   425
         Left            =   720
         Top             =   675
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   741
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmMain2-SSRCS.frx":78E76
         BackColor       =   12632064
         ForeColor       =   16711680
         BorderColor     =   12632064
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMain2-SSRCS.frx":78EB6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMain2-SSRCS.frx":78ED6
      End
      Begin HexUniControls.ctlUniLabel lbl_samplemode 
         Height          =   420
         Left            =   3795
         Top             =   675
         Width           =   5800
         _ExtentX        =   10239
         _ExtentY        =   741
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmMain2-SSRCS.frx":78EF2
         BackColor       =   12632064
         ForeColor       =   192
         BorderColor     =   12632064
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMain2-SSRCS.frx":78F12
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMain2-SSRCS.frx":78F32
      End
      Begin HexUniControls.ctlUniLabel lblProd1 
         Height          =   495
         Left            =   3795
         Top             =   150
         Width           =   5800
         _ExtentX        =   10239
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmMain2-SSRCS.frx":78F4E
         BackColor       =   12632064
         ForeColor       =   192
         BorderColor     =   12632064
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMain2-SSRCS.frx":78F6E
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmMain2-SSRCS.frx":78F8E
      End
      Begin HexUniControls.ctlUniLabel lblProductTitle 
         Height          =   495
         Left            =   240
         Top             =   150
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmMain2-SSRCS.frx":78FAA
         BackColor       =   12632064
         ForeColor       =   16711680
         BorderColor     =   12632064
         Alignment       =   1
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmMain2-SSRCS.frx":78FE8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         BackColorOut    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMain2-SSRCS.frx":79008
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_start 
      Height          =   700
      Left            =   1380
      TabIndex        =   7
      Top             =   3940
      Visible         =   0   'False
      Width           =   1200
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
      Caption         =   "frmMain2-SSRCS.frx":79024
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":7904E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":7906E
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_stop 
      Height          =   700
      Left            =   1380
      TabIndex        =   6
      Top             =   3940
      Width           =   1200
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
      Caption         =   "frmMain2-SSRCS.frx":7908A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":790B2
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":790D2
   End
   Begin HexUniControls.ctlUniTextBoxXP txtsampcomment 
      Height          =   450
      Left            =   1260
      TabIndex        =   13
      Top             =   6720
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   794
      BorderColor     =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmMain2-SSRCS.frx":790EE
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
      Tip             =   "frmMain2-SSRCS.frx":7910E
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":7912E
   End
   Begin HexUniControls.ctlUniListBoxXP lstint 
      Height          =   1740
      Left            =   720
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   10440
      Visible         =   0   'False
      Width           =   500
      _ExtentX        =   873
      _ExtentY        =   3069
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
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
      TrapTab         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":7914A
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":7916A
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin HexUniControls.ctlUniListBoxXP lstslope 
      Height          =   1500
      Left            =   120
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   10440
      Visible         =   0   'False
      Width           =   500
      _ExtentX        =   873
      _ExtentY        =   2646
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
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
      TrapTab         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":79186
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":791A6
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin HexUniControls.ctlUniLabel lbl_miltime 
      Height          =   255
      Left            =   9600
      Top             =   3600
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frmMain2-SSRCS.frx":791C2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":791E2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":79202
   End
   Begin HexUniControls.ctlUniListBoxXP lst_nd 
      Height          =   1740
      Left            =   1320
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   10440
      Visible         =   0   'False
      Width           =   500
      _ExtentX        =   873
      _ExtentY        =   3069
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
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
      TrapTab         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":7921E
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":7923E
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin HexUniControls.ctlUniListBoxXP lst_pfexp 
      Height          =   1740
      Left            =   1920
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   10440
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   3069
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
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
      TrapTab         =   0   'False
      Tip             =   "frmMain2-SSRCS.frx":7925A
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frmMain2-SSRCS.frx":7927A
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin SSRCSCLIENTLib.SSRCSClient SSRCSClient 
      Height          =   1125
      Left            =   12360
      TabIndex        =   30
      Top             =   10440
      Visible         =   0   'False
      Width           =   3000
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   238
      _StockProps     =   0
   End
   Begin VB.Timer tmr_sec30 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   2640
      Top             =   5040
   End
   Begin FPUSpreadADO.fpSpread fpspread_pred 
      Height          =   4620
      Left            =   3960
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1635
      Width           =   9120
      _Version        =   458752
      _ExtentX        =   16095
      _ExtentY        =   8149
      _StockProps     =   64
      ColHeaderDisplay=   0
      DAutoHeadings   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   -2147483633
      MaxCols         =   6
      MaxRows         =   64
      OperationMode   =   1
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmMain2-SSRCS.frx":79296
      UnitType        =   2
      UserResize      =   1
   End
   Begin InfoStarIPCServer.IPCServer IPCServer1 
      Height          =   735
      Left            =   10800
      TabIndex        =   32
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   3720
      Top             =   5400
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
End
Attribute VB_Name = "unity_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'declaration of public variables
Public modlname As String   'model name
Public modlindex As Integer  'index of property in model file (for models with multiple properties)
Public tempval As Double
Public repcounter As Integer
Public sumavg As Double
Public tempreal As Double
Public mlrmodname As String
Public mlrnwls As Integer
Public modltype As Integer
Public fullmodelname As String
Public tableok As Boolean
Public tempmdist As Double
Public pukedonpred As Boolean
Public passstring As String
Public tempskew As Single
Public tempbias As Single
Public itbeok As Boolean
Public current_product As String
Public current_sampling As String
Public current_ini As String
Public val1 As Single ' = 1st value/running value
Public val2 As Single 'next value being operated on
Public curfunct As String 'current function being evaluated
Public calc_string As String
Public sec_value As Single
Public utiltoopen As Integer
Public pw_open As Boolean
Public formfrom As Integer  'tells kybd scrn what scrn to return txt to
Public varfrom As Integer   'used to tell kybd scrn where to return txt to
Public os_value As Integer  ' # of lines fron start of qic spreadsheet to start values
Public backdate As Date
Public gotscanname As Boolean  'use with doit in getreference (pass 1 get bane, pass 2 get scan)
Public run_min_gui As Boolean
Public slcal As String
Public calstar_enabled As Boolean 'no=hide all references to calstar
Public firstrep As Boolean 'sname04 after 1st scan / named goes to true
Public errorstring As String

' Instrument/Product Signature Settings
Public m_fileDevID As Long
Public m_fileVersion As String
Public m_smplTable As Long
Public m_sysScanMode As Long

' Instrument/Global Configuration File Settings
Public m_allowBias As Boolean
#If SSTAR Then
Public m_autoSmplrPort As Integer
Public m_batchRptPath As String
#End If
Public m_dbgLevel As LOG_DBG_LEVELS
Public m_darkScan As Boolean
Public m_darkSub As Boolean
Public m_enableGlobalName As Boolean
Public m_enableRunMode As Boolean
Public m_globalBaseCounter As Long
Public m_globalDate As String
Public m_globalDateCounter As Long
Public m_globalNameBase As String
Public m_globalNameMode As String
Public m_minWvln As Double
Public m_maxWvln As Double
Public m_intRefNScans As Integer
Public m_intRefPPT As Integer
Public m_intRefTimeout As Integer
Public m_intRefVerifyRemindTime As Integer
Public m_intRefVerifyTimeout As Integer
Public m_spectrixEnable As Boolean
Public m_sysSerialNum As String        ' instrument serial number
Public m_intRefVerifyAccumTime As Single ' accumulative time between operation for internal ref verification
Public m_dblAccumTime As Double
#If SSTAR Then
Public m_scanDir As SCAN_DIRECTIONS
Public m_useAutoSmplr As Boolean
#End If

' LIMS Configuration File Settings
Public pog_usefile As Integer
Public pogcomm As Integer
Public pogheader As Integer
Public pogpath As String
Public pogfile As String
Public pogappendornew As Integer
Public pogacceptreject As Integer
Public pogformat As String
Public pogusestart As Integer
Public pogstartchar As String
Public pogdelimtype As String
Public pogdelimchar As String
Public pogdatetime As Integer
Public pogdate As Integer
Public pogtime As Integer
Public pogSerialNum As Integer
Public pogcomment As Integer
Public pogproduct As Integer
Public pogmodelid As Integer
Public pogpropname As Integer
Public pogpropvalue As Integer
Public pogpropoutlier As Integer
Public pogmdist As Integer
Public pogsresid As Integer
Public pognd As Integer
Public pogintercept As Integer
Public pogslope As Integer
Public pogport As Integer
Public pogbps As Long
Public pogdatabits As Integer
Public pogstopbits As Single
Public pogflowctrl As String
Public pogparity As String
Public pogsampid As Integer
Public remoteproduct As Integer
Public limsinpath As String
Public limsinfile As String

' Ticket Printer Configuration File Settings
Public m_tktComment As Integer
Public m_tktDate As Integer
Public m_tktHeader1 As String
Public m_tktHeader2 As String
Public m_tktPostDelimiter As Integer
Public m_tktPostFF As Integer
Public m_tktPostLFs As Integer
Public m_tktPreDelimiter As Integer
Public m_tktPreFF As Integer
Public m_tktPreLFs As Integer
Public m_tktPrinterBold As Integer
Public m_tktPrinterFont As Integer
Public m_tktPrinterName As String
Public m_tktProduct As Integer
Public m_tktSampId As Integer
Public m_tktSerialNum As Integer
Public m_tktTime As Integer

' CSV Report Configuration File Settings
Public m_csvDataQuotes As Integer
Public m_csvDateTime As Integer
Public m_csvDate As Integer
Public m_csvTime As Integer
Public m_csvSerialNum As Integer
Public m_csvProdName As Integer
Public m_csvSampleID As Integer
Public m_csvSampleComment As Integer
Public m_csvPropName As Integer
Public m_csvPropValue As Integer
Public m_csvPropMDist As Integer
Public m_csvPropSResid As Integer
Public m_csvPropOutlier As Integer
Public m_csvPropND As Integer
Public m_csvPropIntercept As Integer
Public m_csvPropSlope As Integer

' Dynamic Report Configuration File Settings
Public m_rptAddHeader As Boolean
Public m_rptAddManPrefix As Boolean
Public m_rptAddManSuffix As Boolean
Public m_rptAddQuotes As Boolean
Public m_rptAddTrailer As Boolean
Public m_rptBaseCounter As Long
Public m_rptBaseName As String
Public m_rptDate As String
Public m_rptDateCounter As Long
Public m_rptFieldDelim As String
Public m_rptFileExt As String
Public m_rptFileFormat As String
Public m_rptFilePath As String
Public m_rptHdrNumFields As Integer
Public m_rptManPrefix As String
Public m_rptManSuffix As String
Public m_rptNameEntry As String
Public m_rptNameMode As String
Public m_rptRecNumFields As Integer
Public m_rptTrlNumFields As Integer
Public m_rptUsrNumFields As Integer
Public m_rptUsrPos As String

' Spectrum Treatment Configuration File Settings
Public m_enableTreatment As Boolean
Public m_enableSmooth As Boolean
Public m_endSmoothWvln As Double
Public m_endSmoothNumPts As Integer
Public m_progSmoothRate As Integer
Public m_saveSpectra As String
Public m_smoothNumPts As Integer
Public m_smoothTypeEnum As SMOOTH_TYPES
Public m_smoothType As String
Public m_startSmoothWvln As Double
Public m_startSmoothNumPts As Integer
Public m_useProgSmooth As Boolean

' Product Configuration File Settings
Public m_adapterType As String
Public m_alarmMD As Integer
Public m_alarmND As Integer
Public m_alarmRR As Integer
Public m_alarmProp As Integer
Public m_backFreq As REF_FREQS
Public m_bType As String
Public m_clrManualName As Boolean
Public m_clrUserInputs As Boolean
Public m_dayCounter As Long                 ' used in date-counter naming
Public m_doLIMS As Integer
Public m_hideValCol As Boolean
Public m_lastBackFreq As REF_FREQS
Public m_lastEndWvln As Double
Public m_lastStartWvln As Double
Public m_makePred As String
Public m_multiCupType As String
Public m_nameBase As String
Public m_nameCounter As Long
Public m_nameScanType As String
Public m_noOLVal As String
Public m_noPredVal As String
Public m_numModelVars As Integer            ' num of model variables
Public m_olFormat As Boolean                ' if true=pics if false = numbers
Public m_productName As String
Public m_repsAvg As Integer
#If SSTAR Then
Public m_rotateDir As TRAY_ROTATE_DIRS
Public m_rotateIndexSteps As Integer
Public m_rotateMoveMode As TRAY_ROTATE_MOVEMENTS
Public m_rotateSpeed As Integer
Public m_rotateStepSteps As Integer
#End If
Public m_saveCSV As Boolean
Public m_saveCSVFile As String
Public m_saveDir As String
Public m_saveDynRpt As Boolean
Public m_saveIt As String
Public m_savePredFile As String
Public m_savePredictions As Boolean
Public m_saveReps As Boolean                ' save indiv replicate scans
Public m_smplNScans As Integer
Public m_sNameMode As Integer               ' 0 = no save, 1 = enter, 2 = comment, 3 = counter,
                                            ' 4 = date_counter, 5 = Stat name each, 6 = stat counter
Public m_sType As String
Public m_trayNum As Integer
Public m_useExtRefTrayCfg As Boolean
Public m_useMIV As Boolean
Public m_valueBound As Integer              ' bounds reported values to max/min 0=act value 1=bound min 2 = bound max, 3 = both
Public m_writeTkt As Integer                ' 0 = no, 1 = always, 2 = on demand

' Product Ucal Model Settings
Public m_expSVFChanged As Boolean
Public m_numConstituents As Long
Public m_prodSVFChanged As Boolean
Public m_prdFileName As String
Public m_prdModelType As Boolean
Public m_stfFileName As String
Public m_stfFileValid As Boolean
Public m_stfMasterSerNum As String
Public m_svfEndWvln As Double
Public m_svfFileName As String
Public m_svfIsStd As Long
Public m_svfStartWvln As Double
Public m_svfWaveCnvtFlg As Boolean
Public m_svfWaveInc As Double

' System scan variables
#If SSTAR Then
Public m_intRefFileLoadFlg As Boolean
Public m_lastDrwrPos As DRWR_POSITIONS
Public m_trayStatus As Long
#End If
Public m_scanDataType As SCAN_DATA_TYPES
Public m_scanMode As Long
Public m_scanState As SCAN_STATES
Public m_scanTmrState As SCAN_TMR_STATES
Public m_scanTimestamp As String
Public m_scanDblTimestamp As Double

' Internal reference scan variables
Public m_intRefAutoScan As Boolean
Public m_intRefManualScan As Boolean
Public m_intRefPPTScan As Boolean
Public m_intRefPPTEndWvln As Double
Public m_intRefPPTStartWvln As Double
Public m_intRefPPTFileSetup As Boolean

' External reference qualification scan variables
Public m_extRefPPTAdapterType As String
Public m_extRefPPTAdaptIndx As Integer
Public m_extRefPPTEndWvln As Double
Public m_extRefPPTFileName As String
Public m_extRefPPTFileSetup As Boolean
Public m_extRefPPTMultiCupIndx As Integer
Public m_extRefPPTMultiCupType As String
Public m_extRefPPTNScans As Integer
#If SSTAR Then
Public m_extRefPPTRotateDir As TRAY_ROTATE_DIRS
Public m_extRefPPTRotateIndexSteps As Integer
Public m_extRefPPTRotateMoveMode As TRAY_ROTATE_MOVEMENTS
Public m_extRefPPTRotateSpeed As Integer
Public m_extRefPPTRotateStepSteps As Integer
#End If
Public m_extRefPPTScan As Boolean
Public m_extRefPPTStartWvln As Double
Public m_extRefPPTTrayNum As Integer

' External reference scan variables
Public m_extRefAdapterType As String
Public m_extRefAdaptIndx As Integer
Public m_extRefAutoScan As Boolean
Public m_extRefEndWvln As Double
Public m_extRefFileSetup As Boolean
Public m_extRefFileName As String
Public m_extRefManualScan As Boolean
Public m_extRefMultiCupIndx As Integer
Public m_extRefMultiCupType As String
Public m_extRefNScans As Integer
Public m_extRefPosition As Integer
Public m_extRefPPT As Integer
#If SSTAR Then
Public m_extRefRotateDir As TRAY_ROTATE_DIRS
Public m_extRefRotateIndexSteps As Integer
Public m_extRefRotateMoveMode As TRAY_ROTATE_MOVEMENTS
Public m_extRefRotateSpeed As Integer
Public m_extRefRotateStepSteps As Integer
#End If
Public m_extRefStartWvln As Double
Public m_extRefTimeout As Integer
Public m_extRefTimeoutSecs As Integer
Public m_extRefTimer As Integer
Public m_extRefTimerIgnore As Boolean
Public m_extRefTrayNum As Integer

' Offline reference scan variables
Public m_olRefFileName As String
Public m_olRefFileSetup As Boolean

' Remote scan variables
Public m_remoteRefScan As Boolean
Public m_remoteSmplScan As Boolean

' Product sample scan variables
Public m_smplAutoScan As Boolean
Public m_smplEndWvln As Double
Public m_smplManualScan As Boolean
Public m_smplPPTScan As Boolean
Public m_smplPPTEndWvln As Double
Public m_smplPPTFileName As String
Public m_smplPPTStartWvln As Double
Public m_smplPPT As Integer
Public m_smplRepacks As Integer
Public m_smplStartWvln As Double

' Auto sampler/batch scanning variables
Public m_batchRunFlg As Boolean
Public m_batchRptFile As String

' Internal reference calibration management variables
#If SSTAR Then
Public m_allowIntRefCalAccess As Boolean
Public m_calibFunc As CALIB_FUNC
Public m_intRefCalFlg As Boolean
Public m_intRefVerifyTimer As Single
Public m_intRefVerReminderCtr As Integer
Public m_intRefVerReminderFlg As Boolean
Public m_intRefVerReqdFlg As Boolean
Public m_sys2500 As Boolean
#End If

' Misc variables
Public m_acceptLims As Boolean
Public m_ansiErrMsg As String
Public m_badIniVal As Boolean
Public m_defltsLoaded As Boolean
Public m_iniString As String
Public m_instModel As String
Public m_uniErrMsg As String

Public m_remoteProdSelect As Integer
Public m_remoteProdName As String
Public m_remoteSocketIndx As Integer

Public m_netFWInstalled As Boolean
Public m_prdConstituent As String
Public m_prdEnabled As Boolean    ' true = PRD modeling enabled

#If SSRCS Then
Public m_ssrcsConnected As Boolean
Public m_ssrcsConnectTime As Integer
Public m_ssrcsIPAddr As String
Public m_ssrcsRspTime As Long
#End If

#If ABBFT Then
Public m_mb3000 As New clsMB3000
#Else
Private m_ms11srvGNEventQ As Collection
#End If

Public Sub system_startup(reconnectFlg As Boolean)
  Dim rc As Boolean
  Dim verOCXif As Long
  Dim errMsg, errCodeMsg As String
  Dim uniMsg As String


#If SSTAR Then
#If SSRCS Then
  unity_main.errorstring = "Connected to SpectraStar: " & unity_main.m_ssrcsIPAddr
  unity_main.write_error

  SSRCSClientError = SSRCSClient.GetOCXIFVer(verOCXif)
#Else
  verOCXif = MS11srv.verOCXif
#End If

  ' Check that InfoStar can interface with MS11srv.ocx
  If ((verOCXif And &HFFFFFF00) > MS11SRV_IF_VER) Then
    errCodeMsg = "V" & Hex(verOCXif) & " > V" & Hex(MS11SRV_IF_VER)
    errMsg = ("MS11srv.ocx interface version newer than InfoStar (" & errCodeMsg & ")")
    uniMsg = MLSupport.GGS_Params("MS11srv.errMsg2", "MS11srv.ocx interface version newer than InfoStar (%1)", errCodeMsg)
    GoTo StartupError2
  End If
  
  ' Get instrument configuration data
  rc = Get_MS11srv_Inst_Cfg()
 
  If (rc = False) Then
    errMsg = "Error getting SpectraStar configuration data"
    uniMsg = MLSupport.GSS("MS11srv", "errMsg6", "Error getting SpectraStar configuration data")
    GoTo StartupError
  End If
  
  unity_main.m_sysSerialNum = CStr(MS11CfgData.sysSerialNum)
  
  ' Get instrument default data
  rc = Get_MS11srv_Defaults()
 
  If (rc = False) Then
    errMsg = "Error getting SpectraStar default data"
    uniMsg = MLSupport.GSS("MS11srv", "errMsg7", "Error getting SpectraStar default data")
    GoTo StartupError
  End If
  
  
  ' Check that instrument has tray(s) configured
  If (MS11CfgData.nTrays > 0) Then
    rc = Get_MS11srv_Tray_Info
      
    ' Get instrument tray information
    If (rc = False) Then
      errMsg = "Error getting SpectraStar tray information"
      uniMsg = MLSupport.GSS("MS11srv", "errMsg8", "Error getting SpectraStar tray information")
      GoTo StartupError
    End If
  Else
    errMsg = "No trays configured within SpectraStar"
    uniMsg = MLSupport.GSS("MS11srv", "errMsg9", "No trays configured within SpectraStar")
    GoTo StartupError2
  End If
 
  ' Get current drawer position
  If (MS11CfgData.devID = DTID_DRAWER0) Or (MS11CfgData.devID = DTID_DRAWER1) Then
#If SSRCS Then
    SSRCSClientError = SSRCSClient.GetTrayStatus(m_trayStatus)
#Else
    unity_main.m_lastDrwrPos = MS11srv.trayStatus And &H300
#End If
    unity_main.m_lastDrwrPos = m_trayStatus And &H300
  Else
    ' Allow access to Internal Reference Calibration Management button if calibrated TW system
    If ((MS11CfgData.sysScanMode And &HC00) = &HC00) Then
      unity_main.m_allowIntRefCalAccess = True
    End If
  End If
  
#If SSRCS Then
  SSRCSClientError = SSRCSClient.SetupEvents(1, 1)
#End If
#End If
 
updateRegistry
loadAccumSettings
  ' Setup # of buttons supported for IPC clients
  IPCServer1.SetNumButtons cmd_clientAppl.Count
 
  ' Get default printer name
  On Error Resume Next
  frmReport.m_dfltPrinterName = "None"
  frmReport.m_dfltPrinterName = Printer.DeviceName
  
  chk_ucal_files_installed
  chk_net_framework_installed
  seeifcalstar

  build_inst_model
  init_defaults
  
#If ABBFT Then
  load_unitymain
  init_spectrum_plot
#Else
  init_spectrum_plot
  load_unitymain
#End If
  
  setup_spreadsheets_maxrows

#If SSTAR Then
  unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status10", "Ready for New Sample")
#End If
  
  If (reconnectFlg = False) Then
    ' Check if run mode enabled
    If (unity_main.m_enableRunMode = True) Then
      frm_guipw.loadguipw
      frm_guilevel.Visible = True
      frm_guilevel.pwpassed = False
    Else
      frm_guilevel.forcemaxgui
    End If
  Else
    Call unity_main.restart_loop(LOG_DBG_LEVEL3, "Connected to SpectraStar")
  End If

#If SSTAR Then
  clear_GN_eventQ
#End If
  Exit Sub
  
#If SSTAR Then
StartupError:
  ' Report MS11srv error codes
  Call Get_MS11_Errorcodes_Msg(errCodeMsg)
  errMsg = (errMsg & " " & errCodeMsg)
  uniMsg = (uniMsg & " " & errCodeMsg)
  
StartupError2:
  errMsg = (errMsg & ". InfoStar will shutdown automatically")
  unity_main.errorstring = errMsg
  unity_main.write_error
  uniMsg = MLSupport.GGS_Params("errMsg2", "%1. InfoStar will shutdown automatically", uniMsg)
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  unity_main.unloadallforms "unity_main"
  Unload Me
  End
#End If
End Sub

Public Sub updateRegistry()
  Dim AppName As String
  Dim section As String
  Dim key As String
  Dim setting As String
  Dim ifVer(1 To 4) As Integer
  Dim tmp As Long
  Dim nn As Long

  tmp = MS11CfgData.verOCXif
  ifVer(1) = CInt(tmp / &H1000000)
  tmp = tmp - (ifVer(1) * &H1000000)
  ifVer(2) = CInt(tmp / &H10000)
  tmp = tmp - (ifVer(2) * &H10000)
  ifVer(3) = CInt(tmp / &H100)
  ifVer(4) = tmp - (ifVer(3) * &H100)
  
   
  Dim sKey As String
  Dim sValue As String
  Dim vSetting As Variant
  Dim sType As Long
  
  sKey = "Software\UnityScientfic\unet\diags"
   
  RegistryFunc.CreateNewKey sKey, RegistryFunc.HKEY_CURRENT_USERX
  
  sValue = "SerialNum"
  sType = REG_SZ
  
  Dim keyType As Long
  
  keyType = RegistryFunc.HKEY_CURRENT_USERX
  vSetting = unity_main.m_sysSerialNum
  RegistryFunc.SetKeyValue keyType, sKey, sValue, vSetting, sType
  
  sValue = "FirmwareVer"
  sType = REG_SZ
  vSetting = StrConv(MS11CfgData.verFW, vbFromUnicode)
  RegistryFunc.SetKeyValue keyType, sKey, sValue, vSetting, sType
  
  sValue = "MinWave"
  vSetting = MS11CfgData.minWvln
  RegistryFunc.SetKeyValue keyType, sKey, sValue, vSetting, sType
  
  sValue = "MaxWave"
  vSetting = MS11CfgData.maxWvln
  RegistryFunc.SetKeyValue keyType, sKey, sValue, vSetting, sType
  
  sValue = "OcxVer"
  vSetting = (ifVer(1) & "." & ifVer(2) & "." & ifVer(3) & "." & ifVer(4))
  RegistryFunc.SetKeyValue keyType, sKey, sValue, vSetting, sType
  
  sValue = "DevId"
  vSetting = MS11CfgData.devID
  RegistryFunc.SetKeyValue keyType, sKey, sValue, vSetting, sType
  
  sValue = "SoftwareDir"
  vSetting = StrConv(MS11CfgData.swDriverDir, vbFromUnicode)
  RegistryFunc.SetKeyValue keyType, sKey, sValue, vSetting, sType
  
  


End Sub
Public Function check_ext_ref_ppt_file(extRefFilePPTName As String, spcFilename As String) As Boolean

  ' Check for external reference qualification file
  spcFilename = (REFERENCES_DIR & extRefFilePPTName & PPT_SCAN_FILE & SPC_FILE_EXT)
  check_ext_ref_ppt_file = CFile.st_FileExist(spcFilename)
End Function

Public Function check_int_ref_ppt_file(startWvln As Double, endWvln As Double, spcFilename As String) As Boolean

  ' Check for reference qualification file
  spcFilename = (REFERENCES_DIR & startWvln & "-" & endWvln & INT_REF_PPT_SCAN_FILE & SPC_FILE_EXT)
  check_int_ref_ppt_file = CFile.st_FileExist(spcFilename)
End Function

#If SSTAR Then
Public Sub clear_GN_eventQ()

  ' Remove and destroy anything on GN event queue
  If Not (m_ms11srvGNEventQ Is Nothing) Then
    While (m_ms11srvGNEventQ.Count <> 0)
      m_ms11srvGNEventQ.Remove 1
    Wend
  End If
End Sub
#End If

Public Function load_prod_file(prodFile As String, prodSelect As Boolean) As Boolean
  Dim sampling As String
  Dim tempstring As String
  Dim inString As String
  Dim varStr As Variant
  Dim prodName As String
  Dim tmpFile As String
  Dim tried_ini As String
  Dim lineCnt As Integer
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim rc As Boolean
  Dim errMsg As String
  Dim uniMsg As String

  unity_main.m_badIniVal = False
  load_prod_file = False
  
  ' Clear Main screen labels if loading selected product
  If (prodSelect = True) Then
    unity_main.lblProd1.Caption = ""
    unity_main.lbl_samplemode.Caption = ""
    unity_main.lbl_movement.Caption = ""
    clearpredtablefull
  End If

  If (prodFile = "") Then
    get_last_product
    
    If (unity_main.passstring = "") Then
      Exit Function
    End If
  Else
    unity_main.passstring = prodFile
  End If
    
  tried_ini = Trim(unity_main.passstring)
  prodFile = (PRODUCTS_CFG_DIR & Trim(unity_main.passstring))

  ' Setup products defaults before loading .ini file
  If (tried_ini = PROD_DFLTS_CFG_FILE) Then
    Call setup_prod_default_data(True)
  Else
    Call setup_prod_default_data(False)
  End If

  If (uniFile.st_FileExist(prodFile) = False) Then
    unity_main.errorstring = (prodFile & " file not found")
    unity_main.write_error
  
    If (tried_ini = PROD_DFLTS_CFG_FILE) Then
      frm_createdefault.Show 1
    Else
      uniMsg = MLSupport.GGS_Params("fileErrMsg6", "%1 file not found", prodFile)
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    End If
    
    ' Check if load product file due to remote appl. request
    If (unity_main.m_remoteProdSelect = 2) Then
      send_current_product unity_main.m_remoteSocketIndx, InfoStarStatusCodes.ProductNoFileStat
    End If
    
    GoTo LOAD_ERROR
  End If

  'If it made it here, the config file exists, now copy to temp file and load it
  On Error GoTo BAD_FILE
  tmpFile = (PRODUCTS_CFG_DIR & TMP_LOAD_PROD_CFG_FILE)
  uniFile.st_CopyFile prodFile, tmpFile

  frm_collect.m_numModelVars = 0
  
  If (uniFile.OpenFileRead(tmpFile) = False) Then GoTo BAD_FILE
  
  On Error GoTo NO_MODELS
  fEncoding = uniFile.ReadBOM

  lineCnt = 0
  
  ' Find first model info line in .ini file
  While Not (uniFile.EOF())
    lineCnt = lineCnt + 1
    
    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(inString)
    Else
      rc = uniFile.ReadUnicodeLine(inString)
    End If
      
    If (rc = False) Then GoTo BAD_FILE
      
    If Trim(inString) = "[analysis models]" Then
      While Not (uniFile.EOF())
        lineCnt = lineCnt + 1
      
        If (fEncoding = fe_ANSI) Then
          rc = uniFile.ReadAnsiLine(tempstring)
        Else
          rc = uniFile.ReadUnicodeLine(tempstring)
        End If
    
        If (rc = False) Then GoTo BAD_FILE
      
        ' Check if pre-V3.9.0 format
        If (InStr(tempstring, Chr(34) & "," & Chr(34)) = 0) Then
          varStr = Split(tempstring, ",")
        Else
          varStr = Split(tempstring, Chr(34) & "," & Chr(34))
        End If
        
        ' Determine number of variables for models
        frm_collect.m_numModelVars = UBound(varStr) + 1
      
        ' Chk if have proper # of parameters for property
        If (frm_collect.m_numModelVars < 14) And (frm_collect.m_numModelVars > 18) Then
          errMsg = (prodFile & " line " & CStr(lineCnt) & " had invalid number (" & CStr(frm_collect.m_numModelVars) & ") of property parameters")
          unity_main.errorstring = errMsg
          unity_main.write_error
          uniMsg = MLSupport.GGS_Params("unity_main.errMsg1", "%1 line %2 had invalid number (%3) of property parameters", prodFile, CStr(lineCnt), CStr(frm_collect.m_numModelVars))
          CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
          GoTo LOAD_ERROR
        End If
      Wend
    End If
  Wend
  
NO_MODELS:
  ' Reset file position to beginning
  uniFile.SetFilePos 0, eupp_FILE_BEGIN
  On Error GoTo BAD_FILE
  fEncoding = uniFile.ReadBOM
  
  lineCnt = 0
  lineCnt = lineCnt + 1
    
  If (fEncoding = fe_ANSI) Then
    rc = uniFile.ReadAnsiLine(inString)
  Else
    rc = uniFile.ReadUnicodeLine(inString)
  End If
      
  If (rc = False) Or (inString <> "[product type, sampling]") Then GoTo BAD_FILE

  While Not (uniFile.EOF())
    Select Case (inString)
      Case "[product type, sampling]"
        ' Get next line
        lineCnt = lineCnt + 1
  
        If (fEncoding = fe_ANSI) Then
          rc = uniFile.ReadAnsiLine(inString)
        Else
          rc = uniFile.ReadUnicodeLine(inString)
        End If
      
        If (rc = False) Then GoTo BAD_FILE
        
        varStr = Split(inString, ",")
        unity_main.m_productName = varStr(0)
        sampling = varStr(1)
        unity_main.current_ini = Trim(unity_main.passstring)
        
        lineCnt = lineCnt + 1
        
        If (fEncoding = fe_ANSI) Then
          rc = uniFile.ReadAnsiLine(inString)
        Else
          rc = uniFile.ReadUnicodeLine(inString)
        End If
      
        If (rc = False) Then GoTo BAD_FILE
        
      Case "[signature settings]"
        Call load_file_signature_vals(prodFile, uniFile, fEncoding, lineCnt, m_badIniVal)
        inString = unity_main.m_iniString
       
      Case "[analyzer settings]"
        Call load_prod_file_vals(prodFile, uniFile, fEncoding, lineCnt)
        inString = unity_main.m_iniString
        
        If (inString <> "") Then
          If (process_file_signature_vals(prodFile, True, m_badIniVal) = False) Then
            ' Check if load product file due to remote appl. request
            If (unity_main.m_remoteProdSelect = 2) Then
              send_current_product unity_main.m_remoteSocketIndx, InfoStarStatusCodes.ProductNotLoadStat
            End If
          
            GoTo LOAD_ERROR
          End If
        End If
        
        If (inString = "[analysis models]") Then
          GoTo PROCESS_VARS
        End If
        
      Case "[analysis models]"
PROCESS_VARS:
        Call process_prod_file_vars(prodFile, uniFile, fEncoding, lineCnt, prodSelect)
        
        If (frm_collect.m_numModelVars = 0) Then
          unity_main.fpspread_pred.MaxRows = MAX_NUM_PROPS
          GoTo FILE_PROCESSED
        End If
      Case Else
        GoTo BAD_FILE
    End Select
  Wend
  
FILE_PROCESSED:
  unity_main.passstring = unity_main.current_ini   '1/11/04
  
  ' Check if successfully loaded product defaults file
  If (unity_main.current_ini = PROD_DFLTS_CFG_FILE) Then
    unity_main.m_defltsLoaded = True
  End If
  
  ' Update main screen if loading selected product
  If (prodSelect = True) Then
    unity_main.lblProd1.Caption = unity_main.m_productName
    unity_main.current_product = unity_main.m_productName
    clearpredtable
    
    ' Check if user inputs configured to be used
    If (unity_main.m_useMIV = True) Then
      frm_scanname.lbl_prod.Caption = unity_main.m_productName
    End If
    
    send_current_product 0, InfoStarStatusCodes.GoodStat
    
#If ABBFT Then
    ' Signal to reconfigure interferometer
    unity_main.m_mb3000.m_cfgStatus = False
#End If
  End If
        
  frm_collect.lblprod.Caption = unity_main.m_productName
  unity_main.current_sampling = sampling
  
  ' Check if ini file had bad value
  If (unity_main.m_badIniVal = True) Then
    unity_main.errorstring = (prodFile & " had incompatible value(s). Updated file with proper/default value(s).")
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg7", "%1 had incompatible value(s). Updated file with proper/default value(s)", prodFile)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg3", "%1. Refer to system log for more details", uniMsg), vbOKOnly
    Call frm_collect.savescansettings(False)
    
    ' Check if load product file due to remote appl. request
    If (unity_main.m_remoteProdSelect = 2) Then
      send_current_product unity_main.m_remoteSocketIndx, InfoStarStatusCodes.ProductInvCfgStat
    End If
  End If
  
  uniFile.CloseFile
  uniFile.st_RmFile tmpFile
  load_prod_file = True
  Exit Function

BAD_FILE:
  If (lineCnt = 0) Then
    errMsg = (prodFile & " file open error." & Error$)
    uniMsg = MLSupport.GGS_Params("fileErrMsg2", "%1 file open error. %2", prodFile, Error$)
  Else
    errMsg = (prodFile & " file has error on line " & CStr(lineCnt) & ". " & Error$)
    uniMsg = MLSupport.GGS_Params("fileErrMsg3", "%1 file has error on line %2. %3", prodFile, CStr(lineCnt), Error$)
  End If
  
  unity_main.errorstring = errMsg
  unity_main.write_error
  
  ' Check if problem with product defaults file, make a new one.
  If (tried_ini = PROD_DFLTS_CFG_FILE) Then
    frm_createdefault.Show 1
  Else
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    
    ' Check if load product file due to remote appl. request
    If (unity_main.m_remoteProdSelect = 2) Then
      send_current_product unity_main.m_remoteSocketIndx, InfoStarStatusCodes.ProductNotLoadStat
    End If
  End If
  
LOAD_ERROR:
  uniFile.CloseFile
  unity_main.passstring = ""
  
  If (tmpFile <> "") And (uniFile.st_FileExist(tmpFile) = True) Then
    uniFile.st_RmFile tmpFile
  End If
End Function

Sub load_file_signature_vals(ByVal fileName As String, ByVal uniFile As clsUniFile, ByVal fEncoding As eFileEncoding, ByRef lineCnt As Integer, ByRef badIniVal)
  Dim inString As String
  Dim xx As String
  Dim tmpStrg As String
  Dim cfgVar As String
  Dim varVal As String
  Dim yy, strlen As Integer
  Dim rc As Boolean

  unity_main.m_iniString = ""

  ' Process each line in signature settings section
  While Not (uniFile.EOF())
    ' Read line from file
    On Error GoTo FILE_ERROR
    lineCnt = lineCnt + 1

    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(inString)
    Else
      rc = uniFile.ReadUnicodeLine(inString)
    End If
      
    If (rc = False) Then GoTo FILE_ERROR
          
    If (InStr(1, inString, "[") <> 0) Then
      unity_main.m_iniString = inString
      Exit Sub
    End If
          
    ' Get variable name and its value
    xx = InStr(1, inString, "=")
    yy = Len(inString)
    tmpStrg = Trim(Mid(inString, 1, xx - 1))
    cfgVar = LCase(tmpStrg)
    varVal = Trim(Mid(inString, xx + 1))
          
    ' Process value by variable name
    On Error GoTo BAD_INI_VALUE
    Select Case (cfgVar)
      Case "devid"
        unity_main.m_fileDevID = CLng(varVal)
      Case "smpltable"
        unity_main.m_smplTable = CLng(varVal)
      Case "scanmode"
        unity_main.m_sysScanMode = CLng(varVal)
      Case "version"
        unity_main.m_fileVersion = varVal
    End Select
  Wend
  
  Exit Sub
  
BAD_INI_VALUE:
  unity_main.errorstring = (fileName & " line " & CStr(lineCnt) & " had incompatible value. " & Trim(tmpStrg) & " - " & varVal & "; will use default value")
  unity_main.write_error
  badIniVal = True
  Resume Next
  
FILE_ERROR:
  unity_main.m_iniString = ""
End Sub

Public Function process_file_signature_vals(fileName As String, prodFile As Boolean, badIniVal As Boolean) As Boolean

  process_file_signature_vals = True
  
#If ABBFT Then
  ' Check for invalid signature device ID
  If (unity_main.m_fileDevID <> DTID_ABBFT) Then
    ' Check if virgin signature value
    If (unity_main.m_fileDevID = -1) Then
      unity_main.errorstring = (fileName & " had incompatible value. DevID was " & unity_main.m_fileDevID & "; updated to " & DTID_ABBFT)
      unity_main.m_fileDevID = DTID_ABBFT
      badIniVal = True
    Else
      unity_main.errorstring = (fileName & " had incompatible signature value. DevID was " & unity_main.m_fileDevID & "; Instrument " & DTID_ABBFT)
      process_file_signature_vals = False
    End If
    
    unity_main.write_error
  End If

  ' Check for invalid signature file version
  If (unity_main.m_fileVersion <> INFOSTAR_VER) Then
    unity_main.errorstring = (fileName & " had incompatible value. Version was " & unity_main.m_fileVersion & "; updated to " & INFOSTAR_VER)
    unity_main.write_error
    unity_main.m_fileVersion = INFOSTAR_VER
    badIniVal = True
  End If
  
  If (process_file_signature_vals = False) Then
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("unity_main.errMsg2", "%1 had incompatible signature; Please contact Unity Scientific technical Support!", fileName), vbCritical
    Exit Function
  End If
#Else
  ' Check for invalid signature device ID
  If (unity_main.m_fileDevID <> MS11CfgData.devID) Then
    ' Check if virgin signature value
    If (unity_main.m_fileDevID = -1) Then
      unity_main.errorstring = (fileName & " had incompatible value. DevID was " & unity_main.m_fileDevID & "; updated to " & MS11CfgData.devID)
      unity_main.m_fileDevID = MS11CfgData.devID
      badIniVal = True
    Else
      unity_main.errorstring = (fileName & " had incompatible signature value. DevID was " & unity_main.m_fileDevID & "; SpectraStar " & MS11CfgData.devID)
      process_file_signature_vals = False
    End If
    
    unity_main.write_error
  End If
  
  ' Check for invalid signature sample table
  If (unity_main.m_smplTable <> MS11CfgData.smplTblIX) Then
    ' Check if virgin signature value
    If (unity_main.m_smplTable = -1) Then
      unity_main.errorstring = (fileName & " had incompatible value. SmplTable was " & unity_main.m_smplTable & "; updated to " & MS11CfgData.smplTblIX)
      unity_main.m_smplTable = MS11CfgData.smplTblIX
      badIniVal = True
    Else
      unity_main.errorstring = (fileName & " had incompatible signature value. SmplTable was " & unity_main.m_smplTable & "; SpectraStar " & MS11CfgData.smplTblIX)
      process_file_signature_vals = False
    End If
    
    unity_main.write_error
  End If
  
  ' Check for invalid signature scan mode
  If (unity_main.m_sysScanMode <> MS11CfgData.sysScanMode) Then
    ' Check if virgin signature value
    If (unity_main.m_sysScanMode = -1) Then
      unity_main.errorstring = (fileName & " had incompatible value. ScanMode was " & unity_main.m_sysScanMode & "; updated to " & MS11CfgData.sysScanMode)
      unity_main.m_sysScanMode = MS11CfgData.sysScanMode
      badIniVal = True
    Else
      unity_main.errorstring = (fileName & " had incompatible signature value. ScanMode was " & unity_main.m_sysScanMode & "; SpectraStar " & MS11CfgData.sysScanMode)
      process_file_signature_vals = False
    End If
    
    unity_main.write_error
  End If

  ' Check for invalid signature file version
  If (unity_main.m_fileVersion <> INFOSTAR_VER) Then
    unity_main.errorstring = (fileName & " had incompatible value. Version was " & unity_main.m_fileVersion & "; updated to " & INFOSTAR_VER)
    unity_main.write_error
    unity_main.m_fileVersion = INFOSTAR_VER
    badIniVal = True
  End If
  
  If (process_file_signature_vals = False) Then
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("unity_main.errMsg2", "%1 had incompatible signature; Please contact Unity Scientific technical Support!", fileName), vbCritical
    Exit Function
  End If

  ' Check if loading product file
  If (prodFile = True) Then
    ' Check if invalid starting wavelength value
    If (frm_collect.m_smplStartWvln < unity_main.m_minWvln) Or (frm_collect.m_smplStartWvln > unity_main.m_maxWvln) Then
      unity_main.errorstring = (fileName & " had incompatible value. StartWvln was " & frm_collect.m_smplStartWvln & "; SpectraStar " & ProdDfltData.startWvln)
      unity_main.write_error
      process_file_signature_vals = False
    End If

    ' Check if invalid ending wavelength value
    If (frm_collect.m_smplEndWvln < unity_main.m_minWvln) Or (frm_collect.m_smplEndWvln > unity_main.m_maxWvln) Then
      unity_main.errorstring = (fileName & " had incompatible value. EndWvln was " & frm_collect.m_smplEndWvln & "; SpectraStar " & ProdDfltData.endWvln)
      unity_main.write_error
      process_file_signature_vals = False
    End If

    ' Check if ending <= starting wavelength value
    If (frm_collect.m_smplEndWvln <= frm_collect.m_smplStartWvln) Then
      unity_main.errorstring = (fileName & " had incompatible value. StartWvln was " & frm_collect.m_smplStartWvln & "; EndWvln was " & frm_collect.m_smplEndWvln)
      unity_main.write_error
      process_file_signature_vals = False
    End If
  End If
  
  If (process_file_signature_vals = False) Then
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("unity_main.errMsg3", "%1 had incompatible starting/ending wavelengths; Please contact Unity Scientific technical Support!", fileName), vbCritical
  End If
#End If
End Function

Public Sub get_last_product()
  Dim in1 As String
  Dim in2 As String
  Dim xx As String
  Dim yy As Integer
  Dim lineCnt As Integer
  Dim fileName As String
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim rc As Boolean
  Dim errMsg As String
  Dim uniMsg As String

  fileName = (CFG_DIR & LAST_PROD_CFG_FILE)

  ' Open file and get encoding
  If (uniFile.OpenFileRead(fileName) = True) Then
    On Error GoTo BAD_FILE
    fEncoding = uniFile.ReadBOM
    lineCnt = lineCnt + 1
     
    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(in1)
    Else
      rc = uniFile.ReadUnicodeLine(in1)
    End If
     
    If (rc = False) Then GoTo BAD_FILE
     
    lineCnt = lineCnt + 1
    
    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(in2)
    Else
      rc = uniFile.ReadUnicodeLine(in2)
    End If
     
    If (rc = False) Then GoTo BAD_FILE
  
    xx = InStr(1, in2, "=")
    yy = Len(in2)
    unity_main.passstring = Trim(Mid(in2, xx + 1))
    
    If (unity_main.passstring = "") Then GoTo BAD_FILE
  Else
BAD_FILE:
    If (lineCnt = 0) Then
      errMsg = (fileName & " file open error." & Error$)
      uniMsg = MLSupport.GGS_Params("fileErrMsg2", "%1 file open error. %2", fileName, Error$)
    Else
      errMsg = (fileName & " file has error on line " & CStr(lineCnt) & ". " & Error$)
      uniMsg = MLSupport.GGS_Params("fileErrMsg3", "%1 file has error on line %2. %3", fileName, CStr(lineCnt), Error$)
    End If
  
    unity_main.errorstring = errMsg
    unity_main.write_error
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    unity_main.passstring = ""
  End If
  
  uniFile.CloseFile
End Sub

Public Sub save_last_product(prodName As String)
  Dim uniFile As New clsUniFile
  Dim errMsg As String
  Dim uniMsg As String

  On Error GoTo FILE_ERROR
  
  If (uniFile.OpenFileWrite(CFG_DIR & LAST_PROD_CFG_FILE) = True) Then
    uniFile.WriteBOM fe_UTF16LE
    uniFile.WriteUnicodeLine "[Last Product Analyzed]"
    uniFile.WriteUnicodeLine "Last Product=" & prodName
    uniFile.Flush
  Else
FILE_ERROR:
    errMsg = (CFG_DIR & LAST_PROD_CFG_FILE & " file write error. " & Error$)
    unity_main.errorstring = errMsg
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", CFG_DIR & LAST_PROD_CFG_FILE, Error$)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  End If
  
  uniFile.CloseFile
End Sub

Public Function find_latest_svf_file(filePath As String, prodName As String, fileName As String, fileNum As Long) As Boolean
  Dim expFile As Boolean
  Dim dateStrg As String
  Dim filePattern As String
  Dim osf As clsSearchFiles
  Dim oc As Collection
  Dim ii As Integer
  Dim osfi As clsSearchFilesFInfo
  Dim hiTimestamp As Double
  Dim loTimestamp As Double
  Dim dTmp1, dTmp2 As Double
  Dim fileIndx As Integer
  Dim pos1, pos2 As Integer
  Dim numStrg As String

  ' Check if expansion SVF file
  If (prodName = EXPANSION_SVF_FILE) Then
    expFile = True
  End If
  
  ' Check if looking for product SVF file
  If (expFile = False) Then
    ' Build SVF file pattern to search for list of existing product SVF files
    dateStrg = CStr(Date)
    dateStrg = Replace(dateStrg, "/", "-")
    filePattern = (filePath & prodName & "_" & dateStrg & "_S*" & SVF_FILE_EXT)
  Else
    ' Build SVF file pattern to search for list of existing expansion SVF files
    filePattern = (filePath & EXPANSION_SVF_FILE & "_S*" & SVF_FILE_EXT)
  End If
  
  Set osf = New clsSearchFiles
  Set oc = osf.SearchPathPattern(filePattern, esfw_only_files, False, False, False, -1, False)

FIND_NEWEST_FILE:
  hiTimestamp = 0
  loTimestamp = 0
  
  ' Check if any files found
  If (oc.Count = 0) Then
    ' No file found, return first filename to create
    If (expFile = False) Then
      fileName = (prodName & "_" & dateStrg & "_S1" & SVF_FILE_EXT)
    Else
      fileName = (EXPANSION_SVF_FILE & "_S1" & SVF_FILE_EXT)
    End If
    
    find_latest_svf_file = False
  Else
    ' Find newest file
    For ii = 1 To oc.Count
      Set osfi = oc.Item(ii)
      
      If (osfi.dwHighDateTimeLastWrite < 0) Then
        dTmp1 = 4294967296# + osfi.dwHighDateTimeLastWrite
      Else
        dTmp1 = osfi.dwHighDateTimeLastWrite
      End If
      
      If (osfi.dwLowDateTimeLastWrite < 0) Then
        dTmp2 = 4294967296# + osfi.dwLowDateTimeLastWrite
      Else
        dTmp2 = osfi.dwLowDateTimeLastWrite
      End If
      
      If (dTmp1 > hiTimestamp) Or _
         ((dTmp1 = hiTimestamp) And (dTmp2 > loTimestamp)) Then
        hiTimestamp = dTmp1
        loTimestamp = dTmp2
        fileIndx = ii
      End If
    Next ii
  
    ' Save newest file name
    Set osfi = oc.Item(fileIndx)
    fileName = osfi.sFileName
    
    ' Get file subnumber
    pos1 = InStrRev(fileName, "_S")
    pos1 = pos1 + 2
    pos2 = InStrRev(fileName, ".")
    numStrg = Mid(fileName, pos1, pos2 - pos1)
    
    If (IsNumeric(numStrg) = True) Then
      fileNum = CLng(numStrg)
    Else
      oc.Remove fileIndx
      GoTo FIND_NEWEST_FILE
    End If
    
    find_latest_svf_file = True
  End If
End Function

Public Sub send_new_products_list(socketIndx As Integer)
  Dim numProducts As Integer
  Dim productNames() As String
  Dim ii As Integer

  numProducts = FRM_SEL_PRODUCT.LSTPRODUCTS.ListCount
  
  If (numProducts > 0) Then
    ReDim productNames(1 To numProducts)
  
    For ii = 1 To numProducts
      productNames(ii) = Trim(FRM_SEL_PRODUCT.LSTPRODUCTS.List(ii - 1))
    Next ii
  End If
  
  IPCServer1.NewProductsList socketIndx, InfoStarStatusCodes.GoodStat, numProducts, productNames
End Sub

Public Sub send_current_product(socketIndx As Integer, Status As InfoStarStatusCodes)
  Dim prodName As String
  Dim spectrumFolderName As String

  If (Status = InfoStarStatusCodes.GoodStat) Then
    prodName = lblProd1.Caption
    spectrumFolderName = unity_main.m_saveDir
  End If

  IPCServer1.NewCurrentProduct socketIndx, Status, prodName, spectrumFolderName
End Sub

Public Function chk_timeout(StartTime As Single, intvTimeSecs As Single) As Boolean
  Dim currTime As Single
  
  chk_timeout = False
  currTime = Timer

  If (currTime > StartTime) Then
    If (currTime > (StartTime + intvTimeSecs)) Then
      chk_timeout = True
    End If
  Else
    If (currTime > (StartTime + intvTimeSecs)) Then
      chk_timeout = True
    End If
  End If
End Function


Public Sub log_scan_status(ByVal dbgLevel As LOG_DBG_LEVELS, statMsg As String, uniMsg As String)
  
  frm_dbug.lst_status.AddItem uniMsg
  unity_main.errorstring = statMsg
  unity_main.write_error (dbgLevel)
End Sub

Public Sub write_error(Optional ByVal dbgLevel As LOG_DBG_LEVELS)
  Dim dbgLvl As LOG_DBG_LEVELS
  Dim fileName As String
  Dim uniMsg As String
  Dim uniFile As New clsUniFile

  If (IsMissing(dbgLevel)) Then
    dbgLvl = LOG_DBG_LEVEL0
  Else
    dbgLvl = dbgLevel
  End If
  
  ' Check if to write error/status message to log file
  If (dbgLvl <= unity_main.m_dbgLevel) Then
    On Error GoTo FILE_ERROR
    fileName = (LOGFILE_DIR & SYSTEM_LOG_FILE)
    
    If (uniFile.OpenFileAppend(fileName) = True) Then
      uniFile.WriteUnicodeLine (CStr(Date) & " , " & CStr(Time) & " , " & unity_main.errorstring)
      uniFile.Flush
    Else
FILE_ERROR:
      uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", fileName, Error$)
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    End If
    
    uniFile.CloseFile
  End If
End Sub

Public Sub checklogfile()
  Dim fileName As String
  Dim uniMsg As String
  Dim uniFile As New clsUniFile
  
  fileName = (LOGFILE_DIR & SYSTEM_LOG_FILE)
  
  ' If no file, create it and write header
  If (uniFile.st_FileExist(fileName) = False) Then
    On Error GoTo FILE_ERROR
    
    If (uniFile.OpenFileWrite(fileName) = True) Then
      uniFile.WriteBOM fe_UTF16LE
      uniFile.WriteUnicodeLine MLSupport.GSS("unity_main", "statMsg4", "Date, Time, System Message")
      uniFile.Flush
    Else
FILE_ERROR:
      uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", fileName, Error$)
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    End If
    
    uniFile.CloseFile
  End If
End Sub

Public Sub setup_olcols()
  Dim total_ol_cols As Integer
  Dim C3ON As Boolean
  Dim C4ON As Boolean
  Dim C5ON As Boolean
  Dim C6ON As Boolean

  ' Check if to hide prediction value column
  fpspread_pred.Col = 2 ' value
  fpspread_pred.ColHidden = unity_main.m_hideValCol

  'if outlier columns not used, hide them
  total_ol_cols = 0
  fpspread_pred.Col = 3 'md
  
  If (frmedmod.chk_md.Value = 1) Then
    fpspread_pred.ColHidden = False
    C3ON = True
    total_ol_cols = total_ol_cols + 1
  Else
    fpspread_pred.ColHidden = True
  End If
  
  fpspread_pred.Col = 4 'rr
  
  If (frmedmod.chk_rr.Value = 1) Then
    fpspread_pred.ColHidden = False
    C4ON = True
    total_ol_cols = total_ol_cols + 1
  Else
    fpspread_pred.ColHidden = True
  End If

  fpspread_pred.Col = 5

  If (frmedmod.chk_value.Value = 1) Then
    fpspread_pred.ColHidden = False
    C5ON = True
    total_ol_cols = total_ol_cols + 1
  Else
    fpspread_pred.ColHidden = True
  End If

  fpspread_pred.Col = 6

  If (frmedmod.chk_nd.Value = 1) Then
    fpspread_pred.ColHidden = False
    C6ON = True
    total_ol_cols = total_ol_cols + 1
  Else
    fpspread_pred.ColHidden = True
  End If

  Select Case (total_ol_cols)
    Case 0 ' none only prop name and pred value
      fpspread_pred.ColWidth(1) = (fpspread_pred.Width / 1.818) 'was 4.5
      fpspread_pred.ColWidth(2) = (fpspread_pred.Width / 3.7037) 'was 5
    
    Case 1
      fpspread_pred.ColWidth(1) = (fpspread_pred.Width / 2.2522) 'was 4.5
      fpspread_pred.ColWidth(2) = (fpspread_pred.Width / 3.636) 'was 5
      
      If (C3ON = True) Then
        fpspread_pred.ColWidth(3) = (fpspread_pred.Width / 10) 'was 14
      End If
      
      If (C4ON = True) Then
        fpspread_pred.ColWidth(4) = (fpspread_pred.Width / 10) 'was 14
      End If
      
      If (C5ON = True) Then
        fpspread_pred.ColWidth(5) = (fpspread_pred.Width / 10) ' was 14
      End If
    
      If (C6ON = True) Then
        fpspread_pred.ColWidth(6) = (fpspread_pred.Width / 10) ' was 14
      End If
    
    Case 2
      fpspread_pred.ColWidth(1) = (fpspread_pred.Width / 2.53) 'was 4.5
      fpspread_pred.ColWidth(2) = (fpspread_pred.Width / 4.444) 'was 5
      
      If (C3ON = True) Then
        fpspread_pred.ColWidth(3) = (fpspread_pred.Width / 10) 'was 14
      End If
      
      If (C4ON = True) Then
        fpspread_pred.ColWidth(4) = (fpspread_pred.Width / 10) 'was 14
      End If
      
      If (C5ON = True) Then
        fpspread_pred.ColWidth(5) = (fpspread_pred.Width / 10) ' was 14
      End If
    
      If (C6ON = True) Then
        fpspread_pred.ColWidth(6) = (fpspread_pred.Width / 10) ' was 14
      End If
      
    Case 3
      fpspread_pred.ColWidth(1) = (fpspread_pred.Width / 2.9) 'was 4.5
      fpspread_pred.ColWidth(2) = (fpspread_pred.Width / 5.7) 'was 5
      fpspread_pred.ColWidth(3) = (fpspread_pred.Width / 10) 'was 14
      fpspread_pred.ColWidth(4) = (fpspread_pred.Width / 10) 'was 14'
      fpspread_pred.ColWidth(5) = (fpspread_pred.Width / 10) ' was 14
      
    Case 4
      fpspread_pred.ColWidth(1) = (fpspread_pred.Width / 3.2) 'was 2.9
      fpspread_pred.ColWidth(2) = (fpspread_pred.Width / 5.7) 'was 5
      fpspread_pred.ColWidth(3) = (fpspread_pred.Width / 10)  'was 14
      fpspread_pred.ColWidth(4) = (fpspread_pred.Width / 10)  'was 14'
      fpspread_pred.ColWidth(5) = (fpspread_pred.Width / 10)  ' was 14
      fpspread_pred.ColWidth(6) = (fpspread_pred.Width / 10)  ' was 14
  End Select
End Sub

Public Sub clearpredtable()
  Dim ii As Integer
  Dim ht As Long
  
  unity_main.fpspread_pred.Row = 1
  unity_main.fpspread_pred.Col = 2
  unity_main.fpspread_pred.Row2 = fpspread_pred.MaxRows
  unity_main.fpspread_pred.Col2 = unity_main.fpspread_pred.MaxCols
  unity_main.fpspread_pred.BlockMode = True
  unity_main.fpspread_pred.Action = ActionClear
  unity_main.fpspread_pred.BackColor = &HFFFFFF 'reset back to white
  unity_main.fpspread_pred.BlockMode = False
  
  ' Resize row height for any new property added
  ht = fpspread_pred.RowHeight(0)
  
  For ii = 1 To fpspread_pred.MaxRows
    fpspread_pred.RowHeight(ii) = ht
  Next ii
End Sub

Public Sub clearpredtablefull()
  Dim ii As Integer
  Dim ht As Long
  
  unity_main.fpspread_pred.Row = 1
  unity_main.fpspread_pred.Col = 1
  unity_main.fpspread_pred.Row2 = fpspread_pred.MaxRows
  unity_main.fpspread_pred.Col2 = unity_main.fpspread_pred.MaxCols
  unity_main.fpspread_pred.BlockMode = True
  unity_main.fpspread_pred.Action = ActionClear
  unity_main.fpspread_pred.BlockMode = False
  unity_main.fpspread_pred.TopRow = 0
  
  ' Resize row height for any new property added
  ht = fpspread_pred.RowHeight(0)
  
  For ii = 1 To fpspread_pred.MaxRows
    fpspread_pred.RowHeight(ii) = ht
  Next ii
End Sub

Public Sub stripcommas()
  Dim str1 As String
  Dim str2 As String
  Dim onechar As String
  Dim lenstr As Integer
  Dim zz As Integer

  str1 = Trim(unity_main.txtsampcomment.Text)
  str2 = ""
  
  For zz = 1 To Len(str1)
    onechar = Mid(str1, zz, 1)
    
    If onechar = "," Then
      onechar = ";"
    End If
    
    str2 = str2 & onechar
  Next zz
  
  unity_main.txtsampcomment.Text = str2
End Sub

Sub kill_loop(ByVal dbgLevel As LOG_DBG_LEVELS, resetCtr As Boolean, reasonMsg As String)
  
  unity_main.tmr_ref.enabled = False
  unity_main.tmr_all.enabled = False
  unity_main.tmr_sample.enabled = False
  unity_main.errorstring = "Auto operations stopped; " & reasonMsg
  unity_main.write_error (dbgLevel)
  cmd_sample.enabled = True
  If cmd_sample.Visible And cmd_sample.enabled Then
             'cmd_sample.SetFocus
  End If
  If (resetCtr = True) Then
    unity_main.repcounter = 0
  End If
  
  If (cmd_start.Visible = True) Then
    unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status8", "System Stopped") & " - " & MLSupport.GSS("OperStatus", "status1", "Press Start")
  Else
    If (unity_main.m_intRefPPTScan = True) Or (unity_main.m_extRefPPTScan = True) Then
      unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status6", "Reference Qualification Required") & " - " & MLSupport.GSS("OperStatus", "status2", "Press Ref")
    Else
      If (unity_main.m_intRefAutoScan = True) Or (unity_main.m_extRefAutoScan = True) Then
        unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status5", "Reference Required") & " - " & MLSupport.GSS("OperStatus", "status2", "Press Ref")
      Else
        If (unity_main.m_remoteRefScan = True) Then
          unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status9", "Reference Requested Remotely") & " - " & MLSupport.GSS("OperStatus", "status2", "Press Ref")
        Else
          unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status8", "System Stopped") & " - " & MLSupport.GSS("OperStatus", "status3", "Press Scan Sample")
        End If
      End If
    End If
  End If
End Sub

Sub restart_loop(ByVal dbgLevel As LOG_DBG_LEVELS, reasonMsg As String)
  
  If (cmd_start.Visible = True) Then
    cmd_sample.enabled = True
    If cmd_sample.Visible And cmd_sample.enabled Then
      'cmd_sample.SetFocus
    End If
    Exit Sub
  End If
  
  unity_main.m_smplAutoScan = False
  unity_main.tmr_all.enabled = True
  unity_main.errorstring = "Auto operations restarted; " & reasonMsg
  unity_main.write_error (dbgLevel)
      cmd_sample.enabled = True
      If cmd_sample.Visible And cmd_sample.enabled Then
             'cmd_sample.SetFocus
      End If
  If (unity_main.m_intRefPPTScan = True) Or (unity_main.m_extRefPPTScan = True) Then
    unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status6", "Reference Qualification Required")
  Else
    If (unity_main.m_intRefAutoScan = True) Or (unity_main.m_extRefAutoScan = True) Then
      unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status5", "Reference Required")
    Else
      If (unity_main.m_remoteRefScan = True) Then
        unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status9", "Reference Requested Remotely")
      Else
        If (unity_main.m_repsAvg > 1) And (unity_main.repcounter <> 0) And (unity_main.repcounter <= unity_main.m_repsAvg) Then
          lbl_opStatus.Caption = MLSupport.GGS_Params("unity_main.statMsg31", "Insert Repack %1 of %2", CStr(unity_main.repcounter), CStr(unity_main.m_repsAvg))
        Else
          If (lblProd1.Caption = "") Then
            lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status4", "Please Select a Product")
          Else
            unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status10", "Ready for New Sample")
            cmd_sample.enabled = True
            If cmd_sample.Visible And cmd_sample.enabled Then
             'cmd_sample.SetFocus
            End If
          End If
        End If
      End If
    End If
  End If
End Sub

Public Sub hide_taskbar()
 Dim rtn As Long
 
 ' Hide the taskbar
 rtn = FindWindow("Shell_traywnd", "")
 Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
End Sub

Public Sub unloadallforms(Optional formtoignore As String = "")
  Dim f As Form
  Dim rtn As Long
  
#If SSRCS Then
  SSRCSClient.CloseSSRCS
#End If
  
  rtn = FindWindow("Shell_traywnd", "") 'get the Window
  Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW) 'show the Taskbar
  
  For Each f In Forms
    If (f.Name <> formtoignore) Then
      Unload f
      Set f = Nothing
    End If
  Next f
End Sub

Public Sub prepscan()
  
  unity_main.stripcommas      ' remove commas from comment field
  unity_main.setup_olcols     ' set # of outlier columns 'xxx should move to loadprod2?
  clear_avg_tbl               ' Clear table used to calculate averages
  unity_main.firstrep = True
  unity_main.repcounter = 1
  frm_dbug.lst_status.Clear   ' clear the debug status list box
End Sub

#If SSTAR Then
Public Function setup_scan() As Boolean
  Dim nn As Integer
  Dim rc As Boolean
  Dim trayNum As Integer
  Dim refTimeout As Long
  
  ' Setup scan mode
  setup_scan_mode
  
  ' Check if performing internal reference or qualification scan
  If (unity_main.m_scanDataType = SDT_INTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODINTREF) Then
    unity_main.m_intRefFileLoadFlg = False
    unity_main.m_extRefFileSetup = False
    unity_main.m_olRefFileSetup = False
  
    ' Configure reference scan data
    MS11ScanCfgData.wvlnIncr = MS11CfgData.wvlnIncr
    MS11ScanCfgData.scanMode = unity_main.m_scanMode
    MS11ScanCfgData.refTimeout = unity_main.m_intRefTimeout
    MS11ScanCfgData.nScans4Ref = unity_main.m_intRefNScans
    MS11ScanCfgData.nScans4Smpl = 1
    MS11ScanCfgData.smpl4PPT = 0
      
    ' Check if performing internal reference qualification scan
    If (unity_main.m_scanDataType = SDT_INTREFPPT) Then
      MS11ScanCfgData.endWvln = unity_main.m_intRefPPTEndWvln
      MS11ScanCfgData.startWvln = unity_main.m_intRefPPTStartWvln
      MS11ScanCfgData.ref4PPT = 0
      
      unity_main.m_intRefPPTFileSetup = False        ' force to reload internal reference qualification file
      
      ' Must perform new product reference afterwards if reference qualification spectrum range different from product
      If ((unity_main.m_intRefPPTStartWvln <> unity_main.m_smplStartWvln) Or (unity_main.m_intRefPPTEndWvln <> unity_main.m_smplEndWvln)) And _
         (unity_main.m_bType = "internal") Then
        unity_main.m_intRefAutoScan = True
      Else
        ' Don't need to do extra reference
        unity_main.m_intRefAutoScan = False
        unity_main.m_intRefManualScan = False
      End If
    Else     ' product reference scan
      MS11ScanCfgData.endWvln = unity_main.m_smplEndWvln
      MS11ScanCfgData.startWvln = unity_main.m_smplStartWvln
      MS11ScanCfgData.ref4PPT = unity_main.m_intRefPPT
    End If
      
    rc = Set_MS11srv_Scan_Cfg()
    
    If (rc = False) Then
      unity_main.m_ansiErrMsg = "Error configuring internal reference scan data"
      unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status31", "Error configuring internal reference scan data")
      GoTo LeaveRtn
    End If
    
    ' Configure internal reference scan tray
    trayNum = 1
    MS11TrayCfgData.velCont = 0
    MS11TrayCfgData.stps4scn = 0
    MS11TrayCfgData.stps4IX = 0
    rc = Set_MS11srv_Tray_Cfg(trayNum)
    
    If (rc = False) Then
      unity_main.m_ansiErrMsg = "Error configuring internal reference scan tray " & trayNum
      unity_main.m_uniErrMsg = MLSupport.GGS_Params("MS11srv.errMsg3", "Error configuring internal reference scan tray %1", CStr(trayNum))
      GoTo LeaveRtn
    End If
    
    ' Load and setup internal reference qualification if required for product internal reference scan
    If (unity_main.m_scanDataType = SDT_PRODINTREF) And (unity_main.m_intRefPPT <> 0) And (unity_main.m_intRefPPTFileSetup = False) Then
      rc = setup_int_ref_ppt(unity_main.m_smplStartWvln, unity_main.m_smplEndWvln)
      unity_main.m_intRefPPTFileSetup = rc
      
      If (rc = False) Then
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status99", "Error loading/setting up internal reference qualification")
      End If
    End If
  Else
    ' Check if performing external reference or qualification scan
    If (unity_main.m_scanDataType = SDT_EXTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODEXTREF) Then
      unity_main.m_olRefFileSetup = False
      
      ' Configure reference scan data
      MS11ScanCfgData.wvlnIncr = MS11CfgData.wvlnIncr
      MS11ScanCfgData.scanMode = unity_main.m_scanMode
      MS11ScanCfgData.nScans4Smpl = 1
      MS11ScanCfgData.smpl4PPT = 0
      
      ' Check if performing external reference qualification scan
      If (unity_main.m_scanDataType = SDT_EXTREFPPT) Then
        MS11ScanCfgData.endWvln = unity_main.m_extRefPPTEndWvln
        MS11ScanCfgData.startWvln = unity_main.m_extRefPPTStartWvln
        MS11ScanCfgData.nScans4Ref = unity_main.m_extRefPPTNScans
        MS11ScanCfgData.ref4PPT = 0
        MS11ScanCfgData.refTimeout = 0

        unity_main.m_extRefPPTFileSetup = False        ' force to reload external reference qualification file
      
        ' Must perform new product reference afterwards if reference qualification file different from product
        If (unity_main.m_extRefPPTFileName <> unity_main.m_extRefFileName) And (unity_main.m_bType = "external") Then
          unity_main.m_extRefAutoScan = True
        Else
          ' Don't need to do extra reference
          unity_main.m_extRefTimerIgnore = True
          unity_main.m_extRefAutoScan = False
          unity_main.m_extRefManualScan = False
        End If
      Else     ' product reference scan
        MS11ScanCfgData.endWvln = unity_main.m_smplEndWvln
        MS11ScanCfgData.startWvln = unity_main.m_smplStartWvln
        MS11ScanCfgData.nScans4Ref = unity_main.m_extRefNScans
        MS11ScanCfgData.ref4PPT = unity_main.m_extRefPPT
        MS11ScanCfgData.refTimeout = 0
      End If
      
      rc = Set_MS11srv_Scan_Cfg()
    
      If (rc = False) Then
        unity_main.m_ansiErrMsg = "Error configuring external reference scan data"
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status88", "Error configuring external reference scan data")
        GoTo LeaveRtn
      End If
    
      ' Check if performing external reference qualification scan
      If (unity_main.m_scanDataType = SDT_EXTREFPPT) Then
        trayNum = unity_main.m_extRefPPTTrayNum
        
        ' Configure reference scan tray based on movement mode
        Select Case (unity_main.m_rotateMoveMode)
          Case TRM_NONE               ' no movement
            MS11TrayCfgData.velCont = MS11TrayInfoData(trayNum).maxVel
            MS11TrayCfgData.stps4scn = MS11TrayInfoData(trayNum).maxStps4scn
            MS11TrayCfgData.stps4IX = MS11TrayInfoData(trayNum).maxStps4IX
          Case TRM_CONT               ' continuous movement
            MS11TrayCfgData.velCont = unity_main.m_extRefPPTRotateSpeed * unity_main.m_extRefPPTRotateDir
            MS11TrayCfgData.stps4scn = MS11TrayInfoData(trayNum).maxStps4scn
            MS11TrayCfgData.stps4IX = MS11TrayInfoData(trayNum).maxStps4IX
          Case TRM_STEP               ' stepped movement
            MS11TrayCfgData.velCont = MS11TrayInfoData(trayNum).maxVel
            MS11TrayCfgData.stps4scn = unity_main.m_extRefPPTRotateStepSteps * unity_main.m_extRefPPTRotateDir
            MS11TrayCfgData.stps4IX = MS11TrayInfoData(trayNum).maxStps4IX
          Case TRM_INDEX              ' indexed movement
            MS11TrayCfgData.velCont = MS11TrayInfoData(trayNum).maxVel
            MS11TrayCfgData.stps4scn = MS11TrayInfoData(trayNum).maxStps4scn
            MS11TrayCfgData.stps4IX = unity_main.m_extRefPPTRotateIndexSteps * unity_main.m_extRefPPTRotateDir
        End Select
      
      Else    ' performing product external reference scan
        trayNum = unity_main.m_extRefTrayNum
      
        ' Configure reference scan tray based on movement mode
        Select Case (unity_main.m_rotateMoveMode)
          Case TRM_NONE               ' no movement
            MS11TrayCfgData.velCont = MS11TrayInfoData(trayNum).maxVel
            MS11TrayCfgData.stps4scn = MS11TrayInfoData(trayNum).maxStps4scn
            MS11TrayCfgData.stps4IX = MS11TrayInfoData(trayNum).maxStps4IX
          Case TRM_CONT               ' continuous movement
            MS11TrayCfgData.velCont = unity_main.m_extRefRotateSpeed * unity_main.m_extRefRotateDir
            MS11TrayCfgData.stps4scn = MS11TrayInfoData(trayNum).maxStps4scn
            MS11TrayCfgData.stps4IX = MS11TrayInfoData(trayNum).maxStps4IX
          Case TRM_STEP               ' stepped movement
            MS11TrayCfgData.stps4scn = unity_main.m_extRefRotateStepSteps * unity_main.m_extRefRotateDir
            MS11TrayCfgData.velCont = MS11TrayInfoData(trayNum).maxVel
            MS11TrayCfgData.stps4IX = MS11TrayInfoData(trayNum).maxStps4IX
          Case TRM_INDEX              ' indexed movement
            MS11TrayCfgData.velCont = MS11TrayInfoData(trayNum).maxVel
            MS11TrayCfgData.stps4scn = MS11TrayInfoData(trayNum).maxStps4scn
            MS11TrayCfgData.stps4IX = unity_main.m_extRefRotateIndexSteps * unity_main.m_extRefRotateDir
        End Select
      End If
        
      ' Configure external reference scan tray
      rc = Set_MS11srv_Tray_Cfg(trayNum)
    
      If (rc = False) Then
        unity_main.m_ansiErrMsg = "Error configuring external reference scan tray " & trayNum
        unity_main.m_uniErrMsg = MLSupport.GGS_Params("MS11srv.errMsg5", "Error configuring external reference scan tray %1", CStr(trayNum))
        GoTo LeaveRtn
      End If
    
      ' Load and setup external reference qualification if required for product external reference scan
      If (unity_main.m_scanDataType = SDT_PRODEXTREF) And (unity_main.m_extRefPPT <> 0) And (unity_main.m_extRefPPTFileSetup = False) Then
        rc = setup_ext_ref_ppt(unity_main.m_extRefPPTFileName)
        unity_main.m_extRefPPTFileSetup = rc
      
        If (rc = False) Then
          unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status100", "Error loading/setting up external reference qualification")
        End If
      End If
    Else
      ' Configure sample scan data
      MS11ScanCfgData.startWvln = unity_main.m_smplStartWvln
      MS11ScanCfgData.endWvln = unity_main.m_smplEndWvln
      MS11ScanCfgData.wvlnIncr = MS11CfgData.wvlnIncr
      MS11ScanCfgData.scanMode = unity_main.m_scanMode
      MS11ScanCfgData.nScans4Ref = 1
      MS11ScanCfgData.nScans4Smpl = unity_main.m_smplNScans
      MS11ScanCfgData.smpl4PPT = unity_main.m_smplPPT
      
      ' Check if not performing internal reference calibration function
      If (unity_main.m_intRefCalFlg = False) Then
        ' Setup reference qualification
        Select Case (unity_main.m_bType)
          Case "internal"
#If SSRCS Then
            SSRCSClient.GetRefTimeout refTimeout
#Else
            refTimeout = MS11srv.refTimeout
#End If

            MS11ScanCfgData.ref4PPT = unity_main.m_intRefPPT
          
            ' Check if internal reference performed on demand
            If (unity_main.m_backFreq = REF_FREQ_ON_DEMAND) Then
              MS11ScanCfgData.refTimeout = unity_main.m_intRefTimeout
            Else
              MS11ScanCfgData.refTimeout = 0
            End If
          Case "external"
            MS11ScanCfgData.ref4PPT = unity_main.m_extRefPPT
            MS11ScanCfgData.refTimeout = 0
          Case "file"
            MS11ScanCfgData.ref4PPT = 0
            MS11ScanCfgData.refTimeout = 0
        End Select
      Else
        MS11ScanCfgData.ref4PPT = 0
        MS11ScanCfgData.refTimeout = 0
      End If
      
      rc = Set_MS11srv_Scan_Cfg()
    
      If (rc = False) Then
        unity_main.m_ansiErrMsg = "Error configuring sample scan data"
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status32", "Error configuring sample scan data")
        GoTo LeaveRtn
      End If
    
      ' Configure sample scan tray based on movement mode
      trayNum = unity_main.m_trayNum

      Select Case (unity_main.m_rotateMoveMode)
        Case TRM_NONE               ' no movement
          MS11TrayCfgData.velCont = MS11TrayInfoData(trayNum).maxVel
          MS11TrayCfgData.stps4scn = MS11TrayInfoData(trayNum).maxStps4scn
          MS11TrayCfgData.stps4IX = MS11TrayInfoData(trayNum).maxStps4IX
        Case TRM_CONT               ' continuous movement
          MS11TrayCfgData.velCont = unity_main.m_rotateSpeed * unity_main.m_rotateDir
          MS11TrayCfgData.stps4scn = MS11TrayInfoData(trayNum).maxStps4scn
          MS11TrayCfgData.stps4IX = MS11TrayInfoData(trayNum).maxStps4IX
        Case TRM_STEP               ' stepped movement
          MS11TrayCfgData.velCont = MS11TrayInfoData(trayNum).maxVel
          MS11TrayCfgData.stps4scn = unity_main.m_rotateStepSteps * unity_main.m_rotateDir
          MS11TrayCfgData.stps4IX = MS11TrayInfoData(trayNum).maxStps4IX
        Case TRM_INDEX              ' indexed movement
          MS11TrayCfgData.velCont = MS11TrayInfoData(trayNum).maxVel
          MS11TrayCfgData.stps4scn = MS11TrayInfoData(trayNum).maxStps4scn
          MS11TrayCfgData.stps4IX = unity_main.m_rotateIndexSteps * unity_main.m_rotateDir
      End Select
    
      rc = Set_MS11srv_Tray_Cfg(trayNum)
    
      If (rc = False) Then
        unity_main.m_ansiErrMsg = "Error configuring sample scan tray " & trayNum
        unity_main.m_uniErrMsg = MLSupport.GGS_Params("MS11srv.errMsg4", "Error configuring sample scan tray %1", CStr(trayNum))
        GoTo LeaveRtn
      End If
      
      ' Check if not performing internal reference calibration function
      If (unity_main.m_intRefCalFlg = False) Then
        ' Setup reference buffer
        Select Case (unity_main.m_bType)
          Case "internal"
            ' Check if to load saved internal reference dued to internal reference calibration
            If (unity_main.m_intRefFileLoadFlg = True) Then
              rc = setup_int_ref(refTimeout)
              
              If (rc = False) Then
                unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status110", "Error loading/setting up internal reference")
                GoTo LeaveRtn
              End If
            End If

          Case "external"
            ' Load and setup external reference if not already
            If (unity_main.m_extRefFileSetup = False) Then
              rc = frm_extRef.load_ext_ref_cfg_file(unity_main.m_extRefFileName, True)
            
              If (rc = True) Then
                ' Check reference wavelengths against instrument
                rc = frm_collect.check_ref_file_wvlns(unity_main.m_extRefStartWvln, unity_main.m_extRefEndWvln, unity_main.m_extRefFileName)
              
                If (rc = True) Then
                  rc = frm_extRef.setup_ext_ref(unity_main.m_extRefFileName)
                  unity_main.m_extRefFileSetup = rc
                End If
              End If
            
              If (rc = False) Then
                unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status101", "Error loading/setting up external reference")
                GoTo LeaveRtn
              End If
            End If
          Case "file"
            unity_main.m_extRefFileSetup = False
          
            ' Load and setup offline reference if not already
            If (unity_main.m_olRefFileSetup = False) Then
              ' Check reference wavelengths against instrument
              rc = frm_olRef.get_ol_ref_wvlns(unity_main.m_olRefFileName)
            
              If (rc = True) Then
                rc = frm_collect.check_ref_file_wvlns(frm_olRef.m_olRefStartWvln, frm_olRef.m_olRefEndWvln, frm_olRef.m_olRefFileName)
              
                If (rc = True) Then
                  rc = frm_olRef.setup_ol_ref(unity_main.m_olRefFileName)
                  unity_main.m_olRefFileSetup = rc
                End If
              End If
            
              If (rc = False) Then
                unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status102", "Error loading/setting up offline file reference")
                GoTo LeaveRtn
              End If
            End If
        End Select
      End If
    End If
  End If
  
LeaveRtn:
  setup_scan = rc
End Function
#End If

#If SSTAR Then
Public Sub setup_scan_mode()
  Dim refType As String

  ' Check if not performing internal reference calibration function
  If (unity_main.m_intRefCalFlg = False) Then
    ' Base scan mode based on instrument settings
    ' scan direction (bits 0-1), subtract dark selection (bit 2),
    ' no light selection (bit 3), and spectral corrections (bits 10-13)
    unity_main.m_scanMode = unity_main.m_sysScanMode
  Else
    ' Check if performing check reference function
    If (m_calibFunc = CF_CHECK) Then
      ' Base scan mode based on instrument settings
      ' scan direction (bits 0-1), subtract dark selection (bit 2),
      ' no light selection (bit 3), and spectral corrections (bits 10-13)
      m_scanMode = MS11CfgData.sysScanMode
    Else   ' performing update calibration function
      ' Modify scan mode bits to remove spectrix operations
      m_scanMode = MS11CfgData.sysScanMode + Not MS11CfgData.spxScnMod
    End If
  End If
        
  ' Setup scan mode by adding product settings
  ' Add tray motion (bits 4-5)
  ' Check if performing internal reference or qualification scan
  If (unity_main.m_scanDataType = SDT_INTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODINTREF) Then
    unity_main.m_scanMode = unity_main.m_scanMode + (TRM_NONE * 16)
  Else
    ' Check if performing external reference or qualification scan
    If (unity_main.m_scanDataType = SDT_EXTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODEXTREF) Then
      unity_main.m_scanMode = unity_main.m_scanMode + (unity_main.m_extRefRotateMoveMode * 16)
    Else
      ' Performing product sample or qualification scan
      unity_main.m_scanMode = unity_main.m_scanMode + (unity_main.m_rotateMoveMode * 16)
    End If
  End If

  ' Add tray position control (bits 6-7)
  ' Check if performing internal reference or qualification scan
  If (unity_main.m_scanDataType = SDT_INTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODINTREF) Then
    unity_main.m_scanMode = unity_main.m_scanMode + &H40        ' internal reference
  Else
    ' Check if performing external reference or qualification scan
    If (unity_main.m_scanDataType = SDT_EXTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODEXTREF) Then
      unity_main.m_scanMode = unity_main.m_scanMode + &H80        ' external sample
    Else
      ' Performing product sample or qualification scan
      unity_main.m_scanMode = unity_main.m_scanMode + &H80        ' external sample
    End If
  End If

  ' Add PPT qualify spectrum (bit 8)
  ' Set bit if internal ref qualification configured
  If ((unity_main.m_scanDataType = SDT_PRODINTREF) And (unity_main.m_intRefPPT <> 0)) Then
    unity_main.m_scanMode = unity_main.m_scanMode + &H100
  Else
    ' Set bit if external ref qualification configured
    If ((unity_main.m_scanDataType = SDT_PRODEXTREF) And (unity_main.m_extRefPPT <> 0)) Then
      unity_main.m_scanMode = unity_main.m_scanMode + &H100
    Else
      ' Set bit if sample qualification configured
      If (unity_main.m_scanDataType = SDT_PRODPPT) And (unity_main.m_smplPPT <> 0) Then
        unity_main.m_scanMode = unity_main.m_scanMode + &H100
      End If
    End If
  End If

  ' Add reference/sample scan (bit 9)
  If (unity_main.m_scanDataType = SDT_PRODPPT) Or (unity_main.m_scanDataType = SDT_PRODSMPL) Then
    unity_main.m_scanMode = unity_main.m_scanMode + &H200
  End If
End Sub
#End If

#If SSTAR Then
Public Function setup_int_ref(refTimeout As Long) As Boolean
  Dim rc As Boolean
  Dim spcFilename As String
  Dim spcIO As GSpcIOLib.GSPCio
  Dim numSubfiles As Long
  Dim errStrg As String
  Dim subFileIndx As Long
  Dim varXVals As Variant
  Dim varYVals As Variant
   
  ' Load last taken reference spectrum file
  spcFilename = (REFERENCES_DIR & REFERENCE_SCAN_FILE)
  rc = LoadSpcFile(spcFilename, spcIO, numSubfiles, errStrg)
  
  If (rc = True) Then
    ' Get reference spectrum data
    subFileIndx = 0
    rc = GetSpcFileSpectrumData(spcIO, subFileIndx, varXVals, varYVals, errStrg)
    
    If (rc = True) Then
      Dim numPts As Integer
      Dim nn As Long
      
      numPts = (spcIO.LastPoint - spcIO.FirstPoint) / MS11CfgData.wvlnIncr
      ReDim ProdRefXVals(numPts)
      ReDim ProdRefYVals(numPts)
      
      For nn = 0 To numPts
        ProdRefXVals(nn) = varXVals(nn)
        ProdRefYVals(nn) = varYVals(nn)
      Next nn
        
      ' clear any previous errors
      Clear_MS11_Error_Codes
        
#If SSRCS Then
      ' Setup reference spectrum data for qualification
      SSRCSClientError = SSRCSClient.SetRefScan((numPts + 1), ProdRefYVals(0))
        
      If (SSRCSClientError <> 0) Then
        rc = False
      End If
#Else
      ' Setup reference spectrum data for qualification
      rc = MS11srv.SetRefScan(ProdRefYVals(0))
#End If
      
      If (rc = False) Then
        Call Get_MS11_Errorcodes_Msg(errStrg)
        unity_main.m_ansiErrMsg = "Error setting internal reference scan data"
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status109", "Error setting internal reference scan data")
      Else
        If (refTimeout = 0) Then
          refTimeout = 1
        End If
        
#If SSRCS Then
        SSRCSClientError = SSRCSClient.SetRefTimeout(refTimeout)
#Else
        MS11srv.refTimeout = refTimeout
#End If
      End If
    Else
      unity_main.m_ansiErrMsg = "Error reading spectrum file: " & spcFilename
      unity_main.m_uniErrMsg = MLSupport.GGS_Params("errMsg5", "Error reading from spectrum file: %1", spcFilename)
    End If
  
    Set spcIO = Nothing
    CloseSpcFile
  Else
    unity_main.m_ansiErrMsg = "Error opening spectrum file: " & spcFilename
    unity_main.m_uniErrMsg = MLSupport.GGS_Params("errMsg4", "Error opening spectrum file: %1", spcFilename)
  End If
 
  ' Force internal reference collection if internal reference file cannot be loaded
  If (rc = False) Then
    unity_main.m_ansiErrMsg = unity_main.m_ansiErrMsg & ". " & errStrg
    unity_main.m_intRefAutoScan = True
  Else
    unity_main.m_intRefFileLoadFlg = False
  End If
  
  setup_int_ref = rc
End Function
#End If

#If SSTAR Then
Public Function setup_ext_ref_ppt(extRefPPTFileName As String) As Boolean
  Dim rc As Boolean
  Dim spcFilename As String
  Dim spcIO As GSpcIOLib.GSPCio
  Dim numSubfiles As Long
  Dim errStrg As String
  Dim subFileIndx As Long
  Dim varXVals As Variant
  Dim varYVals As Variant
   
  ' Load reference qualification file
  spcFilename = (REFERENCES_DIR & extRefPPTFileName & PPT_SCAN_FILE)
  rc = LoadSpcFile(spcFilename, spcIO, numSubfiles, errStrg)
  
  If (rc = True) Then
    ' Get reference spectrum data
    subFileIndx = 0
    rc = GetSpcFileSpectrumData(spcIO, subFileIndx, varXVals, varYVals, errStrg)
    
    If (rc = True) Then
      Dim numPts As Integer
      Dim nn As Integer
      
      numPts = (spcIO.LastPoint - spcIO.FirstPoint) / MS11CfgData.wvlnIncr
      ReDim RefPPTXVals(numPts)
      ReDim RefPPTYVals(numPts)
      
      For nn = 0 To numPts
        RefPPTXVals(nn) = varXVals(nn)
        RefPPTYVals(nn) = varYVals(nn)
      Next nn
        
      ' clear any previous errors
      Clear_MS11_Error_Codes
        
      ' Setup reference spectrum data for qualification
#If SSRCS Then
      SSRCSClientError = SSRCSClient.SetRefRefScan((numPts + 1), RefPPTYVals(0))
      
      If (SSRCSClientError <> 0) Then
        rc = False
#Else
      rc = MS11srv.SetRefRefScan(RefPPTYVals(0))
     
      If (rc = False) Then
#End If
        Call Get_MS11_Errorcodes_Msg(errStrg)
        unity_main.m_ansiErrMsg = "Error setting external reference qualification scan data"
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status93", "Error setting external reference qualification scan data")
      End If
    Else
      unity_main.m_ansiErrMsg = "Error reading spectrum file: " & spcFilename
      unity_main.m_uniErrMsg = MLSupport.GGS_Params("errMsg5", "Error reading from spectrum file: %1", spcFilename)
    End If
  
    Set spcIO = Nothing
    CloseSpcFile
  Else
    unity_main.m_ansiErrMsg = "Error opening spectrum file: " & spcFilename
    unity_main.m_uniErrMsg = MLSupport.GGS_Params("errMsg4", "Error opening spectrum file: %1", spcFilename)
  End If
 
  ' Force reference qualification collection if qualification file cannot be loaded
  If (rc = False) Then
    unity_main.m_ansiErrMsg = unity_main.m_ansiErrMsg & ". " & errStrg
    unity_main.m_extRefPPTScan = True
  End If
  
  setup_ext_ref_ppt = rc
End Function
#End If

#If SSTAR Then
Public Function setup_int_ref_ppt(startWvln As Double, endWvln As Double) As Boolean
  Dim rc As Boolean
  Dim spcFilename As String
  Dim spcIO As GSpcIOLib.GSPCio
  Dim numSubfiles As Long
  Dim errStrg As String
  Dim subFileIndx As Long
  Dim varXVals As Variant
  Dim varYVals As Variant
   
  ' Load reference qualification file
  spcFilename = (REFERENCES_DIR & startWvln & "-" & endWvln & INT_REF_PPT_SCAN_FILE)
  rc = LoadSpcFile(spcFilename, spcIO, numSubfiles, errStrg)
  
  If (rc = True) Then
    ' Get reference spectrum data
    subFileIndx = 0
    rc = GetSpcFileSpectrumData(spcIO, subFileIndx, varXVals, varYVals, errStrg)
    
    If (rc = True) Then
      Dim numPts As Integer
      Dim nn As Integer
      
      numPts = (endWvln - startWvln) / MS11CfgData.wvlnIncr
      ReDim RefPPTXVals(numPts)
      ReDim RefPPTYVals(numPts)
      
      For nn = 0 To numPts
        RefPPTXVals(nn) = varXVals(nn)
        RefPPTYVals(nn) = varYVals(nn)
      Next nn
        
      ' clear any previous errors
      Clear_MS11_Error_Codes
        
      ' Setup reference spectrum data for qualification
#If SSRCS Then
      SSRCSClientError = SSRCSClient.SetRefRefScan((numPts + 1), RefPPTYVals(0))
      
      If (SSRCSClientError <> 0) Then
        rc = False
#Else
      rc = MS11srv.SetRefRefScan(RefPPTYVals(0))
      
      If (rc = False) Then
#End If
        Call Get_MS11_Errorcodes_Msg(errStrg)
        unity_main.m_ansiErrMsg = "Error setting internal reference qualification scan data"
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status33", "Error setting internal reference qualification scan data")
      End If
    Else
      unity_main.m_ansiErrMsg = "Error reading spectrum file: " & spcFilename
      unity_main.m_uniErrMsg = MLSupport.GGS_Params("errMsg5", "Error reading from spectrum file: %1", spcFilename)
    End If
  
    Set spcIO = Nothing
    CloseSpcFile
  Else
    unity_main.m_ansiErrMsg = "Error opening spectrum file: " & spcFilename
    unity_main.m_uniErrMsg = MLSupport.GGS_Params("errMsg4", "Error opening spectrum file: %1", spcFilename)
  End If
 
  ' Force reference qualification collection if qualification file cannot be loaded
  If (rc = False) Then
    unity_main.m_ansiErrMsg = unity_main.m_ansiErrMsg & ". " & errStrg
    unity_main.m_intRefPPTScan = True
  End If
  
  setup_int_ref_ppt = rc
End Function
#End If

#If SSTAR Then
Public Sub setup_ext_ref_pos()
  Dim extRefTrayInfo As String
  Dim adaptIndx As Integer
  Dim multiCupIndx As Integer
  Dim rotateMoveMode As TRAY_ROTATE_MOVEMENTS
  Dim rotateDir As TRAY_ROTATE_DIRS

  frm_status.lbl_statusCmd.ForeColor = RGB(0, 128, 0)  ' dark green
  
  ' Check if external reference qualification scan
  If (unity_main.m_scanDataType = SDT_EXTREFPPT) Then
    adaptIndx = unity_main.m_extRefPPTAdaptIndx
    multiCupIndx = unity_main.m_extRefPPTMultiCupIndx
  Else    ' external reference product scan
    adaptIndx = unity_main.m_extRefAdaptIndx
    multiCupIndx = unity_main.m_extRefMultiCupIndx
  End If
    
  ' Display adapter/tray type used for external reference scan
  extRefTrayInfo = (unity_main.lblSampleModeTitle.Caption & ": ")
    
  If (unity_main.m_adapterType <> CFG_MULTI_CUP_AT) Then
    extRefTrayInfo = extRefTrayInfo & MS11AdapterInfo(adaptIndx).dspName
  Else        ' Display multi-cup type
    extRefTrayInfo = extRefTrayInfo & MS11MultiCupInfo(multiCupIndx).dspName
  End If
   
  ' Check if external reference qualification scan
  If (unity_main.m_scanDataType = SDT_EXTREFPPT) Then
    rotateMoveMode = unity_main.m_extRefPPTRotateMoveMode
    rotateDir = unity_main.m_extRefPPTRotateDir
  Else    ' external reference product scan
    rotateMoveMode = unity_main.m_extRefRotateMoveMode
    rotateDir = unity_main.m_extRefRotateDir
  End If
    
  ' Display platter movement used for external reference scan
  extRefTrayInfo = (extRefTrayInfo & vbCrLf & unity_main.lblMovementTitle.Caption & ": ")
 
  Select Case (rotateMoveMode)
    Case TRM_NONE
      extRefTrayInfo = extRefTrayInfo & frm_collect.opt_rotateModeNone.Caption
    Case TRM_CONT
      extRefTrayInfo = extRefTrayInfo & frm_collect.opt_rotateModeCont.Caption
    Case TRM_STEP
      extRefTrayInfo = extRefTrayInfo & frm_collect.opt_rotateModeStep.Caption
    Case TRM_INDEX
      extRefTrayInfo = extRefTrayInfo & frm_collect.opt_rotateModeIndex.Caption
  End Select
  
  ' Add rotation direction
  If (rotateMoveMode <> TRM_NONE) Then
    If (rotateDir = TRD_CW) Then
      extRefTrayInfo = extRefTrayInfo & " - " & MLSupport.GSS("unity_main", "rotateCW", "CW")
    Else
      extRefTrayInfo = extRefTrayInfo & " - " & MLSupport.GSS("unity_main", "rotateCCW", "CCW")
    End If
  End If

  frm_status.lbl_statusCmd.Caption = extRefTrayInfo & vbCrLf & MLSupport.GSS("OperStatus", "status103", "Position reference and then press Start")

  unity_main.m_extRefPosition = -1
  frm_status.cmd_cancelScan.Visible = True
  frm_status.cmd_startScan.Visible = True
  unity_main.m_scanTmrState = STS_WAIT_POS_EXT_REF
End Sub
#End If

#If SSTAR Then
Public Function start_scan() As Boolean
  Dim rc As Boolean
  Dim StartTime As Single
    
  rc = True
  
  ' Select tray base on scan type
  ' Check if internal reference qualification or external reference scan
  If (unity_main.m_scanDataType = SDT_INTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODINTREF) Then
#If SSRCS Then
    SSRCSClientError = SSRCSClient.SetTraySelection(1)
#Else
    MS11srv.traySelection = 1
#End If
  Else
      ' Check if external reference qualification scan
    If (unity_main.m_scanDataType = SDT_EXTREFPPT) Then
#If SSRCS Then
      SSRCSClientError = SSRCSClient.SetTraySelection(unity_main.m_extRefPPTTrayNum)
#Else
      MS11srv.traySelection = unity_main.m_extRefPPTTrayNum
#End If
    Else
      ' Check if product external reference scan
      If (unity_main.m_scanDataType = SDT_PRODEXTREF) Then
#If SSRCS Then
        SSRCSClientError = SSRCSClient.SetTraySelection(unity_main.m_extRefTrayNum)
#Else
        MS11srv.traySelection = unity_main.m_extRefTrayNum
#End If
      Else     ' product sample scan
#If SSRCS Then
        SSRCSClientError = SSRCSClient.SetTraySelection(unity_main.m_trayNum)
#Else
        MS11srv.traySelection = unity_main.m_trayNum
#End If
      End If
    End If
  End If
  
  ' Check if first product sample or product qualification scan
  If (unity_main.m_scanDataType = SDT_PRODSMPL) Or (unity_main.m_scanDataType = SDT_PRODPPT) And (unity_main.repcounter = 1) Then
#If SSRCS Then
    SSRCSClientError = SSRCSClient.SetReplicates(0)
#Else
    MS11srv.nReplicates = 0      ' clear averaging sum buffer
#End If
  End If
  
  StartTime = Timer
  
  ' Check if tray busy
  Do While (True)

#If SSRCS Then
    SSRCSClientError = SSRCSClient.GetTrayStatus(m_trayStatus)

    If ((m_trayStatus And &H1000) = 0) Then GoTo TRAY_READY

    If (chk_timeout(StartTime, 8) = True) Then
#Else
    If ((unity_main.MS11srv.trayStatus And &H1000) = 0) Then GoTo TRAY_READY
    
    If (chk_timeout(StartTime, 5) = True) Then
#End If
      Exit Do
    End If
    
    DoEvents
  Loop
 
  ' Report error if tray not ready in 5 seconds
  rc = False
  unity_main.m_ansiErrMsg = "Error tray busy wait timeout"
  unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status34", "Error tray busy wait timeout")
  GoTo LeaveRtn
 
TRAY_READY:
  frm_status.scanProgress.percent = 0
#If SSTAR Then
  frm_batchRun.scanProgress.percent = 0
#End If
  
  ' clear any previous errors and event queue
  Clear_MS11_Error_Codes
  clear_GN_eventQ
  
#If SSRCS Then
  SSRCSClientError = SSRCSClient.ScanStart
  
  If (SSRCSClientError = 0) Then
#Else
  rc = MS11srv.ScanStart
    
  If (rc = True) Then
#End If
    frm_status.lbl_statusCmd.ForeColor = RGB(0, 128, 0)  ' dark green
  
    ' Display buttons indicating scan has started
    frm_status.cmd_exitScan.Visible = False
    frm_status.cmd_resumeScan.Visible = False
    frm_status.cmd_retryScan.Visible = False
    frm_status.cmd_abortScan.Visible = True
  Else
#If SSRCS Then
    rc = False
#End If
  
    If (unity_main.m_scanDataType = SDT_INTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODINTREF) Then
      unity_main.m_ansiErrMsg = "Error starting internal reference scan"
      unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status35", "Error starting internal reference scan")
    Else
      If (unity_main.m_scanDataType = SDT_EXTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODEXTREF) Then
        unity_main.m_ansiErrMsg = "Error starting external reference scan"
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status94", "Error starting external reference scan")
      Else
        unity_main.m_ansiErrMsg = "Error starting sample scan"
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status36", "Error starting sample scan")
      End If
    End If
  End If
  
LeaveRtn:
  start_scan = rc
End Function
#End If

#If SSTAR Then
Public Function check_scan_cmpl() As Boolean
  Dim rc As Integer
  Dim sel As Integer
  Dim gnEvent As clsMS11srvGNEvent
  Dim uniMsg As String
  
  rc = True
  
  ' Check if have any GN events to process
  If (m_ms11srvGNEventQ.Count <> 0) Then
    Set gnEvent = m_ms11srvGNEventQ.Item(1)
    m_ms11srvGNEventQ.Remove 1
    
    ' Process event
    Select Case (gnEvent.noticeType)
      Case EVGN_TRAYFLD           ' Tray Failure
        unity_main.m_ansiErrMsg = "Tray failure"
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status37", "Tray failure")
        rc = False
      
      Case EVGN_SCNRFLD           ' Scanner Failure
        unity_main.m_ansiErrMsg = "Scanner failure"
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status38", "Scanner failure")
        rc = False
      
      Case EVGN_SCANABRTD         ' Current Scan has been Aborted
        unity_main.m_ansiErrMsg = "Current scan has been aborted"
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status39", "Current scan has been aborted")
        rc = False
      
      Case EVGN_SCANSTOPD         ' Scanner Stopped Scanning; Resume can be initiated
        If (unity_main.m_scanState = SS_ABORT) Then
          ' Check if performing batch scan
          If (unity_main.m_batchRunFlg = True) Then
            frm_batchRun.batch_scan_aborted
          Else
            ' Check if performing internal reference calibration function
            If (unity_main.m_intRefCalFlg = True) Then
              frm_intRefCalMgmt.scan_aborted
            Else
              frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status40", "Scanner stopped scanning, press Retry to start again")
            End If
          End If
          
          unity_main.m_scanTmrState = STS_ABORT
        Else
          frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status40", "Scanner stopped scanning, press Retry to start again")
          frm_status.cmd_pauseScan.Visible = False
          frm_status.cmd_retryScan.Visible = False
          frm_status.cmd_exitScan.Visible = True
          frm_status.cmd_resumeScan.Visible = True
          frm_status.cmd_abortScan.Visible = True
        End If

        unity_main.m_scanState = SS_STOP

      Case EVGN_IMPOSERR          ' During Scan/Tray Engine Running, Impossible State Transitions occurred
        unity_main.m_ansiErrMsg = "Scan/Tray engine had impossible state transitions occurred"
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status41", "Scan/Tray engine had impossible state transitions occurred")
        rc = False

      Case EVGN_SMPLSCNOK         ' Sample Data Scan Successfully Completed
        uniMsg = MLSupport.GSS("OperStatus", "status42", "Sample data scan successfully completed")
        frm_status.lbl_statusCmd.Caption = uniMsg
        Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Sample data scan successfully completed", uniMsg)
        unity_main.m_scanTmrState = STS_GET_SCAN
      
      Case EVGN_SMPLSCNBAD        ' Sample Data Scan Collected but Bad (Clipped or Failed PPT Qualification)
        ' Check if scan failed dued to clipping
        If (gnEvent.lastErrCode = MS11SRV_LAST_ERRORS.LEC_BADSPECD) Then
          unity_main.m_ansiErrMsg = "Sample data scan collected but spectral values were clipped"
          unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status104", "Sample data scan collected but spectral values were clipped")
        Else
          unity_main.m_ansiErrMsg = "Sample data scan collected but failed PPT qualification"
          unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status43", "Sample data scan collected but failed PPT qualification")
        End If
        
        rc = False
      
      Case EVGN_REFSCNOK          ' Reference Data Scan Successfully Completed
        uniMsg = MLSupport.GSS("OperStatus", "status44", "Internal reference data scan successfully completed")
        frm_status.lbl_statusCmd.Caption = uniMsg
        Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Internal reference data scan successfully completed", uniMsg)
        unity_main.m_scanTmrState = STS_GET_SCAN
      
      Case EVGN_REFSCNBAD         ' Reference Data Scan Collected but Bad (Clipped or Failed PPT Qualification)
        ' Check if scan failed dued to clipping
        If (gnEvent.lastErrCode = MS11SRV_LAST_ERRORS.LEC_BADSPECD) Then
          If (unity_main.m_bType = "internal") Then
            unity_main.m_ansiErrMsg = "Internal reference data scan collected but spectral values were clipped"
            unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status105", "Internal reference data scan collected but spectral values were clipped")
          Else
            unity_main.m_ansiErrMsg = "External reference data scan collected but spectral values were clipped"
            unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status106", "External reference data scan collected but spectral values were clipped")
          End If
        Else
          If (unity_main.m_bType = "internal") Then
            unity_main.m_ansiErrMsg = "Internal reference data scan collected but failed PPT qualification"
            unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status45", "Internal reference data scan collected but failed PPT qualification")
          Else
            unity_main.m_ansiErrMsg = "External reference data scan collected but failed PPT qualification"
            unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status96", "External reference data scan collected but failed PPT qualification")
          End If
        End If
        
        If (unity_main.get_scan_data = True) Then
          If (unity_main.save_scan_data(True) = True) Then
            unity_main.m_scanTimestamp = MLSupport.GSS("OperStatus", "status46", "Failed PPT Qualification")
            unity_main.plot_spectrum
          End If
        End If
        
        rc = False
      
      Case EVGN_OPENDRWR          ' Request to OPEN Sample Drawer
        frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status47", "Please open sample drawer")
        
      Case EVGN_CLOSEDRWR         ' Request to CLOSE Sample Drawer
        frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status48", "Please close sample drawer")
        
      Case EVGN_POSXREFRNC        ' Request to Position External Reference
        uniMsg = MLSupport.GSS("OperStatus", "status49", "User requested to position external reference")
        Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "User requested to position external reference", uniMsg)
        sel = CWrap.ShowMessageBoxW(MLSupport.GSS("OperStatus", "status50", "Please position external reference and then press OK. Press Cancel to abort scan"), vbOKCancel)

        If (sel = vbCancel) Then
          uniMsg = MLSupport.GSS("OperStatus", "status51", "Scan cancel button pressed by user")
          Call unity_main.log_scan_status(LOG_DBG_LEVEL1, "Scan cancel button pressed by user", uniMsg)
          unity_main.m_scanTmrState = STS_ABORT
          
#If SSRCS Then
          SSRCSClientError = SSRCSClient.ScanStop
          
          If (SSRCSClientError = 0) Then
#Else
          If (MS11srv.ScanStop() = True) Then
#End If
            unity_main.m_scanState = SS_STOP
          Else
            ' Report error codes
            uniMsg = MLSupport.GSS("OperStatus", "status28", "Error stopping scan")
            Call frm_status.report_error_codes("Error stopping scan", uniMsg)
            frm_status.cmd_exitScan.Visible = True
          End If
        Else
          ' Tell InfoStar user request has been done
#If SSRCS Then
          SSRCSClientError = SSRCSClient.UserDone
          
          If (SSRCSClientError <> 0) Then
#Else
          If (MS11srv.UserDone() = False) Then
#End If
            unity_main.m_ansiErrMsg = "Error confirming user request"
            unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status52", "Error confirming user request")
            rc = False
          End If
        End If
        
      Case EVGN_POSXSAMPLE        ' Request to Position External Sample
        uniMsg = MLSupport.GSS("OperStatus", "status53", "User requested to position external sample")
        Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "User requested to position external sample", uniMsg)
        sel = CWrap.ShowMessageBoxW(MLSupport.GSS("OperStatus", "status79", "Please position external sample and then press OK. Press Cancel to abort scan"), vbOKCancel)

        If (sel = vbCancel) Then
          uniMsg = MLSupport.GSS("OperStatus", "status51", "Scan cancel button pressed by user")
          Call unity_main.log_scan_status(LOG_DBG_LEVEL1, "Scan cancel button pressed by user", uniMsg)
          unity_main.m_scanTmrState = STS_ABORT
          
#If SSRCS Then
          SSRCSClientError = SSRCSClient.ScanStop
          
          If (SSRCSClientError = 0) Then
#Else
          If (MS11srv.ScanStop() = True) Then
#End If
            unity_main.m_scanState = SS_STOP
          Else
            ' Report error codes
            uniMsg = MLSupport.GSS("OperStatus", "status28", "Error stopping scan")
            Call frm_status.report_error_codes("Error stopping scan", uniMsg)
            frm_status.cmd_exitScan.Visible = True
          End If
        Else
          ' Tell InfoStar user request has been done
#If SSRCS Then
          SSRCSClientError = SSRCSClient.UserDone
          
          If (SSRCSClientError <> 0) Then
#Else
          If (MS11srv.UserDone() = False) Then
#End If
            unity_main.m_ansiErrMsg = "Error confirming user request"
            unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status52", "Error confirming user request")
            rc = False
          End If
        End If
      
      Case EVGN_MSCOMFLD          ' MS1100if COM Failure: Possible Hard Reset or Similar Event
        unity_main.m_ansiErrMsg = "Communication failure with SpectaStar; Need to perform hardware reset"
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status54", "Communication failure with SpectaStar; Need to perform hardware reset")
        rc = False
    
      Case EVGN_MSSRVDIED         ' MS11srv OCX ME_CODE_DEAD State entered
        unity_main.m_ansiErrMsg = "Operational errors within SpectaStar; Need to reboot system"
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status55", "Operational errors within SpectaStar; Need to reboot system")
        rc = False
    End Select
  End If
  
  check_scan_cmpl = rc
End Function
#End If

#If SSTAR Then
Public Function get_scan_data() As Boolean
#If SSRCS Then
  Dim numPts As Long
#Else
  Dim numPts As Integer
#End If
  Dim rc As Boolean
  Dim nn As Integer
  Dim startWvln As Double
  
  rc = True
  
  ' clear any previous errors
  Clear_MS11_Error_Codes
  
#If SSRCS Then
  SSRCSClientError = SSRCSClient.GetScanPts(numPts)
#Else
  numPts = MS11srv.nScanPts - 1
#End If
  numPts = numPts - 1
  
  ' Check if internal reference scan
  If (unity_main.m_scanDataType = SDT_INTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODINTREF) Then
    ReDim ProdRefXVals(numPts)
    ReDim ProdRefYVals(numPts)

    ' Check if completed internal reference qualification scan
    If (unity_main.m_scanDataType = SDT_INTREFPPT) Then
      startWvln = unity_main.m_intRefPPTStartWvln
    Else
      startWvln = unity_main.m_smplStartWvln
    End If

    ' Setup x-axis wavelengths for plotting
    For nn = 0 To numPts
      ProdRefXVals(nn) = startWvln + (MS11CfgData.wvlnIncr * nn)
    Next nn

#If SSRCS Then
    SSRCSClientError = SSRCSClient.GetRefScan((numPts + 1), ProdRefYVals(0))
    
    If (SSRCSClientError <> 0) Then
      rc = False
#Else
    rc = MS11srv.GetRefScan(ProdRefYVals(0))

    If (rc = False) Then
#End If
      unity_main.m_ansiErrMsg = "Error getting internal reference scan data"
      unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status56", "Error getting internal reference scan data")
    End If
  Else
    ' Check if external reference scan
    If (unity_main.m_scanDataType = SDT_EXTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODEXTREF) Then
      ReDim ProdRefXVals(numPts)
      ReDim ProdRefYVals(numPts)

      ' Check if completed external reference qualification scan
      If (unity_main.m_scanDataType = SDT_EXTREFPPT) Then
        startWvln = unity_main.m_extRefPPTStartWvln
      Else
        startWvln = unity_main.m_smplStartWvln
      End If

      ' Setup x-axis wavelengths for plotting
      For nn = 0 To numPts
        ProdRefXVals(nn) = startWvln + (MS11CfgData.wvlnIncr * nn)
      Next nn

#If SSRCS Then
      SSRCSClientError = SSRCSClient.GetRefScan((numPts + 1), ProdRefYVals(0))
      
      If (SSRCSClientError <> 0) Then
        rc = False
#Else
      rc = MS11srv.GetRefScan(ProdRefYVals(0))

      If (rc = False) Then
#End If
        unity_main.m_ansiErrMsg = "Error getting external reference scan data"
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status56", "Error getting external reference scan data")
      End If
    Else      ' get sample scan
      ReDim ProdSmplXVals(numPts)
      ReDim ProdSmplYVals(numPts)

      ' Check if completed internal reference qualification scan
      If (unity_main.m_scanDataType = SDT_PRODINTREF) Then
        startWvln = unity_main.m_smplPPTStartWvln
      Else
        startWvln = unity_main.m_smplStartWvln
      End If

      ' Setup x-axis wavelengths for plotting
      For nn = 0 To numPts
        ProdSmplXVals(nn) = startWvln + (MS11CfgData.wvlnIncr * nn)
      Next nn

      ' Get sample scan data
#If SSRCS Then
      SSRCSClientError = SSRCSClient.GetSmplScan((numPts + 1), ProdSmplYVals(0))
      
      If (SSRCSClientError <> 0) Then
        rc = False
#Else
      rc = MS11srv.GetSmplScan(ProdSmplYVals(0))
      
      If (rc = False) Then
#End If
        unity_main.m_ansiErrMsg = "Error getting sample scan data"
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status57", "Error getting sample scan data")
      End If
    End If
  End If
  
  get_scan_data = rc
End Function
#End If

#If SSTAR Then
Public Function smooth_scan_data() As Boolean
  Dim uniMsg As String
  Dim startIndx As Long
  Dim endIndx As Long
  Dim numPts As Long
  Dim mathObject As New clsMathComponent
  Dim rc As Long
  
  smooth_scan_data = False
  On Error GoTo OBJECT_ERROR
  
  uniMsg = MLSupport.GGS_Params("unity_main.statMsg42", "Smoothing product sample scan %1 of %2", CStr(unity_main.repcounter), CStr(unity_main.m_smplRepacks))
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Smoothing product sample scan " & unity_main.repcounter & " of " & unity_main.m_smplRepacks), uniMsg)
  
  ' Check if .NET Framework 2.0 is not installed
  If (unity_main.m_netFWInstalled = False) Then
    unity_main.m_ansiErrMsg = "Problem with trying to smooth spectra data without MS .NET Framework 2.0 installed"
    unity_main.m_uniErrMsg = MLSupport.GSS("unity_main", "errMsg9", "Problem with trying to smooth spectra data without MS .NET Framework 2.0 installed")
    Exit Function
  End If
  
  ReDim ProdTreatSmplYVals(UBound(ProdSmplYVals))
  
  ' Check if to use regular smooth algorithm
  If (m_useProgSmooth = False) Then
#If SSRCS Then
    SSRCSClientError = SSRCSClient.GetScanPts(numPts)
    rc = mathObject.makeSmooth_2(m_smoothTypeEnum, ProdSmplYVals, ProdTreatSmplYVals, 0, 0, numPts, CLng(m_smoothNumPts))
#Else
    rc = mathObject.makeSmooth_2(m_smoothTypeEnum, ProdSmplYVals, ProdTreatSmplYVals, 0, 0, MS11srv.nScanPts, CLng(m_smoothNumPts))
#End If
  Else      ' use progressive smooth algorithm
    ' Calculate starting progressive smooth wavelength index
    startIndx = CLng(m_startSmoothWvln - m_smplStartWvln)
    
    If (startIndx < 0) Then
      startIndx = 0
    End If
    
    ' Calculate ending progressive smooth wavelength index
    numPts = (m_smplEndWvln - m_smplStartWvln) / MS11CfgData.wvlnIncr
    endIndx = numPts - (m_smplEndWvln - m_endSmoothWvln)
    
    If (endIndx > numPts) Then
      endIndx = numPts
    End If
    
#If SSRCS Then
    SSRCSClientError = SSRCSClient.GetScanPts(numPts)
    rc = mathObject.makeSmooth_3(m_smoothTypeEnum, ProdSmplYVals, ProdTreatSmplYVals, 0, 0, numPts, CLng(m_startSmoothNumPts), CLng(m_endSmoothNumPts), CLng(m_progSmoothRate), startIndx, endIndx)
#Else
    rc = mathObject.makeSmooth_3(m_smoothTypeEnum, ProdSmplYVals, ProdTreatSmplYVals, 0, 0, MS11srv.nScanPts, CLng(m_startSmoothNumPts), CLng(m_endSmoothNumPts), CLng(m_progSmoothRate), startIndx, endIndx)
#End If
  End If
  
  ' Check if no calculation error
  If (rc = 0) Then
    smooth_scan_data = True
  Else
    unity_main.m_ansiErrMsg = "Sample scan data smoothing error " & rc
    unity_main.m_uniErrMsg = MLSupport.GGS_Params("unity_main.errMsg8", "Sample scan data smoothing error: %1", CStr(rc))
  End If
  
  Exit Function
  
OBJECT_ERROR:
  unity_main.m_ansiErrMsg = "Unity MathComponent.dll component not installed or registered"
  unity_main.m_uniErrMsg = MLSupport.GGS_Params("errMsg7", "Unity %1 component not installed or registered", "MathComponent.dll")
End Function
#End If

#If SSTAR Then
Public Function calc_prod_abs() As Boolean
  Dim numPts As Integer
  Dim mathObject As New clsMathComponent
  Dim rc As Long
  
  calc_prod_abs = False
  
  ' Check if .NET Framework 2.0 is not installed
  If (unity_main.m_netFWInstalled = False) Then
    unity_main.m_ansiErrMsg = "Problem with trying to calculate spectra data absorbance without MS .NET Framework 2.0 installed"
    unity_main.m_uniErrMsg = MLSupport.GSS("unity_main", "errMsg10", "Problem with trying to calculate spectra data absorbance without MS .NET Framework 2.0 installed")
    Exit Function
  End If
  
  On Error GoTo OBJECT_ERROR
  numPts = UBound(ProdSmplYVals)
  ReDim ProdAbsYVals(numPts)
  
  ' Absorbance = Base-10 logarithm of Reflectance
  rc = mathObject.fractionTtoAbsorbance_2(ProdSmplYVals, ProdAbsYVals, 0, 0, numPts + 1, SS_MIN_REFLECT_LIMIT, SS_MAX_REFLECT_LIMIT)
  
  If (rc = 0) Then
    ' Check if spectrum data treated
    If (unity_main.m_enableTreatment = True) Then
      ReDim ProdTreatAbsYVals(numPts)
  
      ' Absorbance = Base-10 logarithm of Reflectance
      rc = mathObject.fractionTtoAbsorbance_2(ProdTreatSmplYVals, ProdTreatAbsYVals, 0, 0, numPts + 1, SS_MIN_REFLECT_LIMIT, SS_MAX_REFLECT_LIMIT)
    End If
  End If
  
  If (rc = 0) Then
    calc_prod_abs = True
  Else
    unity_main.m_ansiErrMsg = "Sample scan data absorbance calculation e error: " & rc
    unity_main.m_uniErrMsg = MLSupport.GGS_Params("unity_main.errMsg9", "Sample scan data absorbance calculation error: %1", CStr(rc))
  End If
  
  Exit Function
  
OBJECT_ERROR:
  unity_main.m_ansiErrMsg = "Unity MathComponent.dll component not installed or registered"
  unity_main.m_uniErrMsg = MLSupport.GGS_Params("errMsg7", "Unity %1 component not installed or registered", "MathComponent.dll")
End Function
#End If

Public Sub sum_prod_avg_abs()
  Dim numPts As Integer
  Dim nn As Integer
  
  numPts = UBound(ProdAbsYVals)
  
  ' Sum absorbance to average later
  For nn = 0 To numPts
    ProdAvgAbsYVals(nn) = ProdAvgAbsYVals(nn) + ProdAbsYVals(nn)
  Next nn
  
  ' Check if spectrum data treated
  If (unity_main.m_enableTreatment = True) Then
    ' Sum treated absorbance to average later
    For nn = 0 To numPts
      ProdTreatAvgAbsYVals(nn) = ProdTreatAvgAbsYVals(nn) + ProdTreatAbsYVals(nn)
    Next nn
  End If
End Sub

Public Sub calc_prod_avg_abs()
  Dim numPts As Integer
  Dim nn As Integer
  Dim uniMsg As String
  
  uniMsg = MLSupport.GSS("OperStatus", "status60", "Performing repack spectra averaging")
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Performing repack spectra averaging", uniMsg)
  numPts = UBound(ProdAvgAbsYVals)
  
  ' Calculate absorbance average
  For nn = 0 To numPts
    ProdAvgAbsYVals(nn) = ProdAvgAbsYVals(nn) / unity_main.m_smplRepacks
  Next nn
  
  ' Check if spectrum data treated
  If (unity_main.m_enableTreatment = True) Then
    ' Calculate absorbance average for treated spectrum
    For nn = 0 To numPts
      ProdTreatAvgAbsYVals(nn) = ProdTreatAvgAbsYVals(nn) / unity_main.m_smplRepacks
    Next nn
  End If
End Sub

#If SSTAR Then
Public Function standardize_spectrum(spectrum() As Double) As Boolean
  Dim rc As Long
  Dim numPts As Integer
  Dim preStdSpectrum() As Double
  Dim errMsg As String
  Dim uniMsg As String
  
  standardize_spectrum = False
  
  ' Check if have transfer stf file to perform standarization
  If (unity_main.m_stfFileName <> "") Then
    On Error GoTo OBJECT_ERROR
    
    ' Check if spectrum different
    If (unity_main.m_svfWaveCnvtFlg = True) Then
      numPts = (m_svfEndWvln - m_svfStartWvln) / m_svfWaveInc
      ReDim preStdSpectrum(numPts)
      rc = PRDObject.cnvtSpectrum(m_smplStartWvln, m_smplEndWvln, MS11CfgData.wvlnIncr, spectrum, m_svfStartWvln, m_svfEndWvln, m_svfWaveInc, preStdSpectrum)
      
      ' Check if any error converting spectrum
      If (rc <> 0) Then
        unity_main.m_ansiErrMsg = "PRDComponent cnvtSpectrum() error: " & rc
        unity_main.m_uniErrMsg = MLSupport.GGS_Params("errMsg8", "%1 %2 error: %3", "PRDComponent", "applyStandardization_3()", CStr(rc))
        Exit Function
      End If
        
      rc = PRDObject.applyStandardization_3(unity_main.m_stfFileName, preStdSpectrum, SVFAbsYVals)
    Else
      rc = PRDObject.applyStandardization_3(unity_main.m_stfFileName, spectrum, SVFAbsYVals)
    End If
    
    ' Check if any error standardizing spectrum
    If (rc <> 0) Then
      unity_main.m_ansiErrMsg = unity_main.m_stfFileName & " applyStandardization_3() error: " & rc
      unity_main.m_uniErrMsg = MLSupport.GGS_Params("errMsg8", "%1 %2 error: %3", unity_main.m_stfFileName, "applyStandardization_3()", CStr(rc))
      Exit Function
    End If
  Else
    SVFAbsYVals = spectrum
  End If
  
  standardize_spectrum = True
  Exit Function
  
OBJECT_ERROR:
  unity_main.m_ansiErrMsg = "Unity PRDComponent.dll component not installed or registered"
  unity_main.m_uniErrMsg = MLSupport.GGS_Params("errMsg7", "Unity %1 component not installed or registered", "PRDComponent.dll")
End Function
#End If

Public Function save_scan_data(Optional ByVal failPPTQual As Boolean) As Boolean
  Dim spcDirName As String
  Dim spcFilename As String
  Dim Comment As String
  Dim scanDataType As SCAN_DATA_TYPES
  Dim rc As Boolean
  Dim noSave As Boolean
  Dim spectrumType As SpectrumTypes
  Dim failPPT As Boolean
  Dim uniFile As New clsUniFile
  Dim errStrg As String
  Dim uniMsg As String
   
  If (IsMissing(failPPTQual)) Then
    failPPT = False
  Else
    failPPT = failPPTQual
  End If

  spectrumType = 0
  CreatePath unity_main.m_saveDir
  
  On Error GoTo OBJECT_ERROR

  ' Check if internal reference scan
  If (unity_main.m_scanDataType = SDT_INTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODINTREF) Then
    ' Check if completed internal reference qualification scan
    If (unity_main.m_scanDataType = SDT_INTREFPPT) Then
      spcDirName = REFERENCES_DIR
      spcFilename = (unity_main.m_intRefPPTStartWvln & "-" & unity_main.m_intRefPPTEndWvln & INT_REF_PPT_SCAN_FILE)
      Comment = MLSupport.GGS_Params("unity_main.statMsg1", "Internal reference qualification scan; wavelengths of %1 to %2", CStr(unity_main.m_intRefPPTStartWvln), CStr(unity_main.m_intRefPPTEndWvln))
      uniMsg = MLSupport.GGS_Params("unity_main.statMsg2", "Saving internal reference qualification scan file: %1", spcDirName & spcFilename & SPC_FILE_EXT)
      Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Saving internal reference qualification scan file: " & spcDirName & spcFilename & SPC_FILE_EXT), uniMsg)
      rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdRefXVals, ProdRefYVals, _
                                   GSpcIOLib.spcYType.spcYIntens, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
                                   
      If (rc = True) Then
        rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
      End If
                                   
      ' Check if reference qualification is same wavelength range as current product
      If (rc = True) And (unity_main.m_intRefPPTStartWvln = unity_main.m_smplStartWvln) And (unity_main.m_intRefPPTEndWvln = unity_main.m_smplEndWvln) Then
        GoTo PRODUCT_INT_REF    ' save scan data as latest product reference
      End If
    Else          ' product reference scan
PRODUCT_INT_REF:
      spectrumType = SpectrumTypes.RefSpectrumType
      scanDataType = SDT_PRODINTREF     ' change type in case was internal reference qualification scan
        
      ' Check if reference failed PPT qualifcation
      If (failPPT = True) Then
        spcDirName = REFERENCES_DIR
        spcFilename = FAIL_PPT_REF_SCAN_FILE
        Comment = MLSupport.GGS_Params("unity_main.statMsg3", "%1 product internal reference scan (failed PPT qualification)", lblProd1.Caption)
        uniMsg = MLSupport.GGS_Params("unity_main.statMsg4", "Saving failed PPT qualification product internal reference scan file: %1", spcDirName & spcFilename & SPC_FILE_EXT)
        Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Saving failed PPT qualification product internal reference scan file: " & spcDirName & spcFilename & SPC_FILE_EXT), uniMsg)
        rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdRefXVals, ProdRefYVals, _
                                     GSpcIOLib.spcYType.spcYIntens, unity_main.m_sysSerialNum, Comment, m_instModel, "", scanDataType, 0, 0, errStrg)
                                   
        If (rc = True) Then
          rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
        End If
      Else    ' passed PPT qualification
        spcDirName = unity_main.m_saveDir
        spcFilename = REFERENCE_SCAN_FILE
        Comment = MLSupport.GGS_Params("unity_main.statMsg5", "%1 product internal reference scan", lblProd1.Caption)
        uniMsg = MLSupport.GGS_Params("unity_main.statMsg6", "Saving product internal reference scan file: %1", spcDirName & spcFilename & SPC_FILE_EXT)
        Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Saving product internal reference scan file: " & spcDirName & spcFilename & SPC_FILE_EXT), uniMsg)
      End If
      
      rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdRefXVals, ProdRefYVals, _
                                   GSpcIOLib.spcYType.spcYIntens, unity_main.m_sysSerialNum, Comment, m_instModel, "", scanDataType, 0, 0, errStrg)
                                     
      If (rc = True) Then
        rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
      End If

      ' Save copy of good reference spectrum in product folder
      If (failPPT = False) Then
        On Error Resume Next
        uniFile.st_CopyFile (spcDirName & spcFilename & SPC_FILE_EXT), (REFERENCES_DIR & spcFilename & SPC_FILE_EXT)
        uniFile.st_CopyFile (spcDirName & spcFilename & SPC_INFO_FILE_EXT), (REFERENCES_DIR & spcFilename & SPC_INFO_FILE_EXT)
        unity_main.backdate = FileDateTime(spcDirName & spcFilename & SPC_FILE_EXT)
      End If
    End If
  Else
    ' Check if external reference scan
    If (unity_main.m_scanDataType = SDT_EXTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODEXTREF) Then
      ' Check if completed external reference qualification scan
      If (unity_main.m_scanDataType = SDT_EXTREFPPT) Then
        spcDirName = REFERENCES_DIR
        spcFilename = (unity_main.m_extRefPPTFileName & PPT_SCAN_FILE)
        Comment = MLSupport.GGS_Params("unity_main.statMsg36", "External reference qualification scan; wavelengths of %1 to %2", CStr(unity_main.m_extRefPPTStartWvln), CStr(unity_main.m_extRefPPTEndWvln))
        uniMsg = MLSupport.GGS_Params("unity_main.statMsg37", "Saving external reference qualification scan file: %1", spcDirName & spcFilename & SPC_FILE_EXT)
        Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Saving external reference qualification scan file: " & spcDirName & spcFilename & SPC_FILE_EXT), uniMsg)
        rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdRefXVals, ProdRefYVals, _
                                     GSpcIOLib.spcYType.spcYIntens, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
                                   
        If (rc = True) Then
          rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
        End If
                                   
        ' Check if reference qualification file is same as current product
        If (rc = True) And (unity_main.m_extRefPPTFileName = unity_main.m_extRefFileName) Then
          GoTo PRODUCT_EXT_REF    ' save scan data as latest product reference
        End If
      Else          ' product reference scan
PRODUCT_EXT_REF:
        spectrumType = SpectrumTypes.RefSpectrumType
        scanDataType = SDT_PRODEXTREF     ' change type in case was external reference qualification scan
        
        ' Check if reference failed PPT qualifcation
        If (failPPT = True) Then
          spcDirName = REFERENCES_DIR
          spcFilename = FAIL_PPT_REF_SCAN_FILE
          Comment = MLSupport.GGS_Params("unity_main.statMsg38", "%1 product external reference scan (failed PPT qualification)", lblProd1.Caption)
          uniMsg = MLSupport.GGS_Params("unity_main.statMsg39", "Saving failed PPT qualification product external reference scan file: %1", spcDirName & spcFilename & SPC_FILE_EXT)
          Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Saving failed PPT qualification product external reference scan file: " & spcDirName & spcFilename & SPC_FILE_EXT), uniMsg)
          rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdRefXVals, ProdRefYVals, _
                                       GSpcIOLib.spcYType.spcYIntens, unity_main.m_sysSerialNum, Comment, m_instModel, "", scanDataType, 0, 0, errStrg)
                                   
          If (rc = True) Then
            rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
          End If
        Else    ' passed PPT qualification
          spcDirName = REFERENCES_DIR
          spcFilename = unity_main.m_extRefFileName
          Comment = MLSupport.GGS_Params("unity_main.statMsg40", "%1 product external reference scan", lblProd1.Caption)
          uniMsg = MLSupport.GGS_Params("unity_main.statMsg41", "Saving product external reference scan file: %1", spcDirName & spcFilename & SPC_FILE_EXT)
          Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Saving product external reference scan file: " & spcDirName & spcFilename & SPC_FILE_EXT), uniMsg)
        End If
      
        rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdRefXVals, ProdRefYVals, _
                                     GSpcIOLib.spcYType.spcYIntens, unity_main.m_sysSerialNum, Comment, m_instModel, "", scanDataType, 0, 0, errStrg)
                                     
        If (rc = True) Then
          rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
        End If

        ' Save copy of good reference spectrum to product folder
        If (failPPT = False) Then
          On Error Resume Next
          uniFile.st_CopyFile (spcDirName & spcFilename & SPC_FILE_EXT), (spcDirName & REFERENCE_SCAN_FILE & SPC_FILE_EXT)
          uniFile.st_CopyFile (spcDirName & spcFilename & SPC_INFO_FILE_EXT), (spcDirName & REFERENCE_SCAN_FILE & SPC_INFO_FILE_EXT)
          uniFile.st_CopyFile (spcDirName & spcFilename & SPC_FILE_EXT), (unity_main.m_saveDir & REFERENCE_SCAN_FILE & SPC_FILE_EXT)
          uniFile.st_CopyFile (spcDirName & spcFilename & SPC_INFO_FILE_EXT), (unity_main.m_saveDir & REFERENCE_SCAN_FILE & SPC_INFO_FILE_EXT)
          unity_main.backdate = FileDateTime(spcDirName & spcFilename & SPC_FILE_EXT)
          spcFilename = REFERENCE_SCAN_FILE
        End If
      End If
    Else
      'Determine which product scan to save
      Select Case (unity_main.m_scanDataType)
        Case SDT_PRODPPT             ' product qualification (reflectance) scan data type
          spcDirName = REFERENCES_DIR
          spcFilename = (PROD_REFLECT_SCAN_FILE)
          Comment = MLSupport.GGS_Params("unity_main.statMsg7", "%1 product qualification scan", lblProd1.Caption)
          uniMsg = MLSupport.GGS_Params("unity_main.statMsg8", "Saving product qualification scan file: %1", spcDirName & spcFilename & SPC_FILE_EXT)
          Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Saving product qualification scan file: " & spcDirName & spcFilename & SPC_FILE_EXT), uniMsg)
          rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdSmplXVals, ProdAbsYVals, _
                                       GSpcIOLib.spcYType.spcYAbsrb, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
                                   
          If (rc = True) Then
            rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
          End If
                                     
        Case SDT_PRODPPTABS          ' product calculated absorbance of qualification scan data type
          spcDirName = REFERENCES_DIR
          spcFilename = unity_main.m_smplPPTFileName
          Comment = MLSupport.GGS_Params("unity_main.statMsg9", "%1 product calculated qualification absorbance", lblProd1.Caption)
          uniMsg = MLSupport.GGS_Params("unity_main.statMsg10", "Saving calculated product qualification absorbance file: %1", spcDirName & spcFilename & SPC_FILE_EXT)
          Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Saving calculated product qualification absorbance file: " & spcDirName & spcFilename & SPC_FILE_EXT), uniMsg)
          rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdSmplXVals, ProdAbsYVals, _
                                       GSpcIOLib.spcYType.spcYAbsrb, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
                                   
          If (rc = True) Then
            rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
          End If

        Case SDT_PRODSMPL            ' product sample (reflectance) scan data type
          spcDirName = unity_main.m_saveDir
          spcFilename = PROD_REFLECT_SCAN_FILE
          Comment = MLSupport.GGS_Params("unity_main.statMsg11", "%1 product sample scan", lblProd1.Caption)
          uniMsg = MLSupport.GGS_Params("unity_main.statMsg12", "Saving product sample scan file: %1", spcDirName & spcFilename & SPC_FILE_EXT)
      
          ' Check if spectrum data treated
          If (unity_main.m_enableTreatment = True) Then
            Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Saving product sample treated scan file: " & spcDirName & spcFilename & SPC_FILE_EXT), uniMsg)
            rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdSmplXVals, ProdTreatSmplYVals, _
                                         GSpcIOLib.spcYType.spcYReflec, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
                      
            If (rc = True) Then
              rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
            End If
                      
            ' Check if to save raw spectrum
            If (unity_main.m_saveSpectra = "Both") Then
              spcFilename = spcFilename & RAW_SCAN_APPEND
              Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Saving product sample raw scan file: " & spcDirName & spcFilename & SPC_FILE_EXT), uniMsg)
              rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdSmplXVals, ProdSmplYVals, _
                                           GSpcIOLib.spcYType.spcYReflec, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
                                           
              If (rc = True) Then
                rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
              End If
            End If
          Else
            Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Saving product sample scan file: " & spcDirName & spcFilename & SPC_FILE_EXT), uniMsg)
            rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdSmplXVals, ProdSmplYVals, _
                                         GSpcIOLib.spcYType.spcYReflec, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
                                   
            If (rc = True) Then
              rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, 0, 0, errStrg)
            End If
          End If
                                     
        Case SDT_PRODSMPLABS         ' product calculated absorbance of sample scan data type
          spectrumType = SpectrumTypes.ProductSpectrumType
          spcDirName = unity_main.m_saveDir
          spcFilename = unity_main.txtsamplename.Text
          Comment = unity_main.txtsampcomment.Text
        
          ' Check if doing replicate scans
          If (unity_main.m_smplRepacks > 1) Then
            Comment = MLSupport.GGS_Params("unity_main.statMsg13", "%1 (repack %2 of %3)", Comment, CStr(unity_main.repcounter), CStr(unity_main.m_smplRepacks))
           
            If (unity_main.m_saveReps = True) Then
              ' Add scan number to file name
              spectrumType = SpectrumTypes.RepackSpectrumType
              spcFilename = (unity_main.txtsamplename.Text & "-" & unity_main.repcounter)
            Else
              spectrumType = 0
              GoTo SAVE_LAST_SCAN
            End If
          End If
        
          If (unity_main.m_saveIt = "save") Then
            uniMsg = MLSupport.GGS_Params("unity_main.statMsg14", "Saving calculated product absorbance file: %1", spcDirName & spcFilename & SPC_FILE_EXT)
            
            ' Check if spectrum data treated
            If (unity_main.m_enableTreatment = True) Then
              Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Saving calculated product treated absorbance file: " & spcDirName & spcFilename & SPC_FILE_EXT), uniMsg)
  
              rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdSmplXVals, ProdTreatAbsYVals, _
                                           GSpcIOLib.spcYType.spcYAbsrb, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
              
              If (rc = True) Then
                rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
              End If
                                         
              ' Check if to save raw spectrum
              If (unity_main.m_saveSpectra = "Both") Then
                spcFilename = spcFilename & RAW_SCAN_APPEND
                uniMsg = MLSupport.GGS_Params("unity_main.statMsg14", "Saving calculated product absorbance file: %1", spcDirName & spcFilename & SPC_FILE_EXT)
                Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Saving calculated product raw absorbance file: " & spcDirName & spcFilename & SPC_FILE_EXT), uniMsg)
                rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdSmplXVals, ProdAbsYVals, _
                                             GSpcIOLib.spcYType.spcYAbsrb, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
                                     
                If (rc = True) Then
                  rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
                End If
              End If
            Else
              Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Saving calculated product absorbance file: " & spcDirName & spcFilename & SPC_FILE_EXT), uniMsg)
              rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdSmplXVals, ProdAbsYVals, _
                                           GSpcIOLib.spcYType.spcYAbsrb, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
                                     
              If (rc = True) Then
                rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
              End If
            End If
          End If
                                       
SAVE_LAST_SCAN:
          ' Save backup of spectrum
          spcFilename = LAST_ABSORB_SCAN_FILE
          
          ' Check if spectrum data treated
          If (unity_main.m_enableTreatment = True) Then
            rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdSmplXVals, ProdTreatAbsYVals, _
                                         GSpcIOLib.spcYType.spcYAbsrb, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
                                     
            If (rc = True) Then
              rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
            End If
            
            ' Check if to save raw spectrum
            If (unity_main.m_saveSpectra = "Both") Then
              spcFilename = spcFilename & RAW_SCAN_APPEND
              rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdSmplXVals, ProdAbsYVals, _
                                           GSpcIOLib.spcYType.spcYAbsrb, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
                                     
              If (rc = True) Then
                rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
              End If
            End If
          Else
            rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdSmplXVals, ProdAbsYVals, _
                                         GSpcIOLib.spcYType.spcYAbsrb, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
                                     
            If (rc = True) Then
              rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
            End If
          End If
           
        Case SDT_PRODAVGABS          ' product calculated average absorbance of sample scans data type
          spectrumType = SpectrumTypes.AvgSpectrumType
          spcDirName = unity_main.m_saveDir
          spcFilename = unity_main.txtsamplename.Text
          Comment = MLSupport.GGS_Params("unity_main.statMsg15", "%1 (Average of %2 repacks)", Trim(unity_main.txtsampcomment.Text), CStr(unity_main.m_smplRepacks))
        
          If (unity_main.m_saveIt = "save") Then
            uniMsg = MLSupport.GGS_Params("unity_main.statMsg14", "Saving calculated product absorbance file: %1", spcDirName & spcFilename & SPC_FILE_EXT)
            
            ' Check if spectrum data treated
            If (unity_main.m_enableTreatment = True) Then
              Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Saving calculated product treated absorbance file: " & spcDirName & spcFilename & SPC_FILE_EXT), uniMsg)
              rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdSmplXVals, ProdTreatAvgAbsYVals, _
                                           GSpcIOLib.spcYType.spcYAbsrb, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
                                     
              If (rc = True) Then
                rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
              End If
              
              ' Check if to save raw spectrum
              If (unity_main.m_saveSpectra = "Both") Then
                spcFilename = spcFilename & RAW_SCAN_APPEND
                uniMsg = MLSupport.GGS_Params("unity_main.statMsg14", "Saving calculated product absorbance file: %1", spcDirName & spcFilename & SPC_FILE_EXT)
                Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Saving calculated product raw absorbance file: " & spcDirName & spcFilename & SPC_FILE_EXT), uniMsg)
                rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdSmplXVals, ProdAvgAbsYVals, _
                                             GSpcIOLib.spcYType.spcYAbsrb, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
                                     
                If (rc = True) Then
                  rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
                End If
              End If
            Else
              Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Saving calculated product absorbance file: " & spcDirName & spcFilename & SPC_FILE_EXT), uniMsg)
              rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdSmplXVals, ProdAvgAbsYVals, _
                                           GSpcIOLib.spcYType.spcYAbsrb, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
                                     
              If (rc = True) Then
                rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
              End If
            End If
          End If
        
          ' Save backup of averaged absorbance spectrum type
          spcFilename = LAST_ABSORB_SCAN_FILE
          
          ' Check if spectrum data treated
          If (unity_main.m_enableTreatment = True) Then
            rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdSmplXVals, ProdTreatAvgAbsYVals, _
                                         GSpcIOLib.spcYType.spcYAbsrb, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
                                     
            If (rc = True) Then
              rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
            End If
              
            ' Check if to save raw spectrum
            If (unity_main.m_saveSpectra = "Both") Then
              spcFilename = spcFilename & RAW_SCAN_APPEND
              rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdSmplXVals, ProdAvgAbsYVals, _
                                           GSpcIOLib.spcYType.spcYAbsrb, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
                                     
              If (rc = True) Then
                rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
              End If
            End If
          Else
            rc = SaveSpcFileSpectrumData(spcDirName, spcFilename, unity_main.m_scanTimestamp, GSpcIOLib.spcFileType.spcFileTypeEven, ProdSmplXVals, ProdAvgAbsYVals, _
                                         GSpcIOLib.spcYType.spcYAbsrb, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
                                     
            If (rc = True) Then
              rc = SaveSpcFileUnicodeData(spcDirName, spcFilename, unity_main.m_scanTimestamp, unity_main.m_sysSerialNum, Comment, m_instModel, "", unity_main.m_scanDataType, unity_main.repcounter, unity_main.m_smplRepacks, errStrg)
            End If
          End If
      End Select
    End If
  End If
  
  If (rc = True) Then
    ' Check if to send new spectrum file info to client appls
    If (spectrumType <> 0) Then
      IPCServer1.NewSpectrum 0, InfoStarStatusCodes.GoodStat, lblProd1.Caption, spectrumType, spcFilename & SPC_FILE_EXT, unity_main.repcounter, unity_main.m_smplRepacks
    End If
  Else      ' error saving spectrum file
    unity_main.m_ansiErrMsg = "Error saving spectrum file: " & spcDirName & spcFilename & SPC_FILE_EXT & ". " & errStrg
    unity_main.m_uniErrMsg = MLSupport.GGS_Params("unity_main.statMsg17", "Error saving spectrum file: %1", spcDirName & spcFilename & SPC_FILE_EXT)
    Call unity_main.log_scan_status(LOG_DBG_LEVEL1, unity_main.m_ansiErrMsg, unity_main.m_uniErrMsg)
    frm_status.lbl_statusCmd.ForeColor = vbRed
    frm_status.lbl_statusCmd.Caption = unity_main.m_uniErrMsg
  End If
  
  save_scan_data = rc
  Exit Function
  
OBJECT_ERROR:
  unity_main.m_ansiErrMsg = "Unity SVFComponent.dll component not installed or registered"
  unity_main.m_uniErrMsg = MLSupport.GGS_Params("errMsg7", "Unity %1 component not installed or registered", "SVFComponent.dll")
  Call unity_main.log_scan_status(LOG_DBG_LEVEL1, unity_main.m_ansiErrMsg, unity_main.m_uniErrMsg)
  frm_status.lbl_statusCmd.ForeColor = vbRed
  frm_status.lbl_statusCmd.Caption = unity_main.m_uniErrMsg
  save_scan_data = False
End Function

Public Sub save_scan_diff_format()
  Dim svfFileExist As Boolean
  Dim fileName As String
  Dim fileNum As Long
  Dim svfFileDiff As Boolean
  Dim labData() As Single
  Dim sampleId As String
  Dim Comment As String
  Dim rc As Long
  Dim errMsg As String
  Dim uniMsg As String
   
  CreatePath unity_main.m_saveDir
   
  ' Check if working with PRD model
  If (unity_main.m_prdModelType = True) Then
    ' Save data only if no STF file or STF file compatible
    If (unity_main.m_stfFileValid = True) Then
      On Error GoTo OBJECT_ERROR
    
      ' Check if SVF file exists to save daily scans spectra data
      svfFileExist = find_latest_svf_file(unity_main.m_saveDir, unity_main.lblProd1.Caption, fileName, fileNum)
      
      ' Check if SVF file different
      If (svfFileExist = True) And (unity_main.m_prodSVFChanged = True) Then
        fileName = Replace(fileName, "_S" & fileNum, "_S" & (fileNum + 1))
        unity_main.m_prodSVFChanged = False
        svfFileDiff = True
      End If

      unity_main.m_svfFileName = (unity_main.m_saveDir & fileName)
      
      ' Create file if does not exists or different
      If (svfFileExist = False) Or (svfFileDiff = True) Then
        rc = SVFObject.createFile_4("Unicode", unity_main.m_svfFileName, unity_main.m_prdFileName, unity_main.m_stfFileName, unity_main.m_sysSerialNum, unity_main.m_stfMasterSerNum, unity_main.m_instModel, "", unity_main.lblProd1.Caption, "", unity_main.m_svfStartWvln, unity_main.m_svfEndWvln, unity_main.m_svfWaveInc, unity_main.m_svfIsStd, SVF_LAB_BASIS, SVF_WAVE_TYPE, unity_main.m_numConstituents, SampleConstituentNames)
        
        If (rc <> 0) Then
          errMsg = unity_main.m_svfFileName & " UCal SVF spectra file error: " & rc
          uniMsg = MLSupport.GGS_Params("unity_main.errMsg7", "%1 UCal SVF spectra file error: %2", unity_main.m_svfFileName, CStr(rc))
          Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
          CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
          Exit Sub
        End If
      End If
    
      ReDim labData(unity_main.m_numConstituents - 1)
      sampleId = unity_main.txtsamplename.Text
      Comment = unity_main.txtsampcomment.Text
    
      rc = SVFObject.SaveSpectrum(unity_main.m_svfFileName, -1, unity_main.m_scanDblTimestamp, sampleId, Comment, SVFAbsYVals, 0, unity_main.m_numConstituents, labData)
      
      If (rc <> 0) Then
        errMsg = unity_main.m_svfFileName & " UCal SVF spectra file error: " & rc
        uniMsg = MLSupport.GGS_Params("unity_main.errMsg7", "%1 UCal SVF spectra file error: %2", unity_main.m_svfFileName, CStr(rc))
        Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
        CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", m_uniErrMsg), vbCritical
      End If
    End If
  End If
  
  Exit Sub
  
OBJECT_ERROR:
  errMsg = "Unity SVFComponent.dll component not installed or registered"
  uniMsg = MLSupport.GGS_Params("errMsg7", "Unity %1 component not installed or registered", "SVFComponent.dll")
  Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", m_uniErrMsg), vbCritical
End Sub

Public Sub plot_spectrum()
  Dim startXPt As Double
  Dim endXpt As Double
  Dim gradXpt As Double
  Dim dblXVals() As Double
  Dim dblYVals() As Double
  Dim snglYVals() As Single
  Dim varXVals As Variant
  Dim varYVals As Variant
  Dim errStrg As String
  Dim errMsg As String
  Dim uniMsg As String
  
  uniMsg = MLSupport.GSS("OperStatus", "status58", "Plotting scan spectrum")
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Plotting scan spectrum", uniMsg)
  
  XYPlot1.UnZoom
  XYPlot1.DisplayCursor False
  XYPlot1.DisplayGrids False
  
  lbl_scanDateTime.Caption = unity_main.m_scanTimestamp
      
  Select Case (unity_main.m_scanDataType)
    Case SDT_INTREFPPT            ' internal reference qualification scan data type
      dblXVals = ProdRefXVals
      dblYVals = ProdRefYVals
      startXPt = unity_main.m_intRefPPTStartWvln
      endXpt = unity_main.m_intRefPPTEndWvln
      XYPlot1.YAxisLabel = MLSupport.GSS("XYPlot", "YAxisLabel1", "Intensity")
    
    Case SDT_EXTREFPPT            ' external reference qualification scan data type
      dblXVals = ProdRefXVals
      dblYVals = ProdRefYVals
      startXPt = unity_main.m_extRefPPTStartWvln
      endXpt = unity_main.m_extRefPPTEndWvln
      XYPlot1.YAxisLabel = MLSupport.GSS("XYPlot", "YAxisLabel1", "Intensity")
      
    Case SDT_PRODINTREF           ' product internal reference scan data type
      dblXVals = ProdRefXVals
      dblYVals = ProdRefYVals
      startXPt = unity_main.m_smplStartWvln
      endXpt = unity_main.m_smplEndWvln
      XYPlot1.YAxisLabel = MLSupport.GSS("XYPlot", "YAxisLabel1", "Intensity")
      
    Case SDT_PRODEXTREF           ' product external reference scan data type
      dblXVals = ProdRefXVals
      dblYVals = ProdRefYVals
      startXPt = unity_main.m_extRefStartWvln
      endXpt = unity_main.m_extRefEndWvln
      XYPlot1.YAxisLabel = MLSupport.GSS("XYPlot", "YAxisLabel1", "Intensity")
      
    Case SDT_PRODPPT              ' product qualification (reflectance) scan data type
      dblXVals = ProdSmplXVals
      
      ' Check if spectrum data treated
      If (unity_main.m_enableTreatment = True) Then
        dblYVals = ProdTreatSmplYVals
      Else
        dblYVals = ProdSmplYVals
      End If
      
      startXPt = unity_main.m_smplPPTStartWvln
      endXpt = unity_main.m_smplPPTEndWvln
      XYPlot1.YAxisLabel = MLSupport.GSS("XYPlot", "YAxisLabel2", "Reflectance")
      
    Case SDT_PRODPPTABS           ' product calculated absorbance of qualification scan data type
      dblXVals = ProdSmplXVals
      
      ' Check if spectrum data treated
      If (unity_main.m_enableTreatment = True) Then
        dblYVals = ProdTreatAbsYVals
      Else
        dblYVals = ProdAbsYVals
      End If
      
      startXPt = unity_main.m_smplPPTStartWvln
      endXpt = unity_main.m_smplPPTEndWvln
      XYPlot1.YAxisLabel = MLSupport.GSS("XYPlot", "YAxisLabel3", "Absorbance")
    
    Case SDT_PRODSMPL             ' product sample (reflectance) scan data type
      dblXVals = ProdSmplXVals
      
      ' Check if spectrum data treated
      If (unity_main.m_enableTreatment = True) Then
        dblYVals = ProdTreatSmplYVals
      Else
        dblYVals = ProdSmplYVals
      End If
      
      startXPt = unity_main.m_smplStartWvln
      endXpt = unity_main.m_smplEndWvln
      XYPlot1.YAxisLabel = MLSupport.GSS("XYPlot", "YAxisLabel2", "Reflectance")
      
    Case SDT_PRODSMPLABS          ' product calculated absorbance of sample scan data type
      dblXVals = ProdSmplXVals
      
      ' Check if spectrum data treated
      If (unity_main.m_enableTreatment = True) Then
        dblYVals = ProdTreatAbsYVals
      Else
        dblYVals = ProdAbsYVals
      End If
      
      startXPt = unity_main.m_smplStartWvln
      endXpt = unity_main.m_smplEndWvln
      XYPlot1.YAxisLabel = MLSupport.GSS("XYPlot", "YAxisLabel3", "Absorbance")
      
    Case SDT_PRODAVGABS           ' product calculated average absorbance of sample scans data type
      dblXVals = ProdSmplXVals
      
      ' Check if spectrum data treated
      If (unity_main.m_enableTreatment = True) Then
        dblYVals = ProdTreatAvgAbsYVals
      Else
        dblYVals = ProdAvgAbsYVals
      End If
      
      startXPt = unity_main.m_smplStartWvln
      endXpt = unity_main.m_smplEndWvln
      XYPlot1.YAxisLabel = MLSupport.GSS("XYPlot", "YAxisLabel3", "Absorbance")
  End Select
      
  Dim ii, nn As Integer
  
  nn = UBound(dblYVals)
  ReDim snglYVals(nn)
  
  varXVals = dblXVals
  
  For ii = 0 To nn
    snglYVals(ii) = dblYVals(ii)
  Next ii
      
  varYVals = snglYVals
  
#If ABBFT Then
  startXPt = unity_main.m_mb3000.m_startWavenum
  endXpt = unity_main.m_mb3000.m_endWavenum
  gradXpt = unity_main.m_mb3000.m_waveNumIncr
#Else
  gradXpt = MS11CfgData.wvlnIncr
#End If

  If (XYPlot1.PlotSpectrum2(0, startXPt, endXpt, gradXpt, SubsetDataTypes.SS_VARIANT, vbNull, vbNull, varXVals, varYVals, False, errStrg) = False) Then
    errMsg = "Error plotting spectrum." & errStrg
    uniMsg = MLSupport.GSS("OperStatus", "status59", "Error plotting spectrum")
    Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
  End If
End Sub

Public Sub do_pred()
  Dim freefilex As Integer
  Dim zz, ii As Integer
  Dim modExt As String
  Dim errMsg As String
  Dim uniMsg As String

  unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status7", "Calculating Property Values")
  uniMsg = MLSupport.GSS("OperStatus", "status61", "Performing property predictions")
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Performing property predictions", uniMsg)

  clear_pred_results
  
  For zz = 1 To frmedmod.numprops.Text
    frmedmod.grid_models.Row = zz
    frmedmod.grid_models.Col = 1
    uniMsg = MLSupport.GGS_Params("unity_main.statMsg25", "Calculating value for property %1", Trim(frmedmod.grid_models.Text))
    Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Calculating value for property " & Trim(frmedmod.grid_models.Text), uniMsg)
    
    frmedmod.grid_models.Col = 2
    unity_main.modlname = Trim(frmedmod.grid_models.Text)
        
    If (unity_main.modlname = "") Then GoTo donehere:
        
    modExt = ("." & LCase(CFile.st_FileExt(unity_main.modlname)))

    Select Case (modExt)
      Case GRAMS_MODEL_FILE_EXT       ' GRAMS PLSIQ Model
        unity_main.modltype = 1
      Case MLR_MODEL_FILE_EXT         ' MLR Model
        unity_main.modltype = 2
      Case SEC_MODEL_FILE_EXT         ' Secondary Model - calculated from other properties
        unity_main.modltype = 3
      Case PRD_MODEL_FILE_EXT         ' PRD model
        unity_main.modltype = 4
      
        ' Check if PRD software is not installed
        If (unity_main.m_prdEnabled = False) Then
          errMsg = "Problem with product trying to use a PRD model without UCal PRD software installed"
          unity_main.errorstring = errMsg
          unity_main.write_error (LOG_DBG_LEVEL1)
          uniMsg = MLSupport.GSS("unity_main", "errMsg8", "Problem with product trying to use a PRD model without UCal PRD software installed")
          CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
          unity_main.pukedonpred = True
          GoTo PRED_FAILED
        End If

        ' Check if .NET Framework 2.0 is not installed
        If (unity_main.m_netFWInstalled = False) Then
          errMsg = "Problem with product trying to use an UCal PRD model without MS .NET Framework 2.0 installed"
          unity_main.errorstring = errMsg
          unity_main.write_error (LOG_DBG_LEVEL1)
          uniMsg = MLSupport.GSS("unity_main", "errMsg6", "Problem with product trying to use an UCal PRD model without MS .NET Framework 2.0 installed")
          CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
          unity_main.pukedonpred = True
          GoTo PRED_FAILED
        End If
      Case CALSTAR_MODEL_FILE_EXT     ' Senslogic CalStar model
        unity_main.modltype = 5
                      
        ' Check if CalStar is not installed
        If (unity_main.calstar_enabled = False) Then
          errMsg = "Problem with product trying to use a CalStar model without CalStar software installed"
          unity_main.errorstring = errMsg
          unity_main.write_error (LOG_DBG_LEVEL1)
          uniMsg = MLSupport.GSS("unity_main", "errMsg1", "Problem with product trying to use a CalStar model without CalStar software installed")
          CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
          unity_main.pukedonpred = True
          GoTo PRED_FAILED
        End If
      Case Else
        If (unity_main.modlname = "RUNTIME ENTRY") Then
          errMsg = "RUNTIME ENTRY model no longer supported"
          uniMsg = MLSupport.GSS("unity_main", "errMsg2", "RUNTIME ENTRY model no longer supported")
        Else
          errMsg = (Trim(unity_main.modlname) & " model file has invalid extension")
          uniMsg = MLSupport.GGS_Params("unity_main.errMsg5", "%1 model file has invalid extension", Trim(unity_main.modlname))
        End If
        
        unity_main.errorstring = errMsg
        unity_main.write_error (LOG_DBG_LEVEL1)
        CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
        unity_main.pukedonpred = True
        GoTo PRED_FAILED
    End Select

    unity_main.lst_modtype.AddItem (unity_main.modltype)
    frmedmod.grid_models.Col = 3
        
    ' Setup additional data based on model type
    Select Case (unity_main.modltype)
      Case 4          ' UCal PRD
        unity_main.m_prdConstituent = Trim(frmedmod.grid_models.Text)
        unity_main.modlindex = 1
        frmedmod.grid_models.Col = 15
      Case 5          ' CalStar
        unity_main.slcal = Trim(frmedmod.grid_models.Text)
        unity_main.modlindex = 1
      Case Else    ' GRAMS, MLR & Secondary
        unity_main.modlindex = Trim(frmedmod.grid_models.Text)
    End Select
        
    frmedmod.grid_models.Col = 4
    unity_main.tempbias = frmedmod.grid_models.Value
    unity_main.lstint.AddItem (unity_main.tempbias)
    frmedmod.grid_models.Col = 5
    unity_main.tempskew = frmedmod.grid_models.Value
    unity_main.lstslope.AddItem (unity_main.tempskew)

    Load mainform
    mainform.Visible = False
    mainform.lmpred
        
    If (unity_main.pukedonpred = True) Then GoTo PRED_FAILED
    
    ' Pause to allow background events
    DoEvents
  Next zz
    
donehere:
  For ii = 0 To Int(frmedmod.numprops.Text) - 1
    frmpreds.gridpreds.Row = ii + 1
    frmpreds.gridpreds.Col = unity_main.repcounter
    unity_main.tempval = unity_main.preds.List(ii)
    frmedmod.grid_models.Row = ii + 1
    frmedmod.grid_models.Col = 5 ' skew
    unity_main.tempskew = frmedmod.grid_models.Value
    frmedmod.grid_models.Col = 4 'bias/intercept
    unity_main.tempbias = frmedmod.grid_models.Value
    frmpreds.gridpreds.Text = unity_main.tempval
      
    If (unity_main.m_smplRepacks = 1) Then
      unity_main.fpspread_pred.Col = 2
      unity_main.fpspread_pred.Row = (ii + 1)
      unity_main.fpspread_pred.Text = unity_main.tempval
    End If
  Next ii
    
  Exit Sub
    
PRED_FAILED:
  unity_main.errorstring = "Problem with predictions model " & unity_main.fullmodelname
  unity_main.write_error
  
  ' Check if not performing batch scan
  If (unity_main.m_batchRunFlg = False) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("unity_main", "errMsg3", "Problem with predictions. Please check system and/or scan log for more details"), vbCritical
  End If
End Sub

Public Sub checkbounds()
  Dim cond2 As Single
  Dim cond4 As Single
  Dim cvalue As Single
  Dim ii As Integer
  Dim uniMsg As String

  ' unity_main.m_valueBound
  ' 0 = act value
  ' 1 = bound min
  ' 2 = bound max
  ' 3 = both
  ' cond2=low fail
  ' cond4=high fail

  uniMsg = MLSupport.GSS("OperStatus", "status62", "Checking property value bounds")
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Checking property value bounds", uniMsg)

  For ii = 0 To Int(frmedmod.numprops.Text) - 1
    frmedmod.grid_models.Row = ii + 1
    frmedmod.grid_models.Col = 12
    cond2 = frmedmod.grid_models.Value
    
    frmedmod.grid_models.Col = 14
    cond4 = frmedmod.grid_models.Value
    
    unity_main.fpspread_pred.Col = 2
    unity_main.fpspread_pred.Row = ii + 1
    cvalue = unity_main.fpspread_pred.Value
    
    If (cvalue > cond4) And (unity_main.m_valueBound > 1) Then
      unity_main.fpspread_pred.Text = cond4
    End If
    
    If (cvalue < cond2) Then
      If (unity_main.m_valueBound = 1) Then
        unity_main.fpspread_pred.Text = cond2
      Else
        If (unity_main.m_valueBound = 3) Then
          unity_main.fpspread_pred.Text = cond2
        End If
      End If
    End If
  Next ii
End Sub

Public Sub chksigfigs()
  Dim sigfig, ii As Integer

  frmedmod.grid_models.Col = 6

  For ii = 1 To Int(frmedmod.numprops.Text)
    frmedmod.grid_models.Row = ii
    sigfig = frmedmod.grid_models.Value
    unity_main.fpspread_pred.Row = ii
    unity_main.fpspread_pred.Col = 2
    unity_main.fpspread_pred.CellType = CellTypeNumber

    Select Case sigfig
      Case 0
        unity_main.fpspread_pred.TypeNumberDecPlaces = 0
      Case 1
        unity_main.fpspread_pred.TypeNumberDecPlaces = 1
      Case 2
        unity_main.fpspread_pred.TypeNumberDecPlaces = 2
      Case 3
        unity_main.fpspread_pred.TypeNumberDecPlaces = 3
      Case 4
        unity_main.fpspread_pred.TypeNumberDecPlaces = 4
      Case 5
        unity_main.fpspread_pred.TypeNumberDecPlaces = 5
      Case 6
        unity_main.fpspread_pred.TypeNumberDecPlaces = 6
      Case 7
        unity_main.fpspread_pred.TypeNumberDecPlaces = 7
      Case 8
        unity_main.fpspread_pred.TypeNumberDecPlaces = 8
    End Select
  Next ii
End Sub

Public Sub clear_pred_results()

  unity_main.pukedonpred = False
  unity_main.lst_modtype.Clear
  unity_main.lstmd.Clear
  unity_main.lstrr.Clear
  unity_main.lstrr2.Clear
  unity_main.lstresrat.Clear
  unity_main.lst_qual.Clear
  unity_main.lst_nd.Clear
  unity_main.lst_pfexp.Clear
  unity_main.lstint.Clear
  unity_main.lstslope.Clear
  unity_main.preds.Clear
End Sub

Public Sub disp_no_pred_results()
  Dim nn As Integer

  For nn = 1 To frmedmod.numprops.Text
    unity_main.fpspread_pred.Row = nn
    unity_main.fpspread_pred.Col = 2
    unity_main.fpspread_pred.Text = unity_main.m_noPredVal
  Next nn
End Sub

Public Sub write_report_file()
  Dim reportFileName As String
  Dim printString As String
  Dim jj As Integer
  Dim errMsg As String
  Dim uniFile As New clsUniFile
  Dim uniMsg As String

  reportFileName = unity_main.m_savePredFile
  
  uniMsg = MLSupport.GGS_Params("unity_main.statMsg19", "Writing sample results to report file: %1", reportFileName)
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Writing sample results to report file: " & reportFileName), uniMsg)
   
  On Error GoTo FILE_ERROR
  CreatePath CFile.st_FilePath(reportFileName)
  
  If (uniFile.st_FileExist(reportFileName) = True) Then
    If (uniFile.OpenFileAppend(reportFileName) = False) Then GoTo FILE_ERROR
  Else
    If (uniFile.OpenFileWrite(reportFileName) = False) Then GoTo FILE_ERROR
    uniFile.WriteBOM fe_UTF16LE
  End If
  
  uniFile.WriteUnicodeLine ""
  uniFile.WriteUnicodeLine "------------------------------------------------------"
  uniFile.WriteUnicodeLine ""
  uniFile.WriteUnicodeLine MLSupport.GSS("unity_main", "statMsg2", "Unity Scientific Analysis Report")
  uniFile.WriteUnicodeLine MLSupport.GGS_Params("unity_main.statMsg20", "Analysis Date: %1", unity_main.lbl_date.Caption)
  uniFile.WriteUnicodeLine MLSupport.GGS_Params("unity_main.statMsg21", "Analysis Time: %1", unity_main.lbl_time.Caption)
  uniFile.WriteUnicodeLine MLSupport.GSS("Headers", "serNum", "Serial No.") & ": " & unity_main.m_sysSerialNum
  uniFile.WriteUnicodeLine MLSupport.GGS_Params("unity_main.statMsg22", "Sample Type: %1", unity_main.lblProd1.Caption)
  uniFile.WriteUnicodeLine MLSupport.GGS_Params("unity_main.statMsg23", "Sample Name: %1", Trim(unity_main.txtsamplename.Text))
  uniFile.WriteUnicodeLine MLSupport.GGS_Params("unity_main.statMsg24", "Sample Comment: %1", unity_main.txtsampcomment.Text)
  uniFile.WriteUnicodeLine ""
  uniFile.WriteUnicodeLine MLSupport.GSS("unity_main", "statMsg3", "Property, Value")
    
  For jj = 1 To Trim(frmedmod.numprops.Text)
    printString = ""
    unity_main.fpspread_pred.Row = jj
    unity_main.fpspread_pred.Col = 1
    printString = (Trim(unity_main.fpspread_pred.Text) & " = ")
    unity_main.fpspread_pred.Col = 2
    printString = (printString & Trim(unity_main.fpspread_pred.Text))
    uniFile.WriteUnicodeLine printString
  Next jj
  
  uniFile.WriteUnicodeLine ""
  uniFile.WriteUnicodeLine "------------------------------------------------------"
  uniFile.WriteUnicodeLine ""
  uniFile.Flush
  uniFile.CloseFile
  Exit Sub
  
FILE_ERROR:
  uniFile.CloseFile
  errMsg = (reportFileName & " file write error. " & Error$)
  unity_main.errorstring = errMsg
  unity_main.write_error
  uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", reportFileName, Error$)
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub

Public Sub add_last50()
  Dim ret As Boolean
  Dim rowCnt, jj As Long
  Dim colCnt As Integer
  Dim maxCol As Long
  Dim uniMsg As String

  rowCnt = ss_last50.DataRowCnt
  
  ' Delete older data if beyond limit
  If (rowCnt > (ss_last50.MaxRows - 1)) Then
    ss_last50.DeleteRows 1, rowCnt - (ss_last50.MaxRows - 1)
    
    If (lst_last50MaxCols.ListCount > 0) Then
      lst_last50MaxCols.RemoveItem 0
    End If
  End If
  
  ' Determine max columns to display
  maxCol = frmedmod.numprops.Text * 2 + MIN_LAST50_COLS
  
  ' Save max columns used for row
  lst_last50MaxCols.AddItem maxCol
  
  For rowCnt = 0 To lst_last50MaxCols.ListCount - 1
    ' Get max column used for row
    colCnt = lst_last50MaxCols.List(rowCnt)
        
    If (colCnt > maxCol) Then
      maxCol = colCnt
    End If
  Next rowCnt

  ss_last50.MaxCols = maxCol
  
  ss_last50.Row = ss_last50.DataRowCnt + 1
  ss_last50.Col = 1
  ss_last50.Text = lbl_date.Caption
  ss_last50.Col = 2
  ss_last50.Text = Trim(lbl_time.Caption)
  ss_last50.Col = 3
  ss_last50.Text = Trim(lblProd1.Caption)
  ss_last50.Col = 4
  ss_last50.Text = Trim(txtsamplename.Text)
  ss_last50.Col = 5
  
  If (Trim(txtsampcomment.Text) <> "") Then
    ss_last50.Text = Trim(txtsampcomment.Text)
  Else
    ss_last50.Text = " "
  End If

  For jj = 1 To Trim(frmedmod.numprops.Text)
    fpspread_pred.Row = jj
    fpspread_pred.Col = 1             'propname
    ss_last50.Col = MIN_LAST50_COLS + (2 * jj) - 1
    ss_last50.Text = Trim(fpspread_pred.Text)
    fpspread_pred.Col = 2             'value
    ss_last50.Col = MIN_LAST50_COLS + (2 * jj)
    ss_last50.Text = Trim(fpspread_pred.Text)
  Next jj
  
  ss_last50.Row = 0
  ss_last50.Col = 0
  ss_last50.Row2 = 0
  ss_last50.Col2 = ss_last50.MaxCols
  ss_last50.BlockMode = True
  ss_last50.Font.Bold = True
  ss_last50.Font.Name = "Arial Unicode MS"
  ss_last50.Font.Size = 8
  ss_last50.BlockMode = False
  
  ss_last50.Row = 1
  ss_last50.Col = 0
  ss_last50.Row2 = ss_last50.MaxRows
  ss_last50.Col2 = ss_last50.MaxCols
  ss_last50.BlockMode = True
  ss_last50.Font.Bold = False
  ss_last50.Font.Name = "Arial Unicode MS"
  ss_last50.Font.Size = 8
  ss_last50.BlockMode = False
  
  pos_last50_toprow
  build_last50_colhdr
  
  uniMsg = MLSupport.GGS_Params("unity_main.statMsg26", "Writing sample results to historical file: %1", (REPORTS_DIR & HIST_LOG_FILE))
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Writing sample results to historical file: " & (REPORTS_DIR & HIST_LOG_FILE)), uniMsg)
  
  ret = ss_last50.SaveToFile((REPORTS_DIR & HIST_LOG_FILE), False)
End Sub

#If SSTAR Then
Public Sub disp_verify_ref_button(visibleState As Boolean)
  
  ' Reset verification reminder
  m_intRefVerReminderFlg = False
  
  If (unity_main.m_intRefVerReqdFlg <> visibleState) Then
    unity_main.m_intRefVerReqdFlg = visibleState
    unity_main.cmd_verifyRef.Visible = visibleState
  
    ' Check if to display Verify Ref button
    If (visibleState = True) Then
      ' Check if in Run/Operator mode
      If (unity_main.run_min_gui = True) Then
        ' Resize and move Verify Reference button
        unity_main.cmd_verifyRef.Left = unity_main.cmd_select.Left
        unity_main.cmd_verifyRef.Width = unity_main.cmd_select.Width
      Else      ' Supervisor mode
        ' Resize Options and Verify Reference button
        unity_main.cmd_options.Width = unity_main.cmd_runBatch.Width
        unity_main.cmd_verifyRef.Left = unity_main.cmd_runBatch.Left
        unity_main.cmd_verifyRef.Width = unity_main.cmd_runBatch.Width
      End If
    Else
      ' Check if in Supervisor mode
      If (unity_main.run_min_gui = False) Then
        ' Resize Options button
        unity_main.cmd_options.Width = unity_main.cmd_select.Width
      End If
    End If
  End If
End Sub
#End If

Private Sub build_inst_model()
  Dim instStrg As String
  Dim modelStrg As String
  Dim devIDStrg As String
  
#If ABBFT Then
  m_instModel = "US STW (No Int Ref)"
#Else
  instStrg = "SS"

  Select Case (MS11CfgData.devID)
    Case DTID_DRAWER0
      devIDStrg = "DRW"
    Case DTID_TOPWIND0
      devIDStrg = "STW (No Int Ref)"
    Case DTID_DRAWER1
      devIDStrg = "SDRW"
    Case DTID_TOPWIND1
      If (MS11CfgData.nTrays = 1) Then
        devIDStrg = "STW"
      Else
        devIDStrg = "RTW"
      End If
  End Select
  
  Select Case (MS11CfgData.smplTblIX)
    Case 2
      modelStrg = "2200"
      
      ' Check if top window
      If (MS11CfgData.devID = DTID_TOPWIND0) Or (MS11CfgData.devID = DTID_TOPWIND1) Then
        ' Check spectral ending wavelen for 1200 model
        If (MS11CfgData.spxEndWvln = 2180) Then
          modelStrg = "1200"
          instStrg = "IS"
        End If
      End If
    Case 4
      modelStrg = "2400"
    Case 6
      modelStrg = "2500X"
    Case 7
      modelStrg = "2500"
      m_sys2500 = True
    Case 9
      modelStrg = "1400"
      instStrg = "IS"
  End Select
  
  m_instModel = instStrg & modelStrg & " " & devIDStrg
#End If
End Sub

Private Sub setup_spreadsheets_maxrows()
  Dim numPts As Integer

  unity_main.fpspread_pred.MaxRows = MAX_NUM_PROPS
  frmedmod.grid_models.MaxRows = MAX_NUM_PROPS
  frm_repacks.ss_repacks.MaxRows = MAX_NUM_PROPS
  frm_edbias.ss_biases.MaxRows = MAX_NUM_PROPS
  frmpreds.gridpreds.MaxRows = MAX_NUM_PROPS
  
#If ABBFT Then
  numPts = (unity_main.m_mb3000.m_endWavenum - unity_main.m_mb3000.m_startWavenum) / unity_main.m_mb3000.m_waveNumIncr + 1
#Else
  numPts = (MS11CfgData.maxWvln - MS11CfgData.minWvln) / MS11CfgData.wvlnIncr + 1
#End If
  frm_mlrscan.gridscan.MaxRows = numPts
End Sub

Private Sub init_defaults()
  Dim trayIndx As Integer
  Dim adaptIndx As Integer
  Dim multiCupIndx As Integer
  
#If SSTAR Then
  ' Check if Spectrix feature enabled
  If (MS11CfgData.spxStartWvln = 0) And (MS11CfgData.spxEndWvln = 0) Then
    ' Save min/max spectrum wavelength used by instrument
    unity_main.m_minWvln = MS11CfgData.minWvln
    unity_main.m_maxWvln = MS11CfgData.maxWvln
    unity_main.m_spectrixEnable = False
  Else
    ' Save min/max spectrum wavelength used by Spectrix file
    unity_main.m_minWvln = MS11CfgData.spxStartWvln
    unity_main.m_maxWvln = MS11CfgData.spxEndWvln
    unity_main.m_spectrixEnable = True
  End If
 
  MS11MultiCupInfo(0).cfgName = CFG_NONE_MCT
  
  ' Setup reference and tray configuration options based on instrument device ID
  Select Case (MS11CfgData.devID)
    Case DTID_DRAWER0            ' SS2200/SS2400 standard drawer system
      frm_collect.optbginternal.Visible = True
      frm_collect.optbgexternal.Visible = True
      frm_collect.optbgfile.Visible = True
      
      ' Process tray info
      For trayIndx = 1 To MS11CfgData.nTrays
        Select Case (MS11TrayInfoData(trayIndx).trayID)
          Case TTID_UNKNOWN          ' Reference/Sample Configuration Unknown
            MS11AdapterInfo(trayIndx).cfgName = ""
            MS11AdapterInfo(trayIndx).dspName = ""
            MS11AdapterInfo(trayIndx).trayNum = trayIndx
          Case TTID_1SAMPLE          ' Universal Tray: Always can Scan 1 Sample
            MS11AdapterInfo(trayIndx).cfgName = CFG_STATIC_CUP_AT
            MS11AdapterInfo(trayIndx).dspName = MLSupport.GSS("frm_collect", "ATStaticCup", "Stationary Sample Cup")
            MS11AdapterInfo(trayIndx).trayNum = trayIndx
          Case TTID_48POS            ' 48 Position Platter
            MS11AdapterInfo(trayIndx).cfgName = CFG_ROTATE_CUP_AT
            MS11AdapterInfo(trayIndx).dspName = MLSupport.GSS("frm_collect", "ATRotateCup", "Rotating Sample Cup")
            MS11AdapterInfo(trayIndx).trayNum = trayIndx
        End Select
      Next trayIndx
      
      For adaptIndx = 1 To MAX_ADAPTER_TYPES
        If (MS11AdapterInfo(adaptIndx).trayNum <> 0) Then
          frm_collect.opt_adapterType(adaptIndx).Caption = MS11AdapterInfo(adaptIndx).dspName
          frm_extRef.opt_adapterType(adaptIndx).Caption = MS11AdapterInfo(adaptIndx).dspName
        Else
          frm_collect.opt_adapterType(adaptIndx).Visible = False
          frm_extRef.opt_adapterType(adaptIndx).Visible = False
        End If
      Next adaptIndx
    
    Case DTID_TOPWIND0           ' Top window w/out internal reflectance
      ' Hide internal reference qualification setup button
      frm_backrestore.cmd_pptRef.Visible = False
      
      frm_collect.optbginternal.Visible = False
      frm_collect.optbgexternal.Visible = True
      frm_collect.optbgexternal.Visible = True
      
      ' Process tray info
      For trayIndx = 1 To MS11CfgData.nTrays
        Select Case (MS11TrayInfoData(trayIndx).trayID)
          Case TTID_UNKNOWN          ' Reference/Sample Configuration Unknown
            MS11AdapterInfo(trayIndx).cfgName = ""
            MS11AdapterInfo(trayIndx).dspName = ""
            MS11AdapterInfo(trayIndx).trayNum = 0
          Case TTID_1SAMPLE          ' Universal Tray
            MS11AdapterInfo(trayIndx).cfgName = CFG_SAMP_WINDOW_AT
            MS11AdapterInfo(trayIndx).dspName = MLSupport.GSS("frm_collect", "ATSampWindow", "Sample Window")
            MS11AdapterInfo(trayIndx).trayNum = trayIndx
        End Select
      Next trayIndx
      
      For adaptIndx = 1 To MAX_ADAPTER_TYPES
        If (MS11AdapterInfo(adaptIndx).trayNum <> 0) Then
          frm_collect.opt_adapterType(adaptIndx).Caption = MS11AdapterInfo(adaptIndx).dspName
          frm_extRef.opt_adapterType(adaptIndx).Caption = MS11AdapterInfo(adaptIndx).dspName
        Else
          frm_collect.opt_adapterType(adaptIndx).Visible = False
          frm_extRef.opt_adapterType(adaptIndx).Visible = False
        End If
      Next adaptIndx
    
    Case DTID_DRAWER1            ' SS2200/SS2400 drawer w/out stepper system
      frm_collect.optbginternal.Visible = True
      frm_collect.optbgexternal.Visible = True
      frm_collect.optbgexternal.Visible = True
      
      ' Process tray info
      For trayIndx = 1 To MS11CfgData.nTrays
        Select Case (MS11TrayInfoData(trayIndx).trayID)
          Case TTID_UNKNOWN          ' Reference/Sample Configuration Unknown
            MS11AdapterInfo(trayIndx).cfgName = ""
            MS11AdapterInfo(trayIndx).dspName = ""
            MS11AdapterInfo(trayIndx).trayNum = 0
          Case TTID_1SAMPLE          ' Universal Tray: Always can Scan 1 Sample
            MS11AdapterInfo(trayIndx).cfgName = CFG_STATIC_CUP_AT
            MS11AdapterInfo(trayIndx).dspName = MLSupport.GSS("frm_collect", "ATStaticCup", "Stationary Sample Cup")
            MS11AdapterInfo(trayIndx).trayNum = trayIndx
        End Select
      Next trayIndx
      
      For adaptIndx = 1 To MAX_ADAPTER_TYPES
        If (MS11AdapterInfo(adaptIndx).trayNum <> 0) Then
          frm_collect.opt_adapterType(adaptIndx).Caption = MS11AdapterInfo(adaptIndx).dspName
          frm_extRef.opt_adapterType(adaptIndx).Caption = MS11AdapterInfo(adaptIndx).dspName
        Else
          frm_collect.opt_adapterType(adaptIndx).Visible = False
          frm_extRef.opt_adapterType(adaptIndx).Visible = False
        End If
      Next adaptIndx
      
    Case DTID_TOPWIND1           ' Top window with internal reflectance
      frm_collect.optbginternal.Visible = True
      frm_collect.optbgexternal.Visible = True
      frm_collect.optbgexternal.Visible = True
      
      ' Process tray info
      adaptIndx = 1
      multiCupIndx = 1
      
      For trayIndx = 1 To MS11CfgData.nTrays
        Select Case (MS11TrayInfoData(trayIndx).trayID)
          Case TTID_UNKNOWN          ' Reference/Sample Configuration Unknown
            MS11AdapterInfo(adaptIndx).cfgName = ""
            MS11AdapterInfo(adaptIndx).dspName = ""
            MS11AdapterInfo(adaptIndx).trayNum = 0
            adaptIndx = adaptIndx + 1
            
          Case TTID_1SAMPLE          ' Universal Tray: Always can Scan 1 Sample
            MS11AdapterInfo(adaptIndx).cfgName = CFG_SAMP_WINDOW_AT
            MS11AdapterInfo(adaptIndx).dspName = MLSupport.GSS("frm_collect", "ATSampWindow", "Sample Window")
            MS11AdapterInfo(adaptIndx).trayNum = trayIndx
            adaptIndx = adaptIndx + 1
            MS11AdapterInfo(adaptIndx).cfgName = CFG_IRIS_AT
            MS11AdapterInfo(adaptIndx).dspName = MLSupport.GSS("frm_collect", "ATIris", "Iris Adapter")
            MS11AdapterInfo(adaptIndx).trayNum = trayIndx
            adaptIndx = adaptIndx + 1
            
            MS11AdapterInfo(adaptIndx).cfgName = CFG_ISI_RING_AT
            MS11AdapterInfo(adaptIndx).dspName = MLSupport.GSS("frm_collect", "ATISIRing", "ISI Ring Adapter")
            MS11AdapterInfo(adaptIndx).trayNum = trayIndx
            adaptIndx = adaptIndx + 1
            
          Case TTID_48POS            ' 48 Position Platter
            MS11AdapterInfo(adaptIndx).cfgName = CFG_SINGLE_POS_ROTATE_AT
            MS11AdapterInfo(adaptIndx).dspName = MLSupport.GSS("frm_collect", "ATSinglePosRotate", "Single Position Rotating Adapter")
            MS11AdapterInfo(adaptIndx).trayNum = trayIndx
            adaptIndx = adaptIndx + 1
            
            MS11MultiCupInfo(multiCupIndx).cfgName = CFG_MEDIUM_MCT
            MS11MultiCupInfo(multiCupIndx).dspName = MLSupport.GSS("frm_collect", "MCTMedium", "Medium or Spout Sample Cup (Black Pos.)")
            MS11MultiCupInfo(multiCupIndx).trayNum = trayIndx
            multiCupIndx = multiCupIndx + 1
            MS11MultiCupInfo(multiCupIndx).cfgName = CFG_BL_LARGE_MCT
            MS11MultiCupInfo(multiCupIndx).dspName = MLSupport.GSS("frm_collect", "MCTBLLarge", "B&&L Large Sample Cup (Black Position)")
            MS11MultiCupInfo(multiCupIndx).trayNum = trayIndx
            multiCupIndx = multiCupIndx + 1
            MS11MultiCupInfo(multiCupIndx).cfgName = CFG_PETRI_MCT
            MS11MultiCupInfo(multiCupIndx).dspName = MLSupport.GSS("frm_collect", "MCTPetri", "Petri Dish Adapter (Black Position)")
            MS11MultiCupInfo(multiCupIndx).trayNum = trayIndx
            multiCupIndx = multiCupIndx + 1

          Case TTID_16POS            ' 16 Position Tray (ISI)
            MS11AdapterInfo(adaptIndx).cfgName = CFG_MULTI_CUP_AT
            MS11AdapterInfo(adaptIndx).dspName = MLSupport.GSS("frm_collect", "ATMultiCup", "Multi-Cup Adapter")
            MS11AdapterInfo(adaptIndx).trayNum = trayIndx
            adaptIndx = adaptIndx + 1
            
            MS11MultiCupInfo(multiCupIndx).cfgName = CFG_ISI_RING_MCT
            MS11MultiCupInfo(multiCupIndx).dspName = MLSupport.GSS("frm_collect", "MCTISIRing", "ISI Ring Cup (Red Position)")
            MS11MultiCupInfo(multiCupIndx).trayNum = trayIndx
            multiCupIndx = multiCupIndx + 1
            MS11MultiCupInfo(multiCupIndx).cfgName = CFG_BL_SMALL_MCT
            MS11MultiCupInfo(multiCupIndx).dspName = MLSupport.GSS("frm_collect", "MCTBLSmall", "B&&L Small Powder Cup (Red Position)")
            MS11MultiCupInfo(multiCupIndx).trayNum = trayIndx
            multiCupIndx = multiCupIndx + 1
            
          Case TTID_100POS           ' 100 Position Tray
            If (MS11AdapterInfo(adaptIndx - 1).cfgName <> CFG_MULTI_CUP_AT) Then
              MS11AdapterInfo(adaptIndx).cfgName = CFG_MULTI_CUP_AT
              MS11AdapterInfo(adaptIndx).dspName = MLSupport.GSS("frm_collect", "ATMultiCup", "Multi-Cup Adapter")
              MS11AdapterInfo(adaptIndx).trayNum = trayIndx
              adaptIndx = adaptIndx + 1
            End If
            
            MS11MultiCupInfo(multiCupIndx).cfgName = CFG_LARGE_MCT
            MS11MultiCupInfo(multiCupIndx).dspName = MLSupport.GSS("frm_collect", "MCTLarge", "Large Sample Cup (Blue Position)")
            MS11MultiCupInfo(multiCupIndx).trayNum = trayIndx
            multiCupIndx = multiCupIndx + 1
        End Select
      Next trayIndx
      
      For adaptIndx = 1 To MAX_ADAPTER_TYPES
        If (MS11AdapterInfo(adaptIndx).trayNum <> 0) Then
          frm_collect.opt_adapterType(adaptIndx).Caption = MS11AdapterInfo(adaptIndx).dspName
          frm_extRef.opt_adapterType(adaptIndx).Caption = MS11AdapterInfo(adaptIndx).dspName
        Else
          frm_collect.opt_adapterType(adaptIndx).Visible = False
          frm_extRef.opt_adapterType(adaptIndx).Visible = False
        End If
      Next adaptIndx
      
      For multiCupIndx = 1 To MAX_MULTI_CUP_TYPES
        frm_collect.opt_multiCupType(multiCupIndx).Caption = MS11MultiCupInfo(multiCupIndx).dspName
        frm_extRef.opt_multiCupType(multiCupIndx).Caption = MS11MultiCupInfo(multiCupIndx).dspName
      Next multiCupIndx
  End Select
  
  unity_main.m_extRefFileSetup = False
  unity_main.m_extRefPPTFileSetup = False
  unity_main.m_intRefPPTFileSetup = False
  unity_main.m_olRefFileSetup = False
  unity_main.m_lastEndWvln = 0
  unity_main.m_lastStartWvln = 0
#End If

  ' Load and save product default configuration file values
  load_prod_default_file
  
End Sub

Private Sub load_prod_default_file()
  Dim errMsg As String

  ' Load product default configuration file
  Call unity_main.load_prod_file(PROD_DFLTS_CFG_FILE, False)
  
  ' Save product defaults
  On Error Resume Next
#If ABBFT Then
  ProdDfltData.delayMeasure = frm_collect.m_delayMeasure
  ProdDfltData.delayStart = frm_collect.m_delayStart
  ProdDfltData.endWavenumIndx = frm_collect.m_endWavenumIndx
  ProdDfltData.gainIndx = frm_collect.m_gainIndx
  ProdDfltData.numMeasures = frm_collect.m_numMeasures
  ProdDfltData.numSamples = frm_collect.m_numSamples
  ProdDfltData.resolutionIndx = frm_collect.m_resolutionIndx
  ProdDfltData.speedIndx = frm_collect.m_speedIndx
  ProdDfltData.startWavenumIndx = frm_collect.m_startWavenumIndx

  ProdDfltData.backFreq = frm_collect.m_backFreq
  ProdDfltData.boundValue = frm_collect.m_valueBound
  ProdDfltData.btype = frm_collect.m_bType
  ProdDfltData.clrManualName = frm_collect.m_clrManualName
  ProdDfltData.clrUserInputs = frm_collect.m_clrUserInputs
  ProdDfltData.dayCounter = frm_collect.m_dayCounter
  ProdDfltData.endWvln = frm_collect.m_smplEndWvln
  ProdDfltData.hideValCol = frm_collect.m_hideValCol
  ProdDfltData.makePred = frm_collect.m_makePred
  ProdDfltData.mdAlarm = frm_collect.m_alarmMD
  ProdDfltData.useMenuInput = frm_collect.m_useMIV
  ProdDfltData.namebase = frm_collect.m_nameBase
  ProdDfltData.nameCounter = frm_collect.m_nameCounter
  ProdDfltData.nameScanType = frm_collect.m_nameScanType
  ProdDfltData.ndAlarm = frm_collect.m_alarmND
  ProdDfltData.noOLVal = frm_collect.m_noOLVal
  ProdDfltData.noPredVal = frm_collect.m_noPredVal
  ProdDfltData.outlierLights = frm_collect.m_olFormat
  ProdDfltData.repsAvg = frm_collect.m_repsAvg
  ProdDfltData.rrAlarm = frm_collect.m_alarmRR
  ProdDfltData.saveIt = frm_collect.m_saveIt
  ProdDfltData.saveScanDir = frm_collect.m_saveDir
  ProdDfltData.saveCSVFile = frm_collect.m_saveCSVFile
  ProdDfltData.savePredFile = frm_collect.m_savePredFile
  ProdDfltData.savePredictions = frm_collect.m_savePredictions
  ProdDfltData.savePredictionsCSV = frm_collect.m_saveCSV
  ProdDfltData.savePredictionsDynRpt = frm_collect.m_saveDynRpt
  ProdDfltData.sendLimsOutput = frm_collect.m_doLIMS
  ProdDfltData.startWvln = frm_collect.m_smplStartWvln
  ProdDfltData.sType = frm_collect.m_sType
  ProdDfltData.valueAlarm = frm_collect.m_alarmProp
  ProdDfltData.writeTkt = frm_collect.m_writeTkt
#Else
  ProdDfltData.adapterType = frm_collect.m_adapterType
  ProdDfltData.backFreq = frm_collect.m_backFreq
  ProdDfltData.boundValue = frm_collect.m_valueBound
  ProdDfltData.btype = frm_collect.m_bType
  ProdDfltData.clrManualName = frm_collect.m_clrManualName
  ProdDfltData.clrUserInputs = frm_collect.m_clrUserInputs
  ProdDfltData.dayCounter = frm_collect.m_dayCounter
  ProdDfltData.endWvln = frm_collect.m_smplEndWvln
  ProdDfltData.extRefFileName = frm_collect.m_extRefFileName
  ProdDfltData.hideValCol = frm_collect.m_hideValCol
  ProdDfltData.makePred = frm_collect.m_makePred
  ProdDfltData.mdAlarm = frm_collect.m_alarmMD
  ProdDfltData.useMenuInput = frm_collect.m_useMIV
  ProdDfltData.multiCupType = frm_collect.m_multiCupType
  ProdDfltData.namebase = frm_collect.m_nameBase
  ProdDfltData.nameCounter = frm_collect.m_nameCounter
  ProdDfltData.nameScanType = frm_collect.m_nameScanType
  ProdDfltData.ndAlarm = frm_collect.m_alarmND
  ProdDfltData.noOLVal = frm_collect.m_noOLVal
  ProdDfltData.noPredVal = frm_collect.m_noPredVal
  ProdDfltData.numSmplScans = frm_collect.m_smplNScans
  ProdDfltData.olRefFileName = frm_collect.m_olRefFileName
  ProdDfltData.outlierLights = frm_collect.m_olFormat
  ProdDfltData.repsAvg = frm_collect.m_repsAvg
  ProdDfltData.rotateDir = frm_collect.m_rotateDir
  ProdDfltData.rotateIndexSteps = frm_collect.m_rotateIndexSteps
  ProdDfltData.rotateMoveMode = frm_collect.m_rotateMoveMode
  ProdDfltData.rotateSpeed = frm_collect.m_rotateSpeed
  ProdDfltData.rotateStepSteps = frm_collect.m_rotateStepSteps
  ProdDfltData.rrAlarm = frm_collect.m_alarmRR
  ProdDfltData.saveIt = frm_collect.m_saveIt
  ProdDfltData.saveScanDir = frm_collect.m_saveDir
  ProdDfltData.saveCSVFile = frm_collect.m_saveCSVFile
  ProdDfltData.savePredFile = frm_collect.m_savePredFile
  ProdDfltData.savePredictions = frm_collect.m_savePredictions
  ProdDfltData.savePredictionsCSV = frm_collect.m_saveCSV
  ProdDfltData.savePredictionsDynRpt = frm_collect.m_saveDynRpt
  ProdDfltData.saveReplicates = frm_collect.m_saveReps
  ProdDfltData.sendLimsOutput = frm_collect.m_doLIMS
  ProdDfltData.smplPPT = frm_collect.m_smplPPT
  ProdDfltData.startWvln = frm_collect.m_smplStartWvln
  ProdDfltData.sType = frm_collect.m_sType
  ProdDfltData.useExtRefTrayCfg = frm_collect.m_useExtRefTrayCfg
  ProdDfltData.valueAlarm = frm_collect.m_alarmProp
  ProdDfltData.writeTkt = frm_collect.m_writeTkt
#End If
End Sub

Private Sub init_spectrum_plot()
  Dim ii As Integer
  Dim pad As Double

  XYPlot1.AxisMinMaxPad = 3
  XYPlot1.FontSize = FontSizes.SMALL
  XYPlot1.MainTitle = " "
  XYPlot1.NumSubsets = 1
  XYPlot1.SubTitle = ""
  
#If ABBFT Then
  XYPlot1.XAxisLabel = MLSupport.GSS("XYPlot", "XAxisLabel2", "Wavenumber (1/cm)")
#Else
  XYPlot1.XAxisLabel = MLSupport.GSS("XYPlot", "XAxisLabel", "Wavelength (nm)")
#End If

  XYPlot1.YAxisLabel = MLSupport.GSS("XYPlot", "YAxisLabel3", "Absorbance")

#If ABBFT Then
  pad = (unity_main.m_mb3000.m_endWavenum - unity_main.m_mb3000.m_startWavenum) / 50
  Call XYPlot1.Initialize(unity_main.m_mb3000.m_startWavenum - pad, unity_main.m_mb3000.m_endWavenum + pad, unity_main.m_mb3000.m_waveNumIncr, _
                          True, False, True, False, LegendStyles.TWO_LINE, False)
#Else
  Call XYPlot1.Initialize(unity_main.m_minWvln, unity_main.m_maxWvln, MS11CfgData.wvlnIncr, _
                          True, False, True, False, LegendStyles.TWO_LINE, False)
#End If

  Call XYPlot1.MoveChart(0, 10, -10, 0)
  Call XYPlot1.SetupManualXAxis(5, 5)
End Sub

Private Sub load_unitymain()
  Dim optVal As Integer

  unity_main.os_value = 7
  
  On Error GoTo probwithopen
  ' load CSV report configuration
  Call frm_csvCfg.load_cfg(True)
  
  ' load global system configuration
  frm_Inst.Loadmyinstini
  
#If SSTAR Then
  ' load spectrum treatment configuration
  frm_spectTreatCfg.load_cfg
#End If
  
  ' load last application configuration
  Call unity_main.load_prod_file("", True)
  
#If ABBFT Then
  ' Start interferometer state machine
  unity_main.m_mb3000.start_interferometer_state
#End If

  setupspread
  unity_main.pw_open = False
  load_last50

#If ABBFT Then
  unity_main.lbl_backdate.Caption = MLSupport.GSS("unity_main", "lbl_backdate_int2", "Internal Reference has Expired")
#Else
  Select Case (unity_main.m_bType)
    Case "internal"
      ' Check if internal reference performed on demand
      If (unity_main.m_backFreq = REF_FREQ_ON_DEMAND) Then
        ' Check if reference timeout configured
        If (unity_main.m_intRefTimeout > 0) Then
          unity_main.lbl_backdate.Caption = MLSupport.GSS("unity_main", "lbl_backdate_int2", "Internal Reference has Expired")
        Else
          unity_main.lbl_backdate.Caption = MLSupport.GSS("unity_main", "lbl_backdate_int1", "Internal Reference Will Never Expire")
        End If
      Else
        unity_main.lbl_backdate.Caption = MLSupport.GSS("unity_main", "lbl_backdate_int3", "Internal Reference Performed Every Sample")
      End If
      
    Case "external"
      ' Check if reference timeout configured
      If (unity_main.m_extRefTimeoutSecs > 0) Then
        unity_main.lbl_backdate.Caption = MLSupport.GSS("unity_main", "lbl_backdate_ext2", "External Reference has Expired")
          
        ' Check if no request for external reference qualification scan
        If (unity_main.m_extRefPPTScan = False) Then
          unity_main.m_extRefAutoScan = True
        End If
      Else
        unity_main.lbl_backdate.Caption = MLSupport.GSS("unity_main", "lbl_backdate_ext1", "External Reference Will Never Expire")
      End If
      
    Case "file"    ' from offline file
      unity_main.lbl_backdate.Caption = MLSupport.GSS("unity_main", "lbl_backdate_ol1", "Offline Reference Will Never Expire")
  End Select
#End If

  On Error Resume Next
  unity_main.backdate = FileDateTime(REFERENCES_DIR & REFERENCE_SCAN_FILE & SPC_FILE_EXT)
  Exit Sub
  
probwithopen:
  optVal = CWrap.ShowMessageBoxW(MLSupport.GSS("unity_main", "errMsg5", "Problem loading system configuration files. Would you like to restore last saved configuration files?"), vbYesNo)

  If (optVal = vbYes) Then
    frm_backrestore.restore_inis
  End If
  
  unity_main.unloadallforms "unity_main"
  Unload Me
  End
End Sub

Private Sub setup_prod_default_data(dfltIni As Boolean)
  Dim adaptIndx As Integer

  unity_main.m_fileDevID = 0
  unity_main.m_fileVersion = INFOSTAR_VER
  unity_main.m_smplTable = 0
  unity_main.m_sysScanMode = &H7C0    ' set bits 4-9

  ' Check if going to load product default.ini file for first time
  If (dfltIni = True) And (unity_main.m_defltsLoaded = False) Then
    ' Setup product values before loading product .ini file
    ' in case of bad or missing file
#If ABBFT Then
    unity_main.current_product = "Default-FT"
  
    ProdDfltData.delayMeasure = 0
    frm_collect.m_delayMeasure = 0
    
    ProdDfltData.delayStart = 0
    frm_collect.m_delayStart = 0
    
    ProdDfltData.endWavenumIndx = 4148
    frm_collect.m_endWavenumIndx = 4148
    
    ProdDfltData.gainIndx = 0
    frm_collect.m_gainIndx = 0
    
    ProdDfltData.numMeasures = 1
    frm_collect.m_numMeasures = 1
    
    ProdDfltData.numSamples = 10
    frm_collect.m_numSamples = 10
    
    ProdDfltData.resolutionIndx = 1
    frm_collect.m_resolutionIndx = 1
    
    ProdDfltData.speedIndx = 5
    frm_collect.m_speedIndx = 5
    
    ProdDfltData.startWavenumIndx = 2074
    frm_collect.m_startWavenumIndx = 2074
    
    ProdDfltData.backFreq = REF_FREQ_ON_DEMAND
    frm_collect.m_backFreq = REF_FREQ_ON_DEMAND
    
    ProdDfltData.boundValue = 0
    frm_collect.m_valueBound = 0
  
    ProdDfltData.btype = "external"
    frm_collect.m_bType = "external"
  
    ProdDfltData.clrManualName = False
    frm_collect.m_clrManualName = False
    
    ProdDfltData.clrUserInputs = False
    frm_collect.m_clrUserInputs = False
    
    ProdDfltData.dayCounter = 1
    frm_collect.m_dayCounter = 1
    
    ProdDfltData.hideValCol = False
    frm_collect.m_hideValCol = False
    
    ProdDfltData.makePred = "yes"
    frm_collect.m_makePred = "yes"
    
    ProdDfltData.mdAlarm = 0
    frm_collect.m_alarmMD = 0
    
    ProdDfltData.useMenuInput = False
    frm_collect.m_useMIV = False
    
    ProdDfltData.namebase = "Sample"
    frm_collect.m_nameBase = "Sample"
    
    ProdDfltData.nameCounter = 0
    frm_collect.m_nameCounter = 0
    
    ProdDfltData.nameScanType = "Date"
    frm_collect.m_nameScanType = "Date"
    
    ProdDfltData.ndAlarm = 0
    frm_collect.m_alarmND = 0
    
    ProdDfltData.noOLVal = "0"
    frm_collect.m_noOLVal = "0"
    
    ProdDfltData.noPredVal = "NA"
    frm_collect.m_noPredVal = "NA"
    
    ProdDfltData.outlierLights = True
    frm_collect.m_olFormat = True
    
    ProdDfltData.repsAvg = 1
    frm_collect.m_repsAvg = 1
    
    ProdDfltData.rrAlarm = 0
    frm_collect.m_alarmRR = 0
    
    ProdDfltData.saveIt = "save"
    frm_collect.m_saveIt = "save"
    
    ProdDfltData.saveScanDir = SPECTRA_DFLT_DIR
    frm_collect.m_saveDir = SPECTRA_DFLT_DIR
    
    ProdDfltData.saveCSVFile = REPORTS_DIR & DFLT_CSV_FILE
    frm_collect.m_saveCSVFile = REPORTS_DIR & DFLT_CSV_FILE
    
    ProdDfltData.savePredFile = REPORTS_DIR & DFLT_REPORT_FILE
    frm_collect.m_savePredFile = REPORTS_DIR & DFLT_REPORT_FILE
    
    ProdDfltData.savePredictions = False
    frm_collect.m_savePredictions = True
    
    ProdDfltData.savePredictionsCSV = True
    frm_collect.m_saveCSV = True
    
    ProdDfltData.savePredictionsDynRpt = False
    frm_collect.m_saveDynRpt = False
    
    ProdDfltData.sendLimsOutput = 0
    frm_collect.m_doLIMS = 0
    
    ProdDfltData.sType = "abs"
    frm_collect.m_sType = "abs"
    
    ProdDfltData.useExtRefTrayCfg = False
    frm_collect.m_useExtRefTrayCfg = False
    
    ProdDfltData.valueAlarm = 0
    frm_collect.m_alarmProp = 0
    
    ProdDfltData.writeTkt = 0
    frm_collect.m_writeTkt = 0
#Else
    unity_main.current_product = "Default"
  
    Select Case (MS11CfgData.devID)
      Case DTID_DRAWER0              ' SS2200/SS2400 standard drawer system
        ' Default to adapter number 2
        ProdDfltData.adapterType = MS11AdapterInfo(2).cfgName
        frm_collect.m_adapterType = MS11AdapterInfo(2).cfgName
      Case DTID_TOPWIND0             ' Top window w/out internal reflectance
        ' Default to adapter number 1
        ProdDfltData.adapterType = MS11AdapterInfo(1).cfgName
        frm_collect.m_adapterType = MS11AdapterInfo(1).cfgName
      Case DTID_DRAWER1              ' SS2200/SS2400 drawer w/out stepper system
        ' Default to adapter number 1
        ProdDfltData.adapterType = MS11AdapterInfo(1).cfgName
        frm_collect.m_adapterType = MS11AdapterInfo(1).cfgName
      Case DTID_TOPWIND1             ' Top window with internal reflectance
        ' Default to adapter number 1
        ProdDfltData.adapterType = MS11AdapterInfo(1).cfgName
        frm_collect.m_adapterType = MS11AdapterInfo(1).cfgName
    End Select
  
    ' Check if instrument's default sample tray not defined
    If (ProdDfltData.adapterType = "") Then
      ' Find first available adapter type
      For adaptIndx = 1 To MAX_ADAPTER_TYPES
        If (MS11AdapterInfo(adaptIndx).cfgName <> "") Then
          ProdDfltData.adapterType = MS11AdapterInfo(adaptIndx).cfgName
          frm_collect.m_adapterType = MS11AdapterInfo(adaptIndx).cfgName
          Exit For
        End If
      Next adaptIndx
    Else
      ' Get adapter type index
      For adaptIndx = 1 To MAX_ADAPTER_TYPES
        If (MS11AdapterInfo(adaptIndx).cfgName = ProdDfltData.adapterType) Then
          Exit For
        End If
      Next adaptIndx
    End If

    ProdDfltData.backFreq = REF_FREQ_ON_DEMAND
    frm_collect.m_backFreq = REF_FREQ_ON_DEMAND
    
    ProdDfltData.boundValue = 0
    frm_collect.m_valueBound = 0
  
    ' Check if Top window system w/out internal reflectance
    If (MS11CfgData.devID = DTID_TOPWIND0) Then
      ProdDfltData.btype = "external"
      frm_collect.m_bType = "external"
    Else
      ProdDfltData.btype = "internal"
      frm_collect.m_bType = "internal"
    End If
  
    ProdDfltData.clrManualName = False
    frm_collect.m_clrManualName = False
    
    ProdDfltData.clrUserInputs = False
    frm_collect.m_clrUserInputs = False
    
    ProdDfltData.dayCounter = 1
    frm_collect.m_dayCounter = 1
    
    ProdDfltData.endWvln = MS11DfltScanCfgData.endWvln
    frm_collect.m_smplEndWvln = MS11DfltScanCfgData.endWvln
    
    ProdDfltData.extRefFileName = ""
    frm_collect.m_extRefFileName = ""
    
    ProdDfltData.hideValCol = False
    frm_collect.m_hideValCol = False
    
    ProdDfltData.makePred = "yes"
    frm_collect.m_makePred = "yes"
    
    ProdDfltData.mdAlarm = 0
    frm_collect.m_alarmMD = 0
    
    ProdDfltData.useMenuInput = False
    frm_collect.m_useMIV = False

    ProdDfltData.multiCupType = CFG_NONE_MCT
    frm_collect.m_multiCupType = CFG_NONE_MCT
    
    ProdDfltData.namebase = "Sample"
    frm_collect.m_nameBase = "Sample"
    
    ProdDfltData.nameCounter = 0
    frm_collect.m_nameCounter = 0
    
    ProdDfltData.nameScanType = "Date"
    frm_collect.m_nameScanType = "Date"
    
    ProdDfltData.ndAlarm = 0
    frm_collect.m_alarmND = 0
    
    ProdDfltData.noOLVal = "0"
    frm_collect.m_noOLVal = "0"
    
    ProdDfltData.noPredVal = "NA"
    frm_collect.m_noPredVal = "NA"
    
    ProdDfltData.numSmplScans = MS11DfltScanCfgData.nScans4Smpl
    frm_collect.m_smplNScans = MS11DfltScanCfgData.nScans4Smpl
    
    ProdDfltData.olRefFileName = ""
    frm_collect.m_olRefFileName = ""
    
    ProdDfltData.outlierLights = True
    frm_collect.m_olFormat = True
    
    ProdDfltData.repsAvg = 1
    frm_collect.m_repsAvg = 1
    
    ProdDfltData.rotateDir = TRD_NONE
    frm_collect.m_rotateDir = TRD_NONE
    
    ProdDfltData.rotateIndexSteps = MS11DfltTrayCfgData(MS11AdapterInfo(adaptIndx).trayNum).stps4IX
    frm_collect.m_rotateIndexSteps = ProdDfltData.rotateIndexSteps
    
    ProdDfltData.rotateMoveMode = TRM_NONE
    frm_collect.m_rotateMoveMode = TRM_NONE
    
    ProdDfltData.rotateSpeed = MS11DfltTrayCfgData(MS11AdapterInfo(adaptIndx).trayNum).velCont
    frm_collect.m_rotateSpeed = ProdDfltData.rotateSpeed
    
    ProdDfltData.rotateStepSteps = MS11DfltTrayCfgData(MS11AdapterInfo(adaptIndx).trayNum).stps4scn
    frm_collect.m_rotateStepSteps = ProdDfltData.rotateStepSteps
    
    ProdDfltData.rrAlarm = 0
    frm_collect.m_alarmRR = 0
    
    ProdDfltData.saveIt = "save"
    frm_collect.m_saveIt = "save"
    
    ProdDfltData.saveScanDir = SPECTRA_DFLT_DIR
    frm_collect.m_saveDir = SPECTRA_DFLT_DIR
    
    ProdDfltData.saveCSVFile = REPORTS_DIR & DFLT_CSV_FILE
    frm_collect.m_saveCSVFile = REPORTS_DIR & DFLT_CSV_FILE
    
    ProdDfltData.savePredFile = REPORTS_DIR & DFLT_REPORT_FILE
    frm_collect.m_savePredFile = REPORTS_DIR & DFLT_REPORT_FILE
    
    ProdDfltData.savePredictions = False
    frm_collect.m_savePredictions = True
    
    ProdDfltData.savePredictionsCSV = True
    frm_collect.m_saveCSV = True
    
    ProdDfltData.savePredictionsDynRpt = False
    frm_collect.m_saveDynRpt = False
    
    ProdDfltData.saveReplicates = True
    frm_collect.m_saveReps = True
    
    ProdDfltData.sendLimsOutput = 0
    frm_collect.m_doLIMS = 0
    
    ProdDfltData.smplPPT = MS11DfltScanCfgData.smpl4PPT
    frm_collect.m_smplPPT = MS11DfltScanCfgData.smpl4PPT
    
    ProdDfltData.startWvln = MS11DfltScanCfgData.startWvln
    frm_collect.m_smplStartWvln = MS11DfltScanCfgData.startWvln
    
    ProdDfltData.sType = "abs"
    frm_collect.m_sType = "abs"
    
    ProdDfltData.useExtRefTrayCfg = False
    frm_collect.m_useExtRefTrayCfg = False
    
    ProdDfltData.valueAlarm = 0
    frm_collect.m_alarmProp = 0
    
    ProdDfltData.writeTkt = 0
    frm_collect.m_writeTkt = 0
#End If
  Else
    ' Initialize values before loading product .ini file
#If ABBFT Then
    frm_collect.m_delayMeasure = ProdDfltData.delayMeasure
    frm_collect.m_delayStart = ProdDfltData.delayStart
    frm_collect.m_endWavenumIndx = ProdDfltData.endWavenumIndx
    frm_collect.m_gainIndx = ProdDfltData.gainIndx
    frm_collect.m_numMeasures = ProdDfltData.numMeasures
    frm_collect.m_numSamples = ProdDfltData.numSamples
    frm_collect.m_resolutionIndx = ProdDfltData.resolutionIndx
    frm_collect.m_speedIndx = ProdDfltData.speedIndx
    frm_collect.m_startWavenumIndx = ProdDfltData.startWavenumIndx

    frm_collect.m_backFreq = ProdDfltData.backFreq
    frm_collect.m_valueBound = ProdDfltData.boundValue
    frm_collect.m_bType = ProdDfltData.btype
    frm_collect.m_clrManualName = ProdDfltData.clrManualName
    frm_collect.m_clrUserInputs = ProdDfltData.clrUserInputs
    frm_collect.m_dayCounter = ProdDfltData.dayCounter
    frm_collect.m_hideValCol = ProdDfltData.hideValCol
    frm_collect.m_makePred = ProdDfltData.makePred
    frm_collect.m_alarmMD = ProdDfltData.mdAlarm
    frm_collect.m_useMIV = ProdDfltData.useMenuInput
    frm_collect.m_nameBase = ProdDfltData.namebase
    frm_collect.m_nameCounter = ProdDfltData.nameCounter
    frm_collect.m_nameScanType = ProdDfltData.nameScanType
    frm_collect.m_alarmND = ProdDfltData.ndAlarm
    frm_collect.m_noOLVal = ProdDfltData.noOLVal
    frm_collect.m_noPredVal = ProdDfltData.noPredVal
    frm_collect.m_olFormat = ProdDfltData.outlierLights
    frm_collect.m_repsAvg = ProdDfltData.repsAvg
    frm_collect.m_alarmRR = ProdDfltData.rrAlarm
    frm_collect.m_saveIt = ProdDfltData.saveIt
    frm_collect.m_saveDir = ProdDfltData.saveScanDir
    frm_collect.m_saveCSVFile = ProdDfltData.saveCSVFile
    frm_collect.m_savePredFile = ProdDfltData.savePredFile
    frm_collect.m_savePredictions = ProdDfltData.savePredictions
    frm_collect.m_saveCSV = ProdDfltData.savePredictionsCSV
    frm_collect.m_saveDynRpt = ProdDfltData.savePredictionsDynRpt
    frm_collect.m_saveReps = ProdDfltData.saveReplicates
    frm_collect.m_doLIMS = ProdDfltData.sendLimsOutput
    frm_collect.m_sType = ProdDfltData.sType
    frm_collect.m_alarmProp = ProdDfltData.valueAlarm
    frm_collect.m_writeTkt = ProdDfltData.writeTkt
#Else
    frm_collect.m_adapterType = ProdDfltData.adapterType
    frm_collect.m_backFreq = ProdDfltData.backFreq
    frm_collect.m_valueBound = ProdDfltData.boundValue
    frm_collect.m_bType = ProdDfltData.btype
    frm_collect.m_clrManualName = ProdDfltData.clrManualName
    frm_collect.m_clrUserInputs = ProdDfltData.clrUserInputs
    frm_collect.m_dayCounter = ProdDfltData.dayCounter
    frm_collect.m_smplEndWvln = ProdDfltData.endWvln
    frm_collect.m_extRefFileName = ProdDfltData.extRefFileName
    frm_collect.m_hideValCol = ProdDfltData.hideValCol
    frm_collect.m_makePred = ProdDfltData.makePred
    frm_collect.m_alarmMD = ProdDfltData.mdAlarm
    frm_collect.m_useMIV = ProdDfltData.useMenuInput
    frm_collect.m_multiCupType = ProdDfltData.multiCupType
    frm_collect.m_nameBase = ProdDfltData.namebase
    frm_collect.m_nameCounter = ProdDfltData.nameCounter
    frm_collect.m_nameScanType = ProdDfltData.nameScanType
    frm_collect.m_alarmND = ProdDfltData.ndAlarm
    frm_collect.m_noOLVal = ProdDfltData.noOLVal
    frm_collect.m_noPredVal = ProdDfltData.noPredVal
    frm_collect.m_smplNScans = ProdDfltData.numSmplScans
    frm_collect.m_olRefFileName = ProdDfltData.olRefFileName
    frm_collect.m_olFormat = ProdDfltData.outlierLights
    frm_collect.m_repsAvg = ProdDfltData.repsAvg
    frm_collect.m_rotateDir = ProdDfltData.rotateDir
    frm_collect.m_rotateIndexSteps = ProdDfltData.rotateIndexSteps
    frm_collect.m_rotateMoveMode = ProdDfltData.rotateMoveMode
    frm_collect.m_rotateSpeed = ProdDfltData.rotateSpeed
    frm_collect.m_rotateStepSteps = ProdDfltData.rotateStepSteps
    frm_collect.m_alarmRR = ProdDfltData.rrAlarm
    frm_collect.m_saveIt = ProdDfltData.saveIt
    frm_collect.m_saveDir = ProdDfltData.saveScanDir
    frm_collect.m_saveCSVFile = ProdDfltData.saveCSVFile
    frm_collect.m_savePredFile = ProdDfltData.savePredFile
    frm_collect.m_savePredictions = ProdDfltData.savePredictions
    frm_collect.m_saveCSV = ProdDfltData.savePredictionsCSV
    frm_collect.m_saveDynRpt = ProdDfltData.savePredictionsDynRpt
    frm_collect.m_saveReps = ProdDfltData.saveReplicates
    frm_collect.m_doLIMS = ProdDfltData.sendLimsOutput
    frm_collect.m_smplPPT = ProdDfltData.smplPPT
    frm_collect.m_smplStartWvln = ProdDfltData.startWvln
    frm_collect.m_sType = ProdDfltData.sType
    frm_collect.m_useExtRefTrayCfg = ProdDfltData.useExtRefTrayCfg
    frm_collect.m_alarmProp = ProdDfltData.valueAlarm
    frm_collect.m_writeTkt = ProdDfltData.writeTkt
#End If
  End If
End Sub

Private Sub perform_ref_scan()
  Dim uniMsg As String

  If (lblProd1.Caption = "") Then
    uniMsg = MLSupport.GSS("OperStatus", "status4", "Please Select a Product")
    lbl_opStatus.Caption = uniMsg
    CWrap.ShowMessageBoxW uniMsg, vbCritical
    Call unity_main.kill_loop(LOG_DBG_LEVEL1, True, "No product selected")
    cmd_sample.enabled = True
    If cmd_sample.Visible And cmd_sample.enabled Then
             'cmd_sample.SetFocus
    End If
    Exit Sub
  End If
  
  tmr_all.enabled = False
  
#If ABBFT Then
  unity_main.m_mb3000.m_newRefReq = True
#Else
  unity_main.m_scanTmrState = STS_SETUP
  tmr_ref.enabled = True
#End If

  frm_status.lbl_status.Caption = lbl_opStatus.Caption
  frm_status.lbl_statusCmd.Caption = ""
  frm_status.Show 1
End Sub

Private Sub perform_smpl_scan(scanReqType As Boolean)
  Dim numPts As Integer
  Dim scanDataType As SCAN_DATA_TYPES
  Dim opStatus As String
  Dim errMsg As String
  Dim uniMsg As String
 
  scanReqType = False
  
  ' Check if a product has been selected
  If (lblProd1.Caption = "") Then
    uniMsg = MLSupport.GSS("OperStatus", "status4", "Please Select a Product")
    lbl_opStatus.Caption = uniMsg
    CWrap.ShowMessageBoxW uniMsg, vbCritical
    cmd_sample.enabled = True
    If cmd_sample.Visible And cmd_sample.enabled Then
             'cmd_sample.SetFocus
    End If
    Exit Sub
  End If
  
  unity_main.tmr_all.enabled = False
  unity_main.m_smplRepacks = unity_main.m_repsAvg
  
  ' Check if first scan for repack
  If (unity_main.repcounter = 0) Then
#If SSTAR Then
    If (unity_main.m_smplRepacks > 1) Then
      numPts = (unity_main.m_smplEndWvln - unity_main.m_smplStartWvln) / MS11CfgData.wvlnIncr
      ReDim ProdAvgAbsYVals(numPts)
      
      ' Check if to treat spectrum data
      If (unity_main.m_enableTreatment = True) Then
        ReDim ProdTreatAvgAbsYVals(numPts)
      End If
    End If
#End If

    ' Check if current product has PRD model type
    If (unity_main.m_prdModelType = True) Then
      ' Check if .NET Framework 2.0 is not installed
      If (unity_main.m_netFWInstalled = False) Then
        errMsg = "Problem with product trying to save spectra data to UCal SVF file without MS .NET Framework 2.0 installed"
        unity_main.errorstring = errMsg
        unity_main.write_error (LOG_DBG_LEVEL1)
        uniMsg = MLSupport.GSS("unity_main", "errMsg7", "Problem with product trying to save spectra data to UCal SVF file without MS .NET Framework 2.0 installed")
        CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
        
        ' Restart auto-loop timer
        unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status10", "Ready for New Sample")
        tmr_all.enabled = True
        Exit Sub
      End If
    End If

    unity_main.prepscan
    unity_main.clearpredtable
    
    unity_main.img_report.Visible = False
    unity_main.img_ticket.Visible = False
    unity_main.img_batchRpt.Visible = False
    unity_main.m_batchRptFile = ""
    
    If (frm_collect.get_samp_name = False) Then
      Exit Sub
    End If
    
    unity_main.cmd_repacks.Visible = False
    
    ' Check if to perform internal reference before sample
    If (unity_main.m_backFreq = REF_FREQ_ALL_SMPLS) Then
      ' Save sample info
      scanDataType = unity_main.m_scanDataType
      opStatus = lbl_opStatus.Caption

      ' Setup reference info
      unity_main.m_scanDataType = SDT_PRODINTREF
      lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status12", "Manual Internal Reference Scan")
      frm_status.m_scanAborted = False
      perform_ref_scan
      
      ' Wait until reference scan completed
      While (tmr_ref.enabled = True)
        DoEvents
      Wend
    
      ' Check if reference scan aborted
      If (frm_status.m_scanAborted = True) Then
        frm_collect.roll_back_name_ctr
        unity_main.txtsamplename.Text = ""
        unity_main.txtsampcomment.Text = ""
        Exit Sub
      End If
      
      ' Restore sample info
      unity_main.m_scanDataType = scanDataType
      lbl_opStatus.Caption = opStatus
    End If
  Else
    ' Clear the debug status list box
    frm_dbug.lst_status.Clear
  End If
  
  
#If ABBFT Then
  unity_main.m_mb3000.m_newSmplReq = True
#Else
  unity_main.m_scanTmrState = STS_SETUP
  tmr_sample.enabled = True
#End If
  
  If (unity_main.m_smplRepacks > 1) Then
    lbl_opStatus.Caption = lbl_opStatus.Caption & " - " & MLSupport.GGS_Params("unity_main.statMsg30", "Repack %1 of %2", CStr(unity_main.repcounter), CStr(unity_main.m_smplRepacks))
  End If
  
  frm_status.lbl_status.Caption = unity_main.txtsamplename.Text & vbCrLf & lbl_opStatus.Caption
  frm_status.lbl_statusCmd.Caption = ""
  frm_status.Show 1
  Exit Sub
End Sub

Private Sub process_prod_file_vars(fileName As String, ByVal uniFile As clsUniFile, ByVal fEncoding As eFileEncoding, ByRef lineCnt As Integer, prodSelect As Boolean)
  Dim strlen, offset As Integer
  Dim tmpStrg As String
  Dim spcFilename As String
  Dim numPts As Long
  Dim waveNumIncr As Double
  Dim wavenumIndx As Long
  Dim wavenum As Double
    
#If ABBFT Then
  ' Interferometer scanning parameters
  frm_collect.lst_resolution.ListIndex = frm_collect.m_resolutionIndx
  frm_collect.lst_speed.ListIndex = frm_collect.m_speedIndx
  frm_collect.lst_gain.ListIndex = frm_collect.m_gainIndx
  
  ' Interferometer CoAddition parameters
  frm_collect.m_numMeasures = 1  ' Only support 1 measurement for now
  frm_collect.txt_numMeasures.Text = frm_collect.m_numMeasures
  frm_collect.numInc_numSamples.Text = frm_collect.m_numSamples
  frm_collect.txt_delayStart.Text = frm_collect.m_delayStart
  frm_collect.txt_delayMeasure.Text = frm_collect.m_delayMeasure
    
  ' Interferometer sample/prediction parameters
  Call frm_collect.calc_wavenum_incr(frm_collect.lst_resolution.List(frm_collect.lst_resolution.ListIndex), numPts, waveNumIncr)
  frm_collect.m_smplNumPts = numPts
  frm_collect.m_waveNumIncr = waveNumIncr
  frm_collect.txt_smplNumPts.Text = numPts
  frm_collect.txt_waveNumIncr.Text = waveNumIncr
  
  wavenumIndx = frm_collect.m_startWavenumIndx
  
  If (frm_collect.calc_wavenum(wavenumIndx, wavenum) = False) Then
    frm_collect.m_startWavenumIndx = wavenumIndx
  End If
  
  frm_collect.m_startWavenum = wavenum
  
  frm_collect.txt_startWavenumIndx.Text = frm_collect.m_startWavenumIndx
  frm_collect.txt_startWavenum.Text = frm_collect.m_startWavenum
  
  wavenumIndx = frm_collect.m_endWavenumIndx
  
  If (frm_collect.calc_wavenum(wavenumIndx, wavenum) = False) Then
    frm_collect.m_endWavenumIndx = wavenumIndx
  End If
  
  frm_collect.m_endWavenum = wavenum
  
  frm_collect.txt_endWavenumIndx.Text = frm_collect.m_endWavenumIndx
  frm_collect.txt_endWavenum.Text = frm_collect.m_endWavenum
#End If
    
  ' Setup limiting report values selection
PROCESS_BOUNDVALUE:
  Select Case (frm_collect.m_valueBound)
    Case 0          ' none
      frm_collect.opt_boundno.Value = True
    Case 1          ' low
      frm_collect.opt_boundl.Value = True
    Case 2          ' high
      frm_collect.opt_boundh.Value = True
    Case 3          ' high & low
      frm_collect.opt_boundhl.Value = True
    Case Else       ' invalid value
      unity_main.errorstring = (fileName & " had incompatible value. Bound_Values was " & frm_collect.m_valueBound & "; updated to " & ProdDfltData.boundValue)
      unity_main.write_error
      frm_collect.m_valueBound = ProdDfltData.boundValue
      unity_main.m_badIniVal = True
      GoTo PROCESS_BOUNDVALUE
  End Select
  
#If SSTAR Then
  frm_collect.build_ref_name_list "external"
  frm_collect.build_ref_name_list "file"
#End If

  ' Setup reference type selection
PROCESS_BTYPE:
  Select Case (frm_collect.m_bType)
#If ABBFT Then
    Case "external"
      frm_collect.m_ignoreEvent = True
      frm_collect.optbgexternal.Value = True
      frm_collect.combo_extRefFileName.Visible = False
      frm_collect.combo_olRefFileName.Visible = False
      frm_collect.chk_useExtRefTrayCfg.Visible = False
      frm_collect.opt_backall.enabled = False
      frm_collect.m_ignoreEvent = False
#Else
    Case "internal"
      If (MS11CfgData.devID = DTID_TOPWIND0) Then GoTo BAD_BTYPE
      
      frm_collect.m_ignoreEvent = True
      frm_collect.optbginternal.Value = True
      frm_collect.combo_extRefFileName.Visible = False
      frm_collect.combo_olRefFileName.Visible = False
      frm_collect.chk_useExtRefTrayCfg.Visible = False
      frm_collect.chk_useExtRefTrayCfg.Value = 0
      frm_collect.opt_backall.enabled = True
      frm_collect.m_ignoreEvent = False
    Case "external"
      frm_collect.m_ignoreEvent = True
      frm_collect.optbgexternal.Value = True
      frm_collect.combo_extRefFileName.Visible = True
      frm_collect.combo_olRefFileName.Visible = False
      frm_collect.chk_useExtRefTrayCfg.Visible = True
      
      If (frm_collect.m_useExtRefTrayCfg = True) Then
        frm_collect.chk_useExtRefTrayCfg.Value = 1
      Else
        frm_collect.chk_useExtRefTrayCfg.Value = 0
      End If
      
      frm_collect.opt_backall.enabled = False
      frm_collect.m_ignoreEvent = False
    Case "file"    ' from offline file
      frm_collect.m_ignoreEvent = True
      frm_collect.optbgfile.Value = True
      frm_collect.combo_extRefFileName.Visible = False
      frm_collect.combo_olRefFileName.Visible = True
      frm_collect.chk_useExtRefTrayCfg.Visible = False
      frm_collect.chk_useExtRefTrayCfg.Value = 0
      frm_collect.opt_backall.enabled = False
      frm_collect.m_ignoreEvent = False
#End If
    Case Else       ' invalid value
BAD_BTYPE:
      unity_main.errorstring = (fileName & " had incompatible value. BType was " & frm_collect.m_bType & "; updated to " & ProdDfltData.btype)
      unity_main.write_error
      frm_collect.m_bType = ProdDfltData.btype
      unity_main.m_badIniVal = True
      GoTo PROCESS_BTYPE
  End Select

  ' Setup reference frequency selection
PROCESS_BACKFREQ:
  Select Case (frm_collect.m_backFreq)
#If ABBFT Then
    Case REF_FREQ_ON_DEMAND           ' on demand
      frm_collect.opt_backdemand.Value = True
#Else
    Case REF_FREQ_ON_DEMAND           ' on demand
      frm_collect.opt_backdemand.Value = True
    Case REF_FREQ_ALL_SMPLS           ' every sample
      If (frm_collect.m_bType = "internal") Then
        frm_collect.opt_backall.Value = True
      Else
        GoTo INV_BACKFREQ
      End If
#End If
    Case Else                         ' invalid value
INV_BACKFREQ:
      unity_main.errorstring = (fileName & " had incompatible value. Background_Frequency was " & frm_collect.m_backFreq & "; updated to " & ProdDfltData.backFreq)
      unity_main.write_error
      frm_collect.m_backFreq = ProdDfltData.backFreq
      unity_main.m_badIniVal = True
      GoTo PROCESS_BACKFREQ
  End Select

  ' Setup clear manual name entry
  If (frm_collect.m_clrManualName = True) Then
    frm_collect.chk_clrManualName = 1
  Else
    frm_collect.chk_clrManualName = 0
  End If
  
  ' Setup clear user inputs entry
  If (frm_collect.m_clrUserInputs = True) Then
    frm_collect.chk_clrUserInputs = 1
  Else
    frm_collect.chk_clrUserInputs = 0
  End If
  
  ' Check for invalid day counter
  If (frm_collect.m_dayCounter < frm_collect.numInc_dateCounter.Min) Or (frm_collect.m_dayCounter > frm_collect.numInc_dateCounter.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. Day_Counter was " & frm_collect.m_dayCounter & "; updated to " & ProdDfltData.dayCounter)
    unity_main.write_error
    frm_collect.m_dayCounter = ProdDfltData.dayCounter
    unity_main.m_badIniVal = True
  End If
  
  ' Setup day counter
  frm_collect.numInc_dateCounter.Text = frm_collect.m_dayCounter
  
  ' Setup external & offline reference files
  frm_collect.m_ignoreEvent = True
  frm_collect.combo_extRefFileName.Text = frm_collect.m_extRefFileName
  frm_collect.combo_olRefFileName.Text = frm_collect.m_olRefFileName
  frm_collect.m_ignoreEvent = False
  
  ' Setup hide prediction value column selection
  If (frm_collect.m_hideValCol = True) Then
    frm_collect.chk_hideValCol.Value = 1
  Else
    frm_collect.chk_hideValCol.Value = 0
  End If

  ' Setup make prediction selections
PROCESS_MAKEPRED:
  Select Case (LCase(frm_collect.m_makePred))
    Case "yes"
      frm_collect.optpredyes.Value = True
    Case "no"
      frm_collect.optpredno.Value = True
    Case Else       ' invalid value
      unity_main.errorstring = (fileName & " had incompatible value. MakePred was " & frm_collect.m_makePred & "; updated to " & ProdDfltData.makePred)
      unity_main.write_error
      frm_collect.m_makePred = ProdDfltData.makePred
      unity_main.m_badIniVal = True
      GoTo PROCESS_MAKEPRED
  End Select

  ' Check for invalid M-Dist alarm value
  If (frm_collect.m_alarmMD < 0) Or (frm_collect.m_alarmMD > 1) Then
    unity_main.errorstring = (fileName & " had incompatible value. MD_Alarm was " & frm_collect.m_alarmMD & "; updated to " & ProdDfltData.mdAlarm)
    unity_main.write_error
    frm_collect.m_alarmMD = ProdDfltData.mdAlarm
    unity_main.m_badIniVal = True
  End If
  
  ' Setup use input button/list selection
  If (frm_collect.m_useMIV = True) Then
    frm_collect.chk_userInputs.Value = 1
    frm_buttoncfg.loadit True
    frm_buttoncfg.loadbuttonform
    frm_buttoncfg.loadbuttonconfig True
  Else
    frm_collect.chk_userInputs.Value = 0
    frm_buttoncfg.clear_form
  End If
  
  ' Setup M-Dist alarm selection
  frmedmod.chk_md.Value = frm_collect.m_alarmMD
  
  ' Setup sample file base name
  frm_collect.txtsampnamebase.Text = frm_collect.m_nameBase
  
  ' Check for invalid name counter
  If (frm_collect.m_nameCounter < frm_collect.numInc_nameCounter.Min) Or (frm_collect.m_nameCounter > frm_collect.numInc_nameCounter.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. NameCounter was " & frm_collect.m_nameCounter & "; updated to " & ProdDfltData.nameCounter)
    unity_main.write_error
    frm_collect.m_nameCounter = ProdDfltData.nameCounter
    unity_main.m_badIniVal = True
  End If
  
  ' Setup name counter
  frm_collect.numInc_nameCounter.Text = frm_collect.m_nameCounter

  ' Setup sample file naming convention selection
PROCESS_NAMESCANTYPE:
  Select Case (frm_collect.m_nameScanType)
    Case "Manual"
      frm_collect.txt_scanname.Text = "Manual Entry"
      frm_collect.optnamemanual.Value = True
      frm_collect.m_sNameMode = 1
      frm_collect.chk_clrManualName.enabled = True
    Case "Counter"
      frm_collect.txt_scanname.Text = frm_collect.txtsampnamebase.Text & frm_collect.numInc_nameCounter.Text
      frm_collect.optnamecounter.Value = True
      frm_collect.m_sNameMode = 3
      frm_collect.chk_clrManualName.enabled = False
    Case "Date"
      Dim dateStrg As String
      Call frm_collect.rebuild_date(Date, dateStrg)
      frm_collect.txt_scanname.Text = dateStrg & "_" & frm_collect.numInc_dateCounter.Text
      frm_collect.opt_namedate.Value = True
      frm_collect.m_sNameMode = 4
      frm_collect.chk_clrManualName.enabled = False
    Case Else       ' invalid value
      unity_main.errorstring = (fileName & " had incompatible value. NameScanType was " & frm_collect.m_nameScanType & "; updated to " & ProdDfltData.nameScanType)
      unity_main.write_error
      frm_collect.m_nameScanType = ProdDfltData.nameScanType
      unity_main.m_badIniVal = True
      GoTo PROCESS_NAMESCANTYPE
  End Select

  ' Check for invalid neighborhood distance alarm value
  If (frm_collect.m_alarmND < 0) Or (frm_collect.m_alarmND > 1) Then
    unity_main.errorstring = (fileName & " had incompatible value. ND_Alarm was " & frm_collect.m_alarmND & "; updated to " & ProdDfltData.ndAlarm)
    unity_main.write_error
    frm_collect.m_alarmND = ProdDfltData.ndAlarm
    unity_main.m_badIniVal = True
  End If
  
  ' Setup no outlier reported value
  frm_collect.txt_noOLVal.Text = frm_collect.m_noOLVal
  
  ' Setup no prediction reported value
  frm_collect.txt_noPredVal.Text = frm_collect.m_noPredVal
  
#If SSTAR Then
  ' Check for invalid number of scans
  If (frm_collect.m_smplNScans < frm_collect.numInc_smplNScans.Min) Or (frm_collect.m_smplNScans > frm_collect.numInc_smplNScans.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. NScansS was " & frm_collect.m_smplNScans & "; updated to " & ProdDfltData.numSmplScans)
    unity_main.write_error
    frm_collect.m_smplNScans = ProdDfltData.numSmplScans
    unity_main.m_badIniVal = True
  End If
  
  ' Setup number of scans
  frm_collect.numInc_smplNScans.Text = frm_collect.m_smplNScans
#End If

  ' Setup neighborhood distance alarm selection
  frmedmod.chk_nd.Value = frm_collect.m_alarmND
    
  ' Setup outlier display format selection
  If (frm_collect.m_olFormat = True) Then
    frm_collect.opt_ollts.Value = True      ' display pass/warn/fail lights
  Else
    frm_collect.opt_olval.Value = True      ' display values
  End If
  
#If SSTAR Then
  ' Check for invalid number of scan repacks
  If (frm_collect.m_repsAvg < frm_collect.numInc_numRepacks.Min) Or (frm_collect.m_repsAvg > frm_collect.numInc_numRepacks.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. RepsAvg was " & frm_collect.m_repsAvg & "; updated to " & ProdDfltData.repsAvg)
    unity_main.write_error
    frm_collect.m_repsAvg = ProdDfltData.repsAvg
    unity_main.m_badIniVal = True
  End If

  ' Setup number of scan repacks
  frm_collect.numInc_numRepacks.Text = frm_collect.m_repsAvg
#End If

  ' Check for invalid residual alarm value
  If (frm_collect.m_alarmRR < 0) Or (frm_collect.m_alarmRR > 1) Then
    unity_main.errorstring = (fileName & " had incompatible value. RR_Alarm was " & frm_collect.m_alarmRR & "; updated to " & ProdDfltData.rrAlarm)
    unity_main.write_error
    frm_collect.m_alarmRR = ProdDfltData.rrAlarm
    unity_main.m_badIniVal = True
  End If
  
  ' Setup residual alarm selection
  frmedmod.chk_rr.Value = frm_collect.m_alarmRR

  ' Setup save sample file selection
PROCESS_SAVEIT:
  Select Case (LCase(frm_collect.m_saveIt))
    Case "save"
      frm_collect.optsavescanyes.Value = True
    Case "nosave"
      frm_collect.optsavescanno.Value = True
    Case Else       ' invalid value
      unity_main.errorstring = (fileName & " had incompatible value. SaveIt was " & frm_collect.m_saveIt & "; updated to " & ProdDfltData.saveIt)
      unity_main.write_error
      frm_collect.m_saveIt = ProdDfltData.saveIt
      unity_main.m_badIniVal = True
      GoTo PROCESS_SAVEIT
  End Select
       
  ' Confirm path contains '\' instead of '/'
  tmpStrg = frm_collect.m_saveDir
  check_filepathname_delimiters tmpStrg
  frm_collect.m_saveDir = tmpStrg
       
  ' Check if save dir is proper for this software version
  If (InStr(1, frm_collect.m_saveDir, UNITY_DIR) = 1) Then
    If (InStr(1, frm_collect.m_saveDir, SPECTRA_DIR) = 0) Then
      strlen = Len(frm_collect.m_saveDir)
      offset = InStrRev(frm_collect.m_saveDir, "\", strlen - 1)
      tmpStrg = (SPECTRA_DIR & Right(frm_collect.m_saveDir, strlen - offset))
      unity_main.errorstring = (fileName & " had incompatible value. SaveScansDir was " & frm_collect.m_saveDir & "; updated to " & tmpStrg)
      unity_main.write_error
      frm_collect.m_saveDir = tmpStrg
      unity_main.m_badIniVal = True
    End If
  End If

  ' Append "\" to spectrum file path if not present
  If (Right(frm_collect.m_saveDir, 1) <> "\") Then
    frm_collect.m_saveDir = frm_collect.m_saveDir & "\"
  End If

  ' Setup spectra save directory
  frm_collect.txt_caldir.Text = frm_collect.m_saveDir
  
  ' Confirm path contains '\' instead of '/'
  tmpStrg = frm_collect.m_saveCSVFile
  check_filepathname_delimiters tmpStrg
  frm_collect.m_saveCSVFile = tmpStrg
  
  If (InStr(frm_collect.m_saveCSVFile, "\") = 0) Then
    frm_collect.m_saveCSVFile = (REPORTS_DIR & tmpStrg)
  End If
  
  ' Setup CSV path/file name
  frm_collect.txt_csvfilename.Text = frm_collect.m_saveCSVFile
  
  ' Confirm path contains '\' instead of '/'
  tmpStrg = frm_collect.m_savePredFile
  check_filepathname_delimiters tmpStrg
  frm_collect.m_savePredFile = tmpStrg
  
  If (InStr(frm_collect.m_savePredFile, "\") = 0) Then
    frm_collect.m_savePredFile = (REPORTS_DIR & tmpStrg)
  End If
  
  ' Setup report path/file name
  frm_collect.txtpredfile.Text = frm_collect.m_savePredFile

  ' Setup analysis ticket report selection
  If (frm_collect.m_savePredictions = True) Then
    frm_collect.optsavepredsyes.Value = True
  Else
    frm_collect.optsavepredsno.Value = True
  End If
  
  ' Setup analysis CSV report selection
  If (frm_collect.m_saveCSV = True) Then
    frm_collect.optcsvyes.Value = True
  Else
    frm_collect.optcsvno.Value = True
  End If

  ' Setup analysis dynamic report selection
  If (frm_collect.m_saveDynRpt = True) Then
    frm_collect.opt_dynRptYes.Value = True
  Else
    frm_collect.opt_DynRptNo.Value = True
  End If

#If SSTAR Then
  ' Setup save individual replicate scan selection
  If (frm_collect.m_saveReps = True) Then
    frm_collect.chk_savereps.Value = 1
  Else
    frm_collect.chk_savereps.Value = 0
  End If
#End If

  ' Check for invalid send data to LIMS output value
  If (frm_collect.m_doLIMS < 0) Or (frm_collect.m_doLIMS > 1) Then
    unity_main.errorstring = (fileName & " had incompatible value. Send_LIMS_Output was " & frm_collect.m_doLIMS & "; updated to " & ProdDfltData.sendLimsOutput)
    unity_main.write_error
    frm_collect.m_doLIMS = ProdDfltData.sendLimsOutput
    unity_main.m_badIniVal = True
  End If
  
  ' Setup send data to LIMS output selection
  frm_collect.chk_lims.Value = frm_collect.m_doLIMS
  
  ' Load LIMS output parameters if output enabled
  If (frm_collect.m_doLIMS = 1) Then
    Call frm_POG.load_lims(True)
  End If
  
  ' Load dynamic report parameters if output enabled
  If (frm_collect.m_saveDynRpt = True) Then
    Call frm_dynRptCfg.load_cfg(True)
  End If

#If SSTAR Then
  ' Check for invalid qualification limits
  If (frm_collect.m_smplPPT < frm_collect.numInc_smplPPT.Min) Or (frm_collect.m_smplPPT > frm_collect.numInc_smplPPT.Max) Then
    unity_main.errorstring = (fileName & " had incompatible value. SmplPPT was " & frm_collect.m_smplPPT & "; updated to " & ProdDfltData.smplPPT)
    unity_main.write_error
    frm_collect.m_smplPPT = ProdDfltData.smplPPT
    unity_main.m_badIniVal = True
  End If

  ' Setup product qualification limits
  frm_collect.numInc_smplPPT.Text = frm_collect.m_smplPPT
#End If

  ' Setup spectrum type
PROCESS_STYPE:
  Select Case (LCase(frm_collect.m_sType))
    Case "abs"
      frm_collect.optabs.Value = True
    Case "back"
      frm_collect.optbk.Value = True
    Case Else       ' invalid value
      unity_main.errorstring = (fileName & " had incompatible value. SType was " & frm_collect.m_sType & "; updated to " & ProdDfltData.sType)
      unity_main.write_error
      frm_collect.m_sType = ProdDfltData.sType
      unity_main.m_badIniVal = True
      GoTo PROCESS_STYPE
  End Select

  ' Check for invalid value alarm value
  If (frm_collect.m_alarmProp < 0) Or (frm_collect.m_alarmProp > 1) Then
    unity_main.errorstring = (fileName & " had incompatible value. Value_Alarm was " & frm_collect.m_alarmProp & "; updated to " & ProdDfltData.valueAlarm)
    unity_main.write_error
    frm_collect.m_alarmProp = ProdDfltData.valueAlarm
    unity_main.m_badIniVal = True
  End If
  
  ' Setup value alarm selection
  frmedmod.chk_value.Value = frm_collect.m_alarmProp

  ' Setup ticket printout selection
PROCESS_WRITETKT:
  Select Case (frm_collect.m_writeTkt)
    Case 0          ' no
      frm_collect.opt_tktno.Value = True
    Case 1          ' always
      frm_collect.opt_tktall.Value = True
    Case 2          ' on demand
      frm_collect.opt_tktdemand.Value = True
    Case Else       ' invalid value
      unity_main.errorstring = (fileName & " had incompatible value. Write_Ticket_Printer was " & frm_collect.m_writeTkt & "; updated to " & ProdDfltData.writeTkt)
      unity_main.write_error
      frm_collect.m_writeTkt = ProdDfltData.writeTkt
      unity_main.m_badIniVal = True
      GoTo PROCESS_WRITETKT
  End Select

  ' Check if to load write ticket cfg file
  If (frm_collect.m_writeTkt > 0) Then
    Call frm_ticket.loadticketfile(True)
  Else
    unity_main.img_ticket.Visible = False
  End If

#If SSTAR Then
  ' Setup product & default starting/ending wavelengths
  frm_collect.txt_dfltEndWvln.Text = ProdDfltData.endWvln
  frm_collect.txt_dfltStartWvln.Text = ProdDfltData.startWvln
  frm_collect.txt_endWvln.Text = frm_collect.m_smplEndWvln
  frm_collect.txt_startWvln.Text = frm_collect.m_smplStartWvln
  
  ' Setup instrument min/max wavelengths
  frm_collect.txt_minWvln.Text = unity_main.m_minWvln
  frm_collect.txt_maxWvln.Text = unity_main.m_maxWvln

  ' Check if wavelengths match product defaults
  If (ProdDfltData.endWvln = frm_collect.m_smplEndWvln) And (ProdDfltData.startWvln = frm_collect.m_smplStartWvln) Then
    frm_collect.chk_dfltWavelens.Value = 1
    frm_collect.txt_dfltStartWvln.Visible = True
    frm_collect.txt_dfltEndWvln.Visible = True
    frm_collect.txt_startWvln.Visible = False
    frm_collect.txt_endWvln.Visible = False
    frm_collect.txt_minWvln.Visible = False
    frm_collect.lbl_minWvln.Visible = False
    frm_collect.txt_maxWvln.Visible = False
    frm_collect.lbl_maxWvln.Visible = False
  Else
    frm_collect.chk_dfltWavelens.Value = 0
    frm_collect.txt_dfltStartWvln.Visible = False
    frm_collect.txt_dfltEndWvln.Visible = False
    frm_collect.txt_startWvln.Visible = True
    frm_collect.txt_endWvln.Visible = True
    frm_collect.txt_minWvln.Visible = True
    frm_collect.lbl_minWvln.Visible = True
    frm_collect.txt_maxWvln.Visible = True
    frm_collect.lbl_maxWvln.Visible = True
  End If

  ' Check if to use external reference tray settings
  If (frm_collect.m_bType = "external") Then
    If (frm_extRef.load_ext_ref_cfg_file(frm_collect.m_extRefFileName, prodSelect) = True) Then
      ' Check reference wavelengths against instrument
      If (frm_collect.check_ref_file_wvlns(frm_extRef.m_extRefStartWvln, frm_extRef.m_extRefEndWvln, frm_extRef.m_extRefFileName) = True) Then
        If (frm_collect.m_useExtRefTrayCfg = True) Then
          frm_collect.m_adapterType = frm_extRef.m_extRefAdapterType
          frm_collect.m_multiCupType = frm_extRef.m_extRefMultiCupType
          frm_collect.m_rotateDir = frm_extRef.m_extRefRotateDir
          frm_collect.m_rotateIndexSteps = frm_extRef.m_extRefRotateIndexSteps
          frm_collect.m_rotateMoveMode = frm_extRef.m_extRefRotateMoveMode
          frm_collect.m_rotateSpeed = frm_extRef.m_extRefRotateSpeed
          frm_collect.m_rotateStepSteps = frm_extRef.m_extRefRotateStepSteps
        End If
      End If
    End If
  End If
 
  ' Check and setup product tray configuration parameters
  frm_collect.setup_prod_tray_cfg
#End If
  
  ' Update main screen if loading selected product
  If (prodSelect = True) Then
#If SSTAR Then
    ' Display adapter/tray type used for product
    If (frm_collect.m_adapterType <> CFG_MULTI_CUP_AT) Then
      unity_main.lbl_samplemode.Caption = MS11AdapterInfo(frm_collect.m_adaptIndx).dspName
    Else        ' Display multi-cup type
      unity_main.lbl_samplemode.Caption = MS11MultiCupInfo(frm_collect.m_multiCupIndx).dspName
    End If
  
    ' Display platter movement used for product
    Select Case (frm_collect.m_rotateMoveMode)
      Case TRM_NONE
        unity_main.lbl_movement.Caption = frm_collect.opt_rotateModeNone.Caption
      Case TRM_CONT
        unity_main.lbl_movement.Caption = frm_collect.opt_rotateModeCont.Caption
      Case TRM_STEP
        unity_main.lbl_movement.Caption = frm_collect.opt_rotateModeStep.Caption
      Case TRM_INDEX
        unity_main.lbl_movement.Caption = frm_collect.opt_rotateModeIndex.Caption
    End Select
  
    ' Add rotation direction
    If (frm_collect.m_rotateMoveMode <> TRM_NONE) Then
      If (frm_collect.m_rotateDir = TRD_CW) Then
        unity_main.lbl_movement.Caption = unity_main.lbl_movement.Caption & " - " & MLSupport.GSS("unity_main", "rotateCW", "CW")
      Else
        unity_main.lbl_movement.Caption = unity_main.lbl_movement.Caption & " - " & MLSupport.GSS("unity_main", "rotateCCW", "CCW")
      End If
    End If
#End If

    unity_main.txtsamplename.Text = ""
    unity_main.txtsampcomment.Text = ""
    unity_main.repcounter = 0
    
    unity_main.cmd_repacks.Visible = False
    unity_main.img_report.Visible = False
    unity_main.img_ticket.Visible = False
    unity_main.img_batchRpt.Visible = False
    unity_main.m_batchRptFile = ""
    
    unity_main.setup_olcols
  End If
  
  frmedmod.clearmodtable
  
  ' Load property info if any configured for product
  If (frm_collect.m_numModelVars <> 0) Then
    Call frmProduct.loadmodelsnew(fileName, uniFile, fEncoding, lineCnt, prodSelect)
  Else
    frmedmod.numprops.Text = 0
    frmedmod.grid_models.MaxRows = 0
    frmedmod.m_prdModelType = False
    frmedmod.m_prdFileName = ""
    frmedmod.m_stfFileName = ""
    frmedmod.lst_prdConstituentNames.Clear
    
    If (prodSelect = True) Then
      frmProduct.m_prodSVFChanged = False
      frmProduct.m_expSVFChanged = False
      frmProduct.m_prdModelType = False
      frmProduct.m_numConstituents = 0
    End If
  End If

  ' Update working parameters if loading selected product
  If (prodSelect = True) Then
    unity_main.m_smplAutoScan = False
    unity_main.m_smplManualScan = False
    unity_main.m_remoteSmplScan = False
    
    ' Copy loaded product config into system operational variables
#If ABBFT Then
    unity_main.m_mb3000.m_delayMeasure = frm_collect.m_delayMeasure
    unity_main.m_mb3000.m_delayStart = frm_collect.m_delayStart
    unity_main.m_mb3000.m_endWavenum = frm_collect.m_endWavenum
    unity_main.m_mb3000.m_endWavenumIndx = frm_collect.m_endWavenumIndx
    unity_main.m_mb3000.m_gainIndx = frm_collect.m_gainIndx
    unity_main.m_mb3000.m_numMeasures = frm_collect.m_numMeasures
    unity_main.m_mb3000.m_numSamples = frm_collect.m_numSamples
    unity_main.m_mb3000.m_resolution = frm_collect.m_resolution
    unity_main.m_mb3000.m_resolutionIndx = frm_collect.m_resolutionIndx
    unity_main.m_mb3000.m_smplNumPts = frm_collect.m_smplNumPts
    unity_main.m_mb3000.m_speedIndx = frm_collect.m_speedIndx
    unity_main.m_mb3000.m_startWavenum = frm_collect.m_startWavenum
    unity_main.m_mb3000.m_startWavenumIndx = frm_collect.m_startWavenumIndx
    unity_main.m_mb3000.m_waveNumIncr = frm_collect.m_waveNumIncr
    
    unity_main.m_mb3000.resize_arrays
    init_spectrum_plot
#End If
    
    unity_main.m_adapterType = frm_collect.m_adapterType
    unity_main.m_alarmMD = frm_collect.m_alarmMD
    unity_main.m_alarmND = frm_collect.m_alarmND
    unity_main.m_alarmRR = frm_collect.m_alarmRR
    unity_main.m_alarmProp = frm_collect.m_alarmProp
    unity_main.m_backFreq = frm_collect.m_backFreq
    unity_main.m_bType = frm_collect.m_bType
    unity_main.m_clrManualName = frm_collect.m_clrManualName
    unity_main.m_clrUserInputs = frm_collect.m_clrUserInputs
    unity_main.m_dayCounter = frm_collect.m_dayCounter
    unity_main.m_doLIMS = frm_collect.m_doLIMS
    unity_main.m_extRefFileName = frm_collect.m_extRefFileName
    unity_main.m_hideValCol = frm_collect.m_hideValCol
    unity_main.m_makePred = frm_collect.m_makePred
    unity_main.m_multiCupType = frm_collect.m_multiCupType
    unity_main.m_nameBase = frm_collect.m_nameBase
    unity_main.m_nameCounter = frm_collect.m_nameCounter
    unity_main.m_nameScanType = frm_collect.m_nameScanType
    unity_main.m_noOLVal = frm_collect.m_noOLVal
    unity_main.m_noPredVal = frm_collect.m_noPredVal
    unity_main.m_numModelVars = frm_collect.m_numModelVars
    unity_main.m_olFormat = frm_collect.m_olFormat
    unity_main.m_olRefFileName = frm_collect.m_olRefFileName
    unity_main.m_repsAvg = frm_collect.m_repsAvg
#If SSTAR Then
    unity_main.m_rotateDir = frm_collect.m_rotateDir
    unity_main.m_rotateIndexSteps = frm_collect.m_rotateIndexSteps
    unity_main.m_rotateMoveMode = frm_collect.m_rotateMoveMode
    unity_main.m_rotateSpeed = frm_collect.m_rotateSpeed
    unity_main.m_rotateStepSteps = frm_collect.m_rotateStepSteps
#End If
    unity_main.m_saveCSV = frm_collect.m_saveCSV
    unity_main.m_saveCSVFile = frm_collect.m_saveCSVFile
    unity_main.m_saveDir = frm_collect.m_saveDir
    unity_main.m_saveIt = frm_collect.m_saveIt
    unity_main.m_savePredFile = frm_collect.m_savePredFile
    unity_main.m_savePredictions = frm_collect.m_savePredictions
    unity_main.m_saveReps = frm_collect.m_saveReps
    unity_main.m_saveDynRpt = frm_collect.m_saveDynRpt
    unity_main.m_smplEndWvln = frm_collect.m_smplEndWvln
    unity_main.m_smplNScans = frm_collect.m_smplNScans
    unity_main.m_smplPPT = frm_collect.m_smplPPT
    unity_main.m_smplStartWvln = frm_collect.m_smplStartWvln
    unity_main.m_sNameMode = frm_collect.m_sNameMode
    unity_main.m_sType = frm_collect.m_sType
    unity_main.m_trayNum = frm_collect.m_trayNum
    unity_main.m_useExtRefTrayCfg = frm_collect.m_useExtRefTrayCfg
    unity_main.m_useMIV = frm_collect.m_useMIV
    unity_main.m_valueBound = frm_collect.m_valueBound
    unity_main.m_writeTkt = frm_collect.m_writeTkt

    CreatePath unity_main.m_saveDir

    ' Check if product is configured to use Ucal model
    If (frmProduct.m_prdModelType = True) Then
      unity_main.m_prdModelType = True
      unity_main.m_expSVFChanged = frmProduct.m_expSVFChanged
      unity_main.m_numConstituents = frmProduct.m_numConstituents
      unity_main.m_prodSVFChanged = frmProduct.m_prodSVFChanged
      unity_main.m_prdFileName = frmProduct.m_prdFileName
      unity_main.m_stfFileName = frmProduct.m_stfFileName
      unity_main.m_stfFileValid = frmProduct.m_stfFileValid
      unity_main.m_stfMasterSerNum = frmProduct.m_stfMasterSerNum
      unity_main.m_svfEndWvln = frmProduct.m_svfEndWvln
      unity_main.m_svfIsStd = frmProduct.m_svfIsStd
      unity_main.m_svfStartWvln = frmProduct.m_svfStartWvln
      unity_main.m_svfWaveCnvtFlg = frmProduct.m_svfWaveCnvtFlg
      unity_main.m_svfWaveInc = frmProduct.m_svfWaveInc
    Else
      unity_main.m_prdModelType = False
      unity_main.m_numConstituents = 0
      unity_main.m_prdFileName = ""
      unity_main.m_stfFileName = ""
      unity_main.m_stfFileValid = True
      unity_main.m_stfMasterSerNum = ""
      unity_main.m_svfEndWvln = frm_collect.m_smplEndWvln
      unity_main.m_svfFileName = ""
      unity_main.m_svfIsStd = 0
      unity_main.m_svfStartWvln = frm_collect.m_smplStartWvln
      unity_main.m_svfWaveCnvtFlg = False
      unity_main.m_svfWaveInc = 0
    End If

    ' Setup reference scanning based on type
    Select Case (unity_main.m_bType)
#If SSTAR Then
      Case "internal"
        ' Check if reference performed every sample
        If (unity_main.m_backFreq = REF_FREQ_ALL_SMPLS) Then
          ' Reset internal reference scan info
          unity_main.m_intRefAutoScan = False
          unity_main.m_intRefManualScan = False
        End If
      
        ' Reset external reference scan info
        unity_main.m_extRefAutoScan = False
        unity_main.m_extRefManualScan = False
        unity_main.m_extRefPPTScan = False
        unity_main.m_extRefFileSetup = False
        unity_main.m_extRefPPTFileSetup = False
        unity_main.m_extRefTimer = 0
        unity_main.m_extRefTimeoutSecs = 0
        
        ' Reset offline reference scan info
        unity_main.m_olRefFileSetup = False

        ' Show "Ref" button
        unity_main.cmd_ref.Visible = True

        ' Check if selected product's wavelength range or reference frequency different from last product
        If ((unity_main.m_smplStartWvln <> unity_main.m_lastStartWvln) Or (unity_main.m_smplEndWvln <> unity_main.m_lastEndWvln) Or (unity_main.m_backFreq <> unity_main.m_lastBackFreq)) Then
          ' Check if reference performed on demand
          If (unity_main.m_backFreq = REF_FREQ_ON_DEMAND) Then
            unity_main.m_intRefAutoScan = True
            unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status5", "Reference Required")
          Else
            ' Stop instrument's reference timer
#If SSRCS Then
            SSRCSClientError = SSRCSClient.SetRefTimeout(0)
#Else
            unity_main.MS11srv.refTimeout = 0
#End If
          End If
          
          unity_main.m_lastStartWvln = unity_main.m_smplStartWvln
          unity_main.m_lastEndWvln = unity_main.m_smplEndWvln
          unity_main.m_intRefPPTEndWvln = unity_main.m_smplEndWvln
          unity_main.m_intRefPPTStartWvln = unity_main.m_smplStartWvln
          unity_main.m_lastBackFreq = unity_main.m_backFreq
          
          ' Check if internal reference qualification required
          If (unity_main.m_intRefPPT <> 0) Then
            unity_main.m_intRefPPTFileSetup = False
        
            ' Check if do not have internal reference qualification file for wavelength range
            If (check_int_ref_ppt_file(unity_main.m_smplStartWvln, unity_main.m_smplEndWvln, spcFilename) = False) Then
              unity_main.errorstring = (spcFilename & " file cannot be found. Need to collect internal reference qualification data")
              unity_main.write_error
              CWrap.ShowMessageBoxW MLSupport.GGS_Params("unity_main.errMsg4", "%1 file cannot be found. Need to collect internal reference qualification data over wavelength range %2 - %3", spcFilename, CStr(m_smplStartWvln), CStr(m_smplEndWvln)), vbOKOnly
              unity_main.m_intRefPPTScan = True
              unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status6", "Reference Qualification Required")
            Else
              unity_main.m_intRefPPTScan = False
            End If
          Else
            unity_main.m_intRefPPTScan = False
          End If
        End If
#End If

      Case "external"
        ' Reset internal reference scan info
        unity_main.m_intRefAutoScan = False
        unity_main.m_intRefManualScan = False
        unity_main.m_intRefPPTScan = False
        unity_main.m_intRefPPTFileSetup = False
        
        unity_main.m_lastStartWvln = 0
        unity_main.m_lastEndWvln = 0
        unity_main.m_smplEndWvln = unity_main.m_extRefEndWvln
        unity_main.m_smplStartWvln = unity_main.m_extRefStartWvln
        
        ' Show "Ref" button
        unity_main.cmd_ref.Visible = True
        
#If SSTAR Then
        ' Reset offline reference scan info
        unity_main.m_olRefFileSetup = False
        
        ' Stop reference timer
#If SSRCS Then
        SSRCSClientError = SSRCSClient.SetRefTimeout(0)
#Else
        unity_main.MS11srv.refTimeout = 0
#End If
        
        ' Check if external reference file does not exists
        If (CFile.st_FileExist(REFERENCES_DIR & m_extRefFileName & SPC_FILE_EXT) = False) Then
          unity_main.m_extRefAutoScan = True
          unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status5", "Reference Required")
        End If
        
        ' Check if external reference qualification required
        If (unity_main.m_extRefPPT <> 0) Then
          unity_main.m_extRefPPTAdapterType = unity_main.m_extRefAdapterType
          unity_main.m_extRefPPTAdaptIndx = unity_main.m_extRefAdaptIndx
          unity_main.m_extRefPPTEndWvln = unity_main.m_extRefEndWvln
          unity_main.m_extRefPPTFileName = unity_main.m_extRefFileName
          unity_main.m_extRefPPTFileSetup = False
          unity_main.m_extRefPPTMultiCupIndx = unity_main.m_extRefMultiCupIndx
          unity_main.m_extRefPPTMultiCupType = unity_main.m_extRefMultiCupType
          unity_main.m_extRefPPTNScans = unity_main.m_extRefNScans
          unity_main.m_extRefPPTRotateDir = unity_main.m_extRefRotateDir
          unity_main.m_extRefPPTRotateIndexSteps = unity_main.m_extRefRotateIndexSteps
          unity_main.m_extRefPPTRotateMoveMode = unity_main.m_extRefRotateMoveMode
          unity_main.m_extRefPPTRotateSpeed = unity_main.m_extRefRotateSpeed
          unity_main.m_extRefPPTRotateStepSteps = unity_main.m_extRefRotateStepSteps
          unity_main.m_extRefPPTStartWvln = unity_main.m_extRefStartWvln
          unity_main.m_extRefPPTTrayNum = unity_main.m_extRefTrayNum
        
          ' Check if do not have external reference qualification file for wavelength range
          If (check_ext_ref_ppt_file(m_extRefPPTFileName, spcFilename) = False) Then
            unity_main.errorstring = (spcFilename & " file cannot be found. Need to collect external reference qualification data")
            unity_main.write_error
            CWrap.ShowMessageBoxW MLSupport.GGS_Params("unity_main.errMsg6", "%1 file cannot be found. Need to collect external reference qualification data over wavelength range %2 - %3", spcFilename, CStr(m_extRefPPTStartWvln), CStr(m_extRefPPTEndWvln)), vbOKOnly
            unity_main.m_extRefPPTScan = True
            unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status6", "Reference Qualification Required")
          Else
            unity_main.m_extRefPPTScan = False
          End If
        Else
          unity_main.m_extRefPPTScan = False
        End If
#End If

#If SSTAR Then
      Case "file"    ' from offline file
        ' Reset internal reference scan info
        unity_main.m_intRefAutoScan = False
        unity_main.m_intRefManualScan = False
        unity_main.m_intRefPPTScan = False
        unity_main.m_intRefPPTFileSetup = False
        unity_main.m_lastStartWvln = 0
        unity_main.m_lastEndWvln = 0
        
        ' Reset external reference scan info
        unity_main.m_extRefAutoScan = False
        unity_main.m_extRefManualScan = False
        unity_main.m_extRefPPTScan = False
        unity_main.m_extRefFileSetup = False
        unity_main.m_extRefPPTFileSetup = False
        unity_main.m_extRefTimer = 0
        unity_main.m_extRefTimeoutSecs = 0
        
        ' Stop reference timer
#If SSRCS Then
        SSRCSClientError = SSRCSClient.SetRefTimeout(0)
#Else
        unity_main.MS11srv.refTimeout = 0
#End If
        
        ' Hide "Ref" button
        unity_main.cmd_ref.Visible = False

        ' Check reference wavelengths against instrument
        If (frm_olRef.get_ol_ref_wvlns(unity_main.m_olRefFileName) = True) Then
          Call frm_collect.check_ref_file_wvlns(frm_olRef.m_olRefStartWvln, frm_olRef.m_olRefEndWvln, frm_olRef.m_olRefFileName)
        End If

        unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status10", "Ready for New Sample")
#End If
    End Select
  End If
End Sub

Private Sub load_prod_file_vals(ByVal fileName As String, ByVal uniFile As clsUniFile, ByVal fEncoding As eFileEncoding, ByRef lineCnt As Integer)
  Dim xx As Variant
  Dim inString As String
  Dim tmpStrg As String
  Dim cfgVar As String
  Dim varVal As String
  Dim strlen As Integer
  Dim rc As Boolean

  unity_main.m_iniString = ""

  ' Process each line in analysis settings section
  While Not (uniFile.EOF())
    ' Read line from file
    On Error GoTo FILE_ERROR
    lineCnt = lineCnt + 1
    
    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(inString)
    Else
      rc = uniFile.ReadUnicodeLine(inString)
    End If
      
    If (rc = False) Then GoTo FILE_ERROR
          
    If (InStr(1, inString, "[") <> 0) Then
      unity_main.m_iniString = inString
      Exit Sub
    End If
          
    ' Get variable name and its value
    xx = InStr(1, inString, "=")
    strlen = Len(inString)
    tmpStrg = Trim(Mid(inString, 1, xx - 1))
    cfgVar = LCase(tmpStrg)
    varVal = Trim(Mid(inString, xx + 1))
          
    ' Process value by variable name
    On Error GoTo BAD_INI_VALUE
    Select Case (cfgVar)
      Case "adaptertype"
        frm_collect.m_adapterType = varVal
      Case "background_frequency"
        frm_collect.m_backFreq = CInt(varVal)
      Case "bound_values"
        frm_collect.m_valueBound = CInt(varVal)
      Case "btype"
        frm_collect.m_bType = varVal
      Case "clrmanualname"
        frm_collect.m_clrManualName = CBool(varVal)
      Case "clruserinputs"
        frm_collect.m_clrUserInputs = CBool(varVal)
      Case "day_counter"
        frm_collect.m_dayCounter = CLng(varVal)
      Case "delaymeasure"
        frm_collect.m_delayMeasure = CLng(varVal)
      Case "delaystart"
        frm_collect.m_delayStart = CLng(varVal)
      Case "endwavenumindx"
        frm_collect.m_endWavenumIndx = CLng(varVal)
      Case "endwvln"
        frm_collect.m_smplEndWvln = CDbl(varVal)
      Case "extreffile"
        frm_collect.m_extRefFileName = varVal
      Case "gainindx"
        frm_collect.m_gainIndx = CInt(varVal)
      Case "hidevalcol"
        frm_collect.m_hideValCol = CBool(varVal)
      Case "makepred"
        frm_collect.m_makePred = varVal
      Case "md_alarm"
        frm_collect.m_alarmMD = CInt(varVal)
      Case "menu_input_buttons"
        frm_collect.m_useMIV = CBool(varVal)
      Case "multicuptype"
        frm_collect.m_multiCupType = varVal
      Case "namebase"
        frm_collect.m_nameBase = varVal
      Case "namecounter"
        frm_collect.m_nameCounter = CLng(varVal)
      Case "namescantype"
        strlen = Len(varVal)
        
        If (strlen > 0) Then
          Dim tmpString As String
          tmpString = UCase(Left(varVal, 1))
          tmpString = tmpString & LCase(Right(varVal, strlen - 1))
          frm_collect.m_nameScanType = tmpString
        Else
          frm_collect.m_nameScanType = varVal
        End If
      Case "nd_alarm"
        frm_collect.m_alarmND = CInt(varVal)
      Case "noolval"
        frm_collect.m_noOLVal = varVal
      Case "nopredval"
        frm_collect.m_noPredVal = varVal
      Case "nscanss"
        frm_collect.m_smplNScans = CInt(varVal)
      Case "nummeasures"
        frm_collect.m_numMeasures = CInt(varVal)
      Case "numsamples"
        frm_collect.m_numSamples = CInt(varVal)
      Case "olreffile"
        frm_collect.m_olRefFileName = varVal
      Case "outlier_lights"
        frm_collect.m_olFormat = CBool(varVal)
      Case "repsavg"
        frm_collect.m_repsAvg = CInt(varVal)
      Case "resolutionindx"
        frm_collect.m_resolutionIndx = CInt(varVal)
#If SSTAR Then
      Case "rotatedir"
        frm_collect.m_rotateDir = CInt(varVal)
      Case "rotateindexsteps"
        frm_collect.m_rotateIndexSteps = CInt(varVal)
      Case "rotatemovemode"
        frm_collect.m_rotateMoveMode = CInt(varVal)
      Case "rotatespeed"
        frm_collect.m_rotateSpeed = CInt(varVal)
      Case "rotatestepsteps"
        frm_collect.m_rotateStepSteps = CInt(varVal)
#End If
      Case "rr_alarm"
        frm_collect.m_alarmRR = CInt(varVal)
      Case "saveit"
        frm_collect.m_saveIt = varVal
      Case "savescansdir"
        frm_collect.m_saveDir = varVal
      Case "save_csvfile"
        frm_collect.m_saveCSVFile = varVal
      Case "save_predfile"
        frm_collect.m_savePredFile = varVal
      Case "save_predictions"
        frm_collect.m_savePredictions = CBool(varVal)
      Case "save_predictions_csv"
        frm_collect.m_saveCSV = CBool(varVal)
      Case "save_predictions_dynrpt"
        frm_collect.m_saveDynRpt = CBool(varVal)
      Case "save_replicates"
        frm_collect.m_saveReps = CBool(varVal)
      Case "send_lims_output"
        frm_collect.m_doLIMS = CInt(varVal)
      Case "smplppt"
        frm_collect.m_smplPPT = CInt(varVal)
      Case "speedindx"
        frm_collect.m_speedIndx = CInt(varVal)
      Case "startwavenumindx"
        frm_collect.m_startWavenumIndx = CLng(varVal)
      Case "startwvln"
        frm_collect.m_smplStartWvln = CDbl(varVal)
      Case "stype"
        frm_collect.m_sType = varVal
      Case "useextreftraycfg"
        frm_collect.m_useExtRefTrayCfg = CBool(varVal)
      Case "value_alarm"
        frm_collect.m_alarmProp = CInt(varVal)
      Case "write_ticket_printer"
        frm_collect.m_writeTkt = CInt(varVal)
    End Select
  Wend
  
  Exit Sub
  
BAD_INI_VALUE:
    unity_main.errorstring = (fileName & " line " & CStr(lineCnt) & " had incompatible value. " & Trim(tmpStrg) & " - " & varVal & "; will use default value")
    unity_main.write_error
    unity_main.m_badIniVal = True
    Resume Next
  
FILE_ERROR:
  unity_main.m_iniString = ""
End Sub

Private Sub seeifcalstar()
  Dim dirname, fileName As String
  Dim fileFound1 As Boolean
  Dim fileFound2 As Boolean
  Dim fileFound3 As Boolean
  
  dirname = "c:\winnt\system32\"
  fileName = "DBAccess.dll"
  fileFound1 = CFile.st_FileExist(dirname & fileName)
  
  fileName = "QuanPredict.dll"
  fileFound2 = CFile.st_FileExist(dirname & fileName)
  
  fileName = "SL_ct32.dll"
  fileFound3 = CFile.st_FileExist(dirname & fileName)
  
  If (fileFound1 = True) And (fileFound2 = True) And (fileFound3 = True) Then
    unity_main.calstar_enabled = True
  Else
    dirname = "c:\windows\system32\"
    fileName = "DBAccess.dll"
    fileFound1 = CFile.st_FileExist(dirname & fileName)
  
    fileName = "QuanPredict.dll"
    fileFound2 = CFile.st_FileExist(dirname & fileName)
  
    fileName = "SL_ct32.dll"
    fileFound3 = CFile.st_FileExist(dirname & fileName)
    
    If (fileFound1 = True) And (fileFound2 = True) And (fileFound3 = True) Then
      unity_main.calstar_enabled = True
    End If
  End If
End Sub

Private Sub chk_ucal_files_installed()
  Dim dirname, fileName As String
  Dim fileFound1 As Boolean
  Dim fileFound2 As Boolean
  Dim prodID As Long
  Dim majVer As Long
  Dim minVer As Long
  Dim features As Integer
  Dim rc As Long
  Dim errMsg As String
  Dim uniMsg As String
  
  dirname = SOFTWARE_DIR
  fileName = "PRDComponent.dll"
  fileFound1 = CFile.st_FileExist(dirname & fileName)
  
  fileName = "SVFComponent.dll"
  fileFound2 = CFile.st_FileExist(dirname & fileName)
  
  If (fileFound1 = True) And (fileFound2 = True) Then
    ' Setup to disable security check
    prodID = -1522756
    majVer = 32239684
    minVer = -81216144
    
    On Error GoTo PRD_OBJECT_ERROR
    Set PRDObject = New PRDComponentClass
    
    ' Initialize component
    rc = PRDObject.initComponent(prodID, majVer, minVer, vbNull, features, vbNull)
    
    If (rc = 0) Then
      On Error GoTo SVF_OBJECT_ERROR
      Set SVFObject = New SVFComponentClass
      rc = SVFObject.initComponent(prodID, majVer, minVer, vbNull, features, vbNull)
      
      If (rc = 0) Then
        unity_main.m_prdEnabled = True
        Exit Sub
      Else
        errMsg = "SVFComponent initComponent() error: " & rc
        uniMsg = MLSupport.GGS_Params("errMsg8", "%1 %2 error: %3", "SVFComponent", "initComponent()", CStr(rc))
        CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
      End If
    Else
      errMsg = "PRDComponent initComponent() error: " & rc
      uniMsg = MLSupport.GGS_Params("errMsg8", "%1 %2 error: %3", "PRDComponent", "initComponent()", CStr(rc))
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    End If
  End If
  
PRD_OBJECT_ERROR:
  unity_main.errorstring = "Unity PRDComponent.dll component not installed or registered"
  unity_main.write_error
  uniMsg = MLSupport.GGS_Params("errMsg7", "Unity %1 component not installed or registered", "PRDComponent.dll")
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  Exit Sub
  
SVF_OBJECT_ERROR:
  unity_main.errorstring = "Unity SVFComponent.dll component not installed or registered"
  unity_main.write_error
  uniMsg = MLSupport.GGS_Params("errMsg7", "Unity %1 component not installed or registered", "SVFComponent.dll")
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub

Private Sub chk_net_framework_installed()
  Dim netKeyStrg As String
  Dim hKey As Long

  ' Check if .NET Framework 2.0 is installed; required for PRD modeling
  netKeyStrg = "SOFTWARE\Microsoft\.NETFramework\policy\v2.0"

  If (RegOpenKey(HKEY_LOCAL_MACHINE, netKeyStrg, 0&, KEY_QUERY_VALUE, hKey) <> RC_SUCCESS) Then
    unity_main.m_netFWInstalled = False
  Else
    unity_main.m_netFWInstalled = True
  End If
  
  RegCloseKey (hKey)
End Sub

Private Sub clear_avg_tbl()
  
  frmpreds.gridpreds.Row = 0
  frmpreds.gridpreds.Col = 0
  frmpreds.gridpreds.Row2 = MAX_NUM_PROPS
  frmpreds.gridpreds.Col2 = MAX_NUM_REPACKS
  frmpreds.gridpreds.BlockMode = True
  frmpreds.gridpreds.Action = 3
  frmpreds.gridpreds.BlockMode = False
End Sub

Private Sub load_last50()
  Dim ret As Boolean
  Dim rowCnt As Long
  Dim colCnt As Long

  ret = ss_last50.LoadFromFile(REPORTS_DIR & HIST_LOG_FILE)
  
  rowCnt = ss_last50.DataRowCnt
  
  If (rowCnt > MAX_LAST50_ROWS) Then
    ss_last50.DeleteRows 1, rowCnt - MAX_LAST50_ROWS
  End If
  
  ' Determine max columns for each row
  If (ret = True) Then
    For rowCnt = 1 To ss_last50.DataRowCnt
      ss_last50.Row = rowCnt
      
      For colCnt = 6 To ss_last50.MaxCols
        ss_last50.Col = colCnt
          
        ' Check if no value for property
        If (ss_last50.Text = "") Then
          Exit For
        Else
          colCnt = colCnt + 1
        End If
      Next colCnt
        
      ' Save max column used for row
      colCnt = colCnt - 1
      lst_last50MaxCols.AddItem colCnt
    Next rowCnt
  End If
  
  ss_last50.StartingRowNumber = 1
  ss_last50.ColHeadersShow = True
  ss_last50.MaxRows = MAX_LAST50_ROWS
  
  ss_last50.Row = 0
  ss_last50.Col = 0
  ss_last50.Row2 = 0
  ss_last50.Col2 = ss_last50.MaxCols
  ss_last50.BlockMode = True
  ss_last50.Font.Bold = True
  ss_last50.Font.Name = "Arial Unicode MS"
  ss_last50.Font.Size = 8
  ss_last50.BlockMode = False
  
  ss_last50.Row = 1
  ss_last50.Col = 0
  ss_last50.Row2 = ss_last50.MaxRows
  ss_last50.Col2 = ss_last50.MaxCols
  ss_last50.BlockMode = True
  ss_last50.Font.Bold = False
  ss_last50.Font.Name = "Arial Unicode MS"
  ss_last50.Font.Size = 8
  ss_last50.BlockMode = False
  
  pos_last50_toprow
  build_last50_colhdr

  If (ret = False) Then
    ss_last50.MaxCols = MIN_LAST50_COLS
  Else
    ss_last50.MaxCols = ss_last50.DataColCnt
  End If
End Sub

Private Sub build_last50_colhdr()
  Dim dd, nn As Integer

  ss_last50.Row = 0
  ss_last50.Col = 1
  ss_last50.Text = MLSupport.GSS("Headers", "date", "Date")
  ss_last50.Col = 2
  ss_last50.Text = MLSupport.GSS("Headers", "time", "Time")
  ss_last50.Col = 3
  ss_last50.Text = MLSupport.GSS("Headers", "product", "Product")
  ss_last50.Col = 4
  ss_last50.Text = MLSupport.GSS("Headers", "sampleID", "Sample ID")
  ss_last50.Col = 5
  ss_last50.Text = MLSupport.GSS("Headers", "comment", "Comment")
  
  nn = 1
  
  For dd = 6 To (ss_last50.MaxCols - 1) Step 2
    ss_last50.Col = dd
    ss_last50.Text = MLSupport.GSS("Headers", "property", "Property") & " " & CStr(nn)
    ss_last50.Col = dd + 1
    ss_last50.Text = MLSupport.GSS("Headers", "value", "Value") & " " & CStr(nn)
    nn = nn + 1
  Next dd
End Sub

Private Sub pos_last50_toprow()
  
#If SSTAR Then
  If (MS11CfgData.devID = DTID_TOPWIND0) Or (MS11CfgData.devID = DTID_TOPWIND1) Then
    If (ss_last50.DataRowCnt > 9) Then
      ss_last50.TopRow = ss_last50.DataRowCnt - 9
    End If
  Else
    If (ss_last50.DataRowCnt > MAX_LAST50_TOPROW) Then
      ss_last50.TopRow = ss_last50.DataRowCnt - MAX_LAST50_TOPROW
    End If
  End If
#Else
  If (ss_last50.DataRowCnt > 9) Then
    ss_last50.TopRow = ss_last50.DataRowCnt - 9
  End If
#End If
End Sub

Private Sub setupspread()
  Dim zz As Integer
  Dim fileName As String
  Dim errMsg As String
  Dim uniMsg As String

  For zz = 0 To MAX_NUM_PROPS
    fpspread_pred.RowHeight(zz) = (fpspread_pred.Height / 9)
  Next zz
  
  fpspread_pred.Row = 0
  fpspread_pred.Col = 1
  fpspread_pred.Text = MLSupport.GSS("Headers", "property", "Property")
  fpspread_pred.Col = 2
  fpspread_pred.Text = MLSupport.GSS("Headers", "value", "Value")
  
  fpspread_pred.Row = 0
  fpspread_pred.Col = 3
  fpspread_pred.Row2 = MAX_NUM_PROPS
  fpspread_pred.Col2 = unity_main.fpspread_pred.MaxCols
  fpspread_pred.BlockMode = True
  fpspread_pred.CellType = CellTypePicture
  fpspread_pred.BlockMode = False

  On Error GoTo BADBMP
  
  For zz = 3 To 6
    fpspread_pred.Col = zz
    fpspread_pred.TypePictStretch = True
    fpspread_pred.TypePictCenter = True
    fpspread_pred.TypePictMaintainScale = True
    fpspread_pred.BackColor = vbBlack
  
    Select Case (zz)
      Case 3
        fileName = (GRAPHICS_DIR & "CHECK_M.BMP")
      Case 4
        fileName = (GRAPHICS_DIR & "CHECK_R.BMP")
      Case 5
        fileName = (GRAPHICS_DIR & "CHECK_V.BMP")
      Case 6
        fileName = (GRAPHICS_DIR & "CHECK_N.BMP")
    End Select
    
    fpspread_pred.TypePictPicture = LoadPicture(fileName)
  Next zz
  
  setup_olcols
  frmedmod.fixthesize
  Exit Sub
  
BADBMP:
  errMsg = (fileName & " file open error." & Error$)
  errMsg = (errMsg & ". InfoStar will shutdown automatically")
  unity_main.errorstring = errMsg
  unity_main.write_error
  uniMsg = MLSupport.GGS_Params("fileErrMsg2", "%1 file open error. %2", fileName, Error$)
  uniMsg = MLSupport.GGS_Params("errMsg2", "%1. InfoStar will shutdown automatically", uniMsg)
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  unity_main.unloadallforms "unity_main"
  Unload Me
  End
End Sub

Private Sub cmd_bias_Click()

  Call unity_main.kill_loop(LOG_DBG_LEVEL3, False, "Main screen Bias button selected")
  unity_main.utiltoopen = 8
  frmLogin.Show 1
End Sub

#If SSTAR Then
Private Sub cmd_verifyRef_Click()

  Call unity_main.kill_loop(LOG_DBG_LEVEL3, False, "Main screen Verify Reference button selected")
  unity_main.utiltoopen = 11
  frmLogin.Show 1
End Sub
#End If

Private Sub cmd_clientAppl_Click(Index As Integer)

  IPCServer1.ClientApplForward Index
End Sub



Private Sub cmd_options_Click()

  Call unity_main.kill_loop(LOG_DBG_LEVEL3, False, "Main screen Options button selected")
  frmUtils.Show 1
End Sub

Private Sub cmd_ref_Click()
  
  If (cmd_start.Visible = True) Then
    cmd_start.Visible = False
    cmd_stop.Visible = True
  End If
  
  Call unity_main.restart_loop(LOG_DBG_LEVEL3, "Main screen Ref button selected")
  
  Select Case (unity_main.m_bType)
    Case "internal"
      unity_main.m_intRefManualScan = True
    Case "external"
      unity_main.m_extRefManualScan = True
  End Select
  
  unity_main.tmr_all.enabled = True
End Sub

Private Sub cmd_repacks_Click()
  
  Call unity_main.kill_loop(LOG_DBG_LEVEL3, False, "Main screen Repacks button selected")
  frm_repacks.uniTitle = MLSupport.GGS_Params("unity_main.statMsg18", "Individual Repack Readings for %1", unity_main.txtsamplename.Text)
  frm_repacks.Show 1
End Sub

#If SSTAR Then
Private Sub cmd_runBatch_Click()
  Dim errMsg As String
  
  Call unity_main.kill_loop(LOG_DBG_LEVEL3, False, "Main screen Run Batch button selected")
  
  If (NumAutoSmplrTowers = 0) Then
    errMsg = MLSupport.GGS_Params("unity_main_errMsg10", "Failure to communicate with auto-sampler. Make sure serial port %1 is not in used by other InfoStar options or programs. Check that the auto-sampler is operating and is connected to the SpectraStar", CStr(m_autoSmplrPort))
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", errMsg), vbCritical
    
    ' Initialize serial port communication w/ auto-sampler
#If SSRCS Then
    Dim parity As String
    parity = AUTO_SMPLR_PARITY
    SSRCSClientError = SSRCSClient.InitASComms(unity_main.m_autoSmplrPort, AUTO_SMPLR_BAUD, parity, AUTO_SMPLR_RSP_TIMEOUT, AUTO_SMPLR_NUM_RETRIES, AUTO_SMPLR_CRC_USAGE)
    
    If (SSRCSClientError = 0) Then
      ' Get max number of tubes supported
      SSRCSClientError = SSRCSClient.GetAllASTubesState
    Else
      If (unity_main.m_ssrcsConnected = True) Then
        CWrap.ShowMessageBoxW MLSupport.GGS_Params("ASErrMsg1", "Problem with configuring Auto-Sampler comm port %1. Verify no other communication option is configured to use this port", CStr(unity_main.m_autoSmplrPort)), vbCritical
      Else
        CWrap.ShowMessageBoxW MLSupport.GSS("frm_ssrcsConnect", "errMsg3", "Not connected to a SpectraStar RCS server"), vbCritical
      End If
    End If
#Else
    If (frm_batchRun.autoSmplr.init_comms(unity_main.m_autoSmplrPort, AUTO_SMPLR_BAUD, AUTO_SMPLR_PARITY, AUTO_SMPLR_RSP_TIMEOUT, AUTO_SMPLR_NUM_RETRIES) = True) Then
      ' Get max number of tubes supported
      frm_batchRun.autoSmplr.get_all_tubes_state AUTO_SMPLR_CRC_USAGE
    Else
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("ASErrMsg1", "Problem with configuring Auto-Sampler comm port %1. Verify no other communication option is configured to use this port", CStr(unity_main.m_autoSmplrPort)), vbCritical
    End If
#End If
  Else
    frm_batchRunCfg.init_cfg False
    frm_batchRunCfg.Show 1
  End If
End Sub
#End If

Private Sub cmd_sample_Click()
  
  If ((unity_main.m_intRefPPTScan = True) Or (unity_main.m_extRefPPTScan = True)) And (unity_main.repcounter = 0) Then
    CWrap.ShowMessageBoxW (MLSupport.GSS("OperStatus", "status6", "Reference Qualification Required") & " - " & MLSupport.GSS("OperStatus", "status2", "Press Ref")), vbOKOnly
  Else
    If ((unity_main.m_intRefAutoScan = True) Or (unity_main.m_extRefAutoScan = True)) And (unity_main.repcounter = 0) Then
      CWrap.ShowMessageBoxW (MLSupport.GSS("OperStatus", "status5", "Reference Required") & " - " & MLSupport.GSS("OperStatus", "status2", "Press Ref")), vbOKOnly
    Else
      If (unity_main.m_remoteRefScan = True) And (unity_main.repcounter = 0) Then
        CWrap.ShowMessageBoxW (MLSupport.GSS("OperStatus", "status9", "Reference Requested Remotely") & " - " & MLSupport.GSS("OperStatus", "status2", "Press Ref")), vbOKOnly
      Else
        If (cmd_start.Visible = True) Then
          cmd_start.Visible = False
          cmd_stop.Visible = True
        End If
  
        Call unity_main.restart_loop(LOG_DBG_LEVEL3, "Main screen Scan Sample button selected")
        unity_main.m_smplManualScan = True
        unity_main.tmr_all.enabled = True
      End If
    End If
    cmd_sample.enabled = False
  End If
  
End Sub

Private Sub cmd_select_Click()

  Call unity_main.kill_loop(LOG_DBG_LEVEL3, False, "Main screen Select Product button selected")
  FRM_SEL_PRODUCT.loadproducts
  FRM_SEL_PRODUCT.Show 1
End Sub

#If SSRCS Then
Private Sub cmd_ssrcsConnect_Click()

  Call unity_main.kill_loop(LOG_DBG_LEVEL3, False, "Main screen Connect to SS button selected")
  frm_ssrcsConnect.initButtons False
  frm_ssrcsConnect.Show 1
End Sub
#End If

Private Sub cmd_start_Click()

  cmd_start.Visible = False
  cmd_stop.Visible = True
  Call unity_main.restart_loop(LOG_DBG_LEVEL1, "Main screen start button selected")
End Sub

Private Sub cmd_stop_Click()
  
  cmd_start.Visible = True
  cmd_stop.Visible = False
  Call unity_main.kill_loop(LOG_DBG_LEVEL1, True, "Main screen Stop button selected")
End Sub

Private Sub Form_Initialize()

#If SSTAR Then
  Set m_ms11srvGNEventQ = New Collection
#End If
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
#If ABBFT Then
  cmd_chkRef.Visible = False
  cmd_runBatch.Visible = False
#End If

  unity_main.checklogfile
End Sub

Private Sub Form_Unload(Cancel As Integer)

#If SSTAR Then
  Set m_ms11srvGNEventQ = Nothing
#End If
End Sub

 

 

#If SSTAR Then
Private Sub img_batchRpt_Click()

  Call unity_main.kill_loop(LOG_DBG_LEVEL3, False, "Main screen View Batch Report icon selected")
  frm_batchRpt.show_report unity_main.m_batchRptFile
  frm_batchRpt.cmd_delete.Visible = False
  frm_batchRpt.m_restartLoop = True
  frm_batchRpt.Show 1
End Sub
#End If

Private Sub img_binocs_Click()
  
  Call unity_main.kill_loop(LOG_DBG_LEVEL3, False, "Main screen View Sample Spectra icon selected")
  frm_scan.Show 1
End Sub

Private Sub img_csv_Click()
  
  Call unity_main.kill_loop(LOG_DBG_LEVEL3, False, "Main screen View Sample CVS File icon selected")
  frm_csvs2.cmd_delete.Visible = False
  frmUtils.loadcsv
End Sub

Private Sub img_help_Click()
  
  Call unity_main.kill_loop(LOG_DBG_LEVEL3, False, "Main screen InfoStar Help icon selected")
  frm_help.Show 1
End Sub

Private Sub img_ticket_Click()
  
  frm_ticket.writeticket
End Sub

Private Sub img_report_Click()
  
  Call unity_main.kill_loop(LOG_DBG_LEVEL3, False, "Main screen Print Sample Analysis Report icon selected")
  show_report
End Sub

Private Sub show_report()
  Dim jj As Integer
  Dim printString As String
    
  If (Trim(unity_main.txtsamplename.Text) = "") Then
    CWrap.ShowMessageBoxW MLSupport.GSS("unity_main", "errMsg4", "No sample scan has been performed")
    Exit Sub
  End If
  
  On Error Resume Next
  frmReport.lst_rept.AddItem MLSupport.GSS("unity_main", "statMsg2", "Unity Scientific Analysis Report")
  frmReport.lst_rept.AddItem MLSupport.GSS("unity_main", "statMsg5", "Analysis Date") & ": " & unity_main.lbl_date.Caption
  frmReport.lst_rept.AddItem MLSupport.GSS("unity_main", "statMsg6", "Analysis Time") & ": " & unity_main.lbl_time.Caption
  frmReport.lst_rept.AddItem MLSupport.GSS("Headers", "serNum", "Serial No.") & ": " & unity_main.m_sysSerialNum
  frmReport.lst_rept.AddItem MLSupport.GGS_Params("unity_main.statMsg22", "Sample Type: %1", unity_main.lblProd1.Caption)
  frmReport.lst_rept.AddItem MLSupport.GGS_Params("unity_main.statMsg23", "Sample Name: %1", Trim(unity_main.txtsamplename.Text))
  frmReport.lst_rept.AddItem MLSupport.GGS_Params("unity_main.statMsg24", "Sample Comment: %1", unity_main.txtsampcomment.Text)
  frmReport.lst_rept.AddItem ""
  frmReport.lst_rept.AddItem MLSupport.GSS("unity_main", "statMsg3", "Property, Value")
  
  For jj = 1 To Trim(frmedmod.numprops.Text)
    printString = ""
    unity_main.fpspread_pred.Row = jj
    unity_main.fpspread_pred.Col = 1
    printString = Trim(unity_main.fpspread_pred.Text) & " = "
    unity_main.fpspread_pred.Col = 2
    printString = printString & Trim(unity_main.fpspread_pred.Text)
    frmReport.lst_rept.AddItem printString
  Next jj
  
  frmReport.loadscancw
  frmReport.Show 1
End Sub

Private Sub IPCServer_GetCurrentProduct(ByVal socketIndx As Integer)
  Dim Status As InfoStarStatusCodes
  
  If (lblProd1.Caption = "") Then
    Status = InfoStarStatusCodes.NoCurrProductStat
  Else
    Status = InfoStarStatusCodes.GoodStat
  End If
    
  send_current_product socketIndx, Status
End Sub

Private Sub IPCServer_GetProductsList(ByVal socketIndx As Integer)

  send_new_products_list socketIndx
End Sub

Private Sub IPCServer_SelectProduct(ByVal socketIndx As Integer, prodName As String)

  unity_main.m_remoteProdName = prodName
  unity_main.m_remoteProdSelect = 1
  unity_main.m_remoteSocketIndx = socketIndx
End Sub

Private Sub IPCServer_SetupButton(ByVal buttonNum As Integer, ByVal buttonVisible As Boolean, buttonLabel As String)

  cmd_clientAppl(buttonNum).Caption = buttonLabel
  cmd_clientAppl(buttonNum).Visible = buttonVisible
End Sub

Private Sub IPCServer_StartScan(ByVal socketIndx As Integer, ByVal scanType As InfoStarIPCServer.ScanTypes)
  Dim Status As InfoStarStatusCodes
  
  If (scanType = InfoStarIPCServer.ScanTypes.RefScanType) Then
    If (lblProd1.Caption = "") Then
      Status = InfoStarStatusCodes.NoCurrProductStat
      IPCServer1.NewSpectrum socketIndx, Status, "", SpectrumTypes.RefSpectrumType, "", 0, 0
    Else
      If (unity_main.m_bType = "file") Then
        Status = InfoStarStatusCodes.InvScanTypeStat
        IPCServer1.NewSpectrum socketIndx, Status, "", SpectrumTypes.RefSpectrumType, "", 0, 0
      Else
        unity_main.m_remoteRefScan = True
      End If
    End If
  Else
    If (lblProd1.Caption = "") Then
      Status = InfoStarStatusCodes.NoCurrProductStat
      IPCServer1.NewSpectrum socketIndx, Status, "", SpectrumTypes.ProductSpectrumType, "", 0, 0
    Else
      unity_main.m_remoteSmplScan = True
    End If
  End If
End Sub

#If MS11srv Then
Private Sub MS11srv_GeneralNotice(ByVal notice As Long)
  Dim gnEvent As New clsMS11srvGNEvent
  Dim trayStatus As Long
    
'  frm_status.txt_GenNotice.Text = notice     ' *** DEBUG USAGE ONLY ***
'  frm_status.txt_TrayStatus.Text = "0x" & Hex(MS11srv.trayStatus)  ' *** DEBUG USAGE ONLY ***

  Select Case (notice)
    Case EVGN_INITDONE            ' Initialization Sequence Passed Successfully
      MS11Initialized = True
      Exit Sub
      
    Case EVGN_TRAYFLD             ' Tray Failure
   
    Case EVGN_SCNRFLD             ' Scanner Failure
      
    Case EVGN_SCANABRTD           ' Current Scan has been Aborted
      
    Case EVGN_SCANSTOPD           ' Scanner Stopped Scanning; Resume can be initiated
      
    Case EVGN_IMPOSERR            ' During Scan/Tray Engine Running, Impossible State Transitions occurred
      
    Case EVGN_REFRNCTMO           ' Reference has Timed Out and is now Expired
      ' Check if internal reference performed on demand
      If (unity_main.m_backFreq = REF_FREQ_ON_DEMAND) Then
        unity_main.m_intRefAutoScan = True
        unity_main.lbl_backdate.Caption = MLSupport.GSS("unity_main", "lbl_backdate_int2", "Internal Reference has Expired")
      
        If ((unity_main.m_batchRunFlg = False) And (unity_main.repcounter = 0)) Then
          unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status5", "Reference Required")
        End If
      Else
        ' Stop instrument's reference timer
        unity_main.MS11srv.refTimeout = 0
      End If
      
      Exit Sub
      
    Case EVGN_SMPLSCNOK           ' Sample Data Scan Successfully Completed
      unity_main.m_scanDblTimestamp = Now
      unity_main.m_scanTimestamp = CDate(m_scanDblTimestamp)
      
    Case EVGN_SMPLSCNBAD          ' Sample Data Scan Collected but Bad (Clipped or Failed PPT Qualification)
      
    Case EVGN_REFSCNOK            ' Reference Data Scan Successfully Completed
      unity_main.m_scanDblTimestamp = Now
      unity_main.m_scanTimestamp = CDate(m_scanDblTimestamp)
      
    Case EVGN_REFSCNBAD           ' Reference Data Scan Collected but Bad (Clipped or Failed PPT Qualification)
      
    Case EVGN_OPENDRWR            ' Request to OPEN Sample Drawer
      frm_status.cmd_exitScan.Visible = False
      frm_status.cmd_resumeScan.Visible = False
      frm_status.cmd_retryScan.Visible = False
      frm_status.cmd_abortScan.Visible = True
      
    Case EVGN_CLOSEDRWR           ' Request to CLOSE Sample Drawer
      frm_status.cmd_exitScan.Visible = False
      frm_status.cmd_resumeScan.Visible = False
      frm_status.cmd_retryScan.Visible = False
      frm_status.cmd_abortScan.Visible = True
      
    Case EVGN_POSXREFRNC          ' Request to Position External Reference
      
    Case EVGN_POSXSAMPLE          ' Request to Position External Sample

    Case EVGN_MSCOMFLD            ' MS1100if COM Failure: Possible Hard Reset or Similar Event
    
    Case EVGN_MSSRVDIED           ' MS11srv OCX ME_CODE_DEAD State entered
      
    Case EVGN_TRAYSTCHGD          ' TrayStatus has changed
      ' Check if tray has been positioned for product sample scan
      trayStatus = MS11srv.trayStatus And &H300
      
      Select Case (trayStatus)
        Case DRWR_POS_OPENED
          unity_main.m_lastDrwrPos = trayStatus

        Case DRWR_POS_CLOSED
          ' Check if internal reference performed on demand
          If (unity_main.m_backFreq = REF_FREQ_ON_DEMAND) Or (unity_main.repcounter <> 0) Then
            If (unity_main.m_lastDrwrPos = DRWR_POS_OPENED) Then
              If (tmr_sample.enabled = False) And (tmr_ref.enabled = False) Then
                unity_main.m_smplAutoScan = True
              End If
            End If
          End If
          
          unity_main.m_lastDrwrPos = trayStatus
      End Select

      Exit Sub
    
    Case Else
'    Case EVGN_INITFLD             ' Initialization Sequence Failed
'    Case EVGN_TRAYRDY             ' Tray Passed Initialization and Ready for Config/Use
'    Case EVGN_SCNRRDY             ' Scanner Passed Initialization and Ready for Config/Use
'    Case EVGN_HEARTBEAT           ' 1 Second (thereabouts) Heart Beat from OCX Timer
'    Case EVGN_MERRORUPD           ' xxxErrorCode and MasterErrors have been updated during non-user event, e.g. Timer.
      Exit Sub
  End Select
  
  ' Save event info and queue for processing later if any timer enabled
  If ((tmr_all.enabled = True) Or (tmr_ref.enabled = True) Or (tmr_sample.enabled = True)) Then
    gnEvent.noticeType = notice
    gnEvent.trayStatus = MS11srv.trayStatus
    gnEvent.masterErrCode = MS11srv.MasterErrors
    gnEvent.firstErrCode = MS11srv.FirstErrorCode
    gnEvent.lastErrCode = MS11srv.LastErrorCode
    gnEvent.trayStatus = m_trayStatus
    
    Get_MS11_Error_Codes
    gnEvent.masterErrCode = MS11MasterErrors
    gnEvent.firstErrCode = MS11FirstErrorCode
    gnEvent.lastErrCode = MS11LastErrorCode
  
    If (m_ms11srvGNEventQ.Count = 0) Then
      m_ms11srvGNEventQ.Add gnEvent
    Else
      m_ms11srvGNEventQ.Add gnEvent, , , m_ms11srvGNEventQ.Count
    End If
  End If
End Sub
#End If

#If MS11srv Then
Private Sub MS11srv_ScanProgress(ByVal percent As Long)

  ' Check if performing batch scan
  If (unity_main.m_batchRunFlg = True) Then
    If (percent <> 0) Then
      If (unity_main.m_scanState = SS_STOP) Or (unity_main.m_scanState = SS_PAUSE) Then
        unity_main.m_scanState = SS_START
      End If
    End If
    
    ' Update scan progress bar
    frm_batchRun.scanProgress.percent = percent
  Else
    If (percent <> 0) Then
      frm_status.lbl_statusCmd.ForeColor = RGB(0, 128, 0)  ' dark green

      If (unity_main.m_scanState = SS_STOP) Or (unity_main.m_scanState = SS_PAUSE) Then
        unity_main.m_scanState = SS_START
        frm_status.cmd_retryScan.Visible = False
        frm_status.cmd_exitScan.Visible = False
        frm_status.cmd_resumeScan.Visible = False
        frm_status.cmd_abortScan.Visible = True
'        frm_status.cmd_pauseScan.Visible = True    ' Comment out for now until firmware updated
      
        Select Case (MS11CfgData.devID)
          Case DTID_DRAWER0            ' SS2200/SS2400 standard drawer system
            frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status63", "Scanning has started, please do NOT move sample drawer")
    
          Case DTID_TOPWIND0           ' Top window w/out internal reflectance
            ' Check if internal reference scan
            If (unity_main.m_scanDataType = SDT_EXTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODEXTREF) Then
              frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status64", "Scanning has started, please do NOT move reference sample")
            Else    ' product scan
              frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status65", "Scanning has started, please do NOT move product sample")
            End If
    
          Case DTID_DRAWER1            ' SS2200/SS2400 drawer w/out stepper system
            frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status63", "Scanning has started, please do NOT move sample drawer")
      
          Case DTID_TOPWIND1           ' Top window with internal reflectance
            ' Check if internal reference scan
            If (unity_main.m_scanDataType = SDT_INTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODINTREF) Or _
               (unity_main.m_scanDataType = SDT_EXTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODEXTREF) Then
              ' Check if internal reference
              If (unity_main.m_bType = "internal") Then
                frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status66", "Scanning has started")
              Else
                frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status64", "Scanning has started, please do NOT move reference sample")
              End If
            Else    ' product scan
              frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status65", "Scanning has started, please do NOT move product sample")
            End If
        End Select
      End If
    End If
    
    ' Update scan progress bar
    frm_status.scanProgress.percent = percent
  End If
End Sub
#End If

 

#If SSRCS Then
Private Sub SSRCSClient_ASCmdCompleteEvnt(ByVal cmdId As Long, cmdIdTxt As String, ByVal commStat As Long, rspData As String)

  frm_batchRun.process_auto_sampler_response cmdId, cmdIdTxt, commStat, rspData
End Sub
#End If

#If SSRCS Then
Private Sub SSRCSClient_ConnectStatusEvnt(ByVal Status As Integer)
  
  If (Status = 0) Then
    m_ssrcsConnected = False
    img_ssrcsConnect.BackColor = vbRed
    unity_main.cmd_ssrcsConnect.Visible = True
  Else
    m_ssrcsConnected = True
    img_ssrcsConnect.BackColor = Frame1.BackColor
    unity_main.cmd_ssrcsConnect.Visible = False
  End If
End Sub
#End If

#If SSRCS Then
Private Sub SSRCSClient_GeneralNoticeEvnt(ByVal notice As Long)
  Dim gnEvent As New clsMS11srvGNEvent
  Dim trayStatus As Long
    
  Select Case (notice)
    Case EVGN_INITDONE            ' Initialization Sequence Passed Successfully
      MS11Initialized = True
      Exit Sub
      
    Case EVGN_TRAYFLD             ' Tray Failure
   
    Case EVGN_SCNRFLD             ' Scanner Failure
      
    Case EVGN_SCANABRTD           ' Current Scan has been Aborted
      
    Case EVGN_SCANSTOPD           ' Scanner Stopped Scanning; Resume can be initiated
      
    Case EVGN_IMPOSERR            ' During Scan/Tray Engine Running, Impossible State Transitions occurred
      
    Case EVGN_REFRNCTMO           ' Reference has Timed Out and is now Expired
      ' Check if internal reference performed on demand
      If (unity_main.m_backFreq = REF_FREQ_ON_DEMAND) Then
        unity_main.m_intRefAutoScan = True
        unity_main.lbl_backdate.Caption = MLSupport.GSS("unity_main", "lbl_backdate_int2", "Internal Reference has Expired")
      
        If ((unity_main.m_batchRunFlg = False) And (unity_main.m_intRefCalFlg = False) And (unity_main.repcounter = 0)) Then
          unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status5", "Reference Required")
        End If
      Else
        ' Stop instrument's reference timer
        SSRCSClientError = SSRCSClient.SetRefTimeout(0)
      End If
      
      Exit Sub
      
    Case EVGN_SMPLSCNOK           ' Sample Data Scan Successfully Completed
      unity_main.m_scanDblTimestamp = Now
      unity_main.m_scanTimestamp = CDate(m_scanDblTimestamp)
      
    Case EVGN_SMPLSCNBAD          ' Sample Data Scan Collected but Bad (Clipped or Failed PPT Qualification)
      
    Case EVGN_REFSCNOK            ' Reference Data Scan Successfully Completed
      unity_main.m_scanDblTimestamp = Now
      unity_main.m_scanTimestamp = CDate(m_scanDblTimestamp)
      
    Case EVGN_REFSCNBAD           ' Reference Data Scan Collected but Bad (Clipped or Failed PPT Qualification)
      
    Case EVGN_OPENDRWR            ' Request to OPEN Sample Drawer
      frm_status.cmd_exitScan.Visible = False
      frm_status.cmd_resumeScan.Visible = False
      frm_status.cmd_retryScan.Visible = False
      frm_status.cmd_abortScan.Visible = True
      
    Case EVGN_CLOSEDRWR           ' Request to CLOSE Sample Drawer
      frm_status.cmd_exitScan.Visible = False
      frm_status.cmd_resumeScan.Visible = False
      frm_status.cmd_retryScan.Visible = False
      frm_status.cmd_abortScan.Visible = True
      
    Case EVGN_POSXREFRNC          ' Request to Position External Reference
      
    Case EVGN_POSXSAMPLE          ' Request to Position External Sample

    Case EVGN_MSCOMFLD            ' MS1100if COM Failure: Possible Hard Reset or Similar Event
    
    Case EVGN_MSSRVDIED           ' MS11srv OCX ME_CODE_DEAD State entered
      
    Case EVGN_TRAYSTCHGD          ' TrayStatus has changed
      ' Check if tray has been positioned for product sample scan
      SSRCSClientError = SSRCSClient.GetTrayStatus(m_trayStatus)
      trayStatus = m_trayStatus And &H300
      
      Select Case (trayStatus)
        Case DRWR_POS_OPENED
          unity_main.m_lastDrwrPos = trayStatus

        Case DRWR_POS_CLOSED
          ' Check if internal reference performed on demand
          If (unity_main.m_backFreq = REF_FREQ_ON_DEMAND) Or (unity_main.repcounter <> 0) Then
            If (unity_main.m_lastDrwrPos = DRWR_POS_OPENED) Then
              If (tmr_sample.enabled = False) And (tmr_ref.enabled = False) Then
                unity_main.m_smplAutoScan = True
              End If
            End If
          End If
          
          unity_main.m_lastDrwrPos = trayStatus
      End Select

      Exit Sub
    
    Case Else
'    Case EVGN_INITFLD             ' Initialization Sequence Failed
'    Case EVGN_TRAYRDY             ' Tray Passed Initialization and Ready for Config/Use
'    Case EVGN_SCNRRDY             ' Scanner Passed Initialization and Ready for Config/Use
'    Case EVGN_HEARTBEAT           ' 1 Second (thereabouts) Heart Beat from OCX Timer
'    Case EVGN_MERRORUPD           ' xxxErrorCode and MasterErrors have been updated during non-user event, e.g. Timer.
      Exit Sub
  End Select
  
  ' Save event info and queue for processing later if any timer enabled
  If ((tmr_all.enabled = True) Or (tmr_ref.enabled = True) Or (tmr_sample.enabled = True)) Then
    gnEvent.noticeType = notice
    
    SSRCSClientError = SSRCSClient.GetTrayStatus(m_trayStatus)
    gnEvent.trayStatus = m_trayStatus
    
    Get_MS11_Error_Codes
    gnEvent.masterErrCode = MS11MasterErrors
    gnEvent.firstErrCode = MS11FirstErrorCode
    gnEvent.lastErrCode = MS11LastErrorCode
  
    If (m_ms11srvGNEventQ.Count = 0) Then
      m_ms11srvGNEventQ.Add gnEvent
    Else
      m_ms11srvGNEventQ.Add gnEvent, , , m_ms11srvGNEventQ.Count
    End If
  End If
End Sub
#End If

#If SSRCS Then
Private Sub SSRCSClient_ScanProgressEvnt(ByVal percent As Long)

  ' Check if performing batch scan
  If (unity_main.m_batchRunFlg = True) Then
    If (percent <> 0) Then
      If (unity_main.m_scanState = SS_STOP) Or (unity_main.m_scanState = SS_PAUSE) Then
        unity_main.m_scanState = SS_START
      End If
    End If
    
    ' Update scan progress bar
    frm_batchRun.scanProgress.percent = percent
  Else
    ' Check if performing internal reference calibration function
    If (unity_main.m_intRefCalFlg = True) Then
       frm_intRefCalMgmt.update_scan_progress percent
    Else
      If (percent <> 0) Then
        frm_status.lbl_statusCmd.ForeColor = RGB(0, 128, 0)  ' dark green

        If (unity_main.m_scanState = SS_STOP) Or (unity_main.m_scanState = SS_PAUSE) Then
          unity_main.m_scanState = SS_START
          frm_status.cmd_retryScan.Visible = False
          frm_status.cmd_exitScan.Visible = False
          frm_status.cmd_resumeScan.Visible = False
          frm_status.cmd_abortScan.Visible = True
'          frm_status.cmd_pauseScan.Visible = True    ' Comment out for now until firmware updated
      
          Select Case (MS11CfgData.devID)
            Case DTID_DRAWER0            ' SS2200/SS2400 standard drawer system
              frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status63", "Scanning has started, please do NOT move sample drawer")
    
            Case DTID_TOPWIND0           ' Top window w/out internal reflectance
              ' Check if internal reference scan
              If (unity_main.m_scanDataType = SDT_EXTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODEXTREF) Then
                frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status64", "Scanning has started, please do NOT move reference sample")
              Else    ' product scan
                frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status65", "Scanning has started, please do NOT move product sample")
              End If
    
            Case DTID_DRAWER1            ' SS2200/SS2400 drawer w/out stepper system
              frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status63", "Scanning has started, please do NOT move sample drawer")
      
            Case DTID_TOPWIND1           ' Top window with internal reflectance
              ' Check if internal reference scan
              If (unity_main.m_scanDataType = SDT_INTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODINTREF) Or _
                 (unity_main.m_scanDataType = SDT_EXTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODEXTREF) Then
                ' Check if internal reference
                If (unity_main.m_bType = "internal") Then
                  frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status66", "Scanning has started")
                Else
                  frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status64", "Scanning has started, please do NOT move reference sample")
                End If
              Else    ' product scan
                frm_status.lbl_statusCmd.Caption = MLSupport.GSS("OperStatus", "status65", "Scanning has started, please do NOT move product sample")
              End If
          End Select
        End If
      End If
    
      ' Update scan progress bar
      frm_status.scanProgress.percent = percent
    End If
  End If
End Sub
#End If

Private Sub tmr_all_Timer()
  Dim uniMsg As String
  
  ' Check if to change product via remote client application
  If (unity_main.m_remoteProdSelect = 1) Then
    unity_main.m_remoteProdSelect = 2
    
    If (frm_POG.remote_product_selection(unity_main.m_remoteProdName, "Remote application product selection") = False) Then
      send_current_product unity_main.m_remoteSocketIndx, InfoStarStatusCodes.InvProductNameStat
    End If
    
    unity_main.m_remoteProdSelect = 0
  Else
    ' Check if to change product via LIMS application
    If (unity_main.run_min_gui = True) And (unity_main.remoteproduct <> 0) Then
      frm_POG.checknewprod
    End If
  End If
  
#If SSTAR Then
  ' Check if to display check internal reference reminder
  If (m_intRefVerReminderFlg = True) Then
    disp_verify_ref_button True
    CWrap.ShowMessageBoxW MLSupport.GSS("OperStatus", "status108", "Time to check internal reference. Press 'Verify Ref' button to access utility."), vbOKOnly, MLSupport.GSS("Headers", "remind", "Reminder")
  End If
#End If
  
  ' Check if internal reference qualification scan required
  If (unity_main.m_intRefPPTScan = True) And (unity_main.repcounter = 0) Then
    ' Unhide "Ref" button if current reference from offline file
    unity_main.cmd_ref.Visible = True
    uniMsg = MLSupport.GSS("OperStatus", "status67", "Auto internal reference qualification scan")
    Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Auto internal reference qualification scan", uniMsg)
    lbl_opStatus.Caption = MLSupport.GGS_Params("unity_main.statMsg32", "Auto Internal Reference Qualification Scan (Wavelength Range %1-%2)", CStr(unity_main.m_intRefPPTStartWvln), CStr(unity_main.m_intRefPPTEndWvln))
    unity_main.m_scanDataType = SDT_INTREFPPT
    perform_ref_scan
  Else
    ' Check if external reference qualification scan required
    If (unity_main.m_extRefPPTScan = True) And (unity_main.repcounter = 0) Then
      ' Unhide "Ref" button if current reference from offline file
      unity_main.cmd_ref.Visible = True
      uniMsg = MLSupport.GSS("OperStatus", "status82", "Auto external reference qualification scan")
      Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Auto external reference qualification scan", uniMsg)
      lbl_opStatus.Caption = MLSupport.GGS_Params("unity_main.statMsg34", "Auto External Reference Qualification Scan (Wavelength Range %1-%2)", CStr(unity_main.m_extRefPPTStartWvln), CStr(unity_main.m_extRefPPTEndWvln))
      unity_main.m_scanDataType = SDT_EXTREFPPT
      perform_ref_scan
    Else
      ' Check if product internal reference scan required
      If (unity_main.m_intRefAutoScan = True) And (unity_main.repcounter = 0) Then
        uniMsg = MLSupport.GSS("OperStatus", "status68", "Auto internal reference scan")
        Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Auto internal reference scan", uniMsg)
        unity_main.m_scanDataType = SDT_PRODINTREF
        lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status11", "Auto Internal Reference Scan")
        perform_ref_scan
      Else
        ' Check if product external reference scan required
        If (unity_main.m_extRefAutoScan = True) And (unity_main.repcounter = 0) Then
          uniMsg = MLSupport.GSS("OperStatus", "status83", "Auto external reference scan")
          Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Auto external reference scan", uniMsg)
          unity_main.m_scanDataType = SDT_PRODEXTREF
          lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status84", "Auto External Reference Scan")
          perform_ref_scan
        Else
          ' Check if product reference scan requested remotely
          If (unity_main.m_remoteRefScan = True) And (unity_main.repcounter = 0) Then
            If (unity_main.m_bType = "internal") Then
              unity_main.m_intRefAutoScan = True
              uniMsg = MLSupport.GSS("OperStatus", "status69", "Internal reference scan requested remotely")
              Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Internal reference scan requested remotely", uniMsg)
              unity_main.m_scanDataType = SDT_PRODINTREF
              lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status81", "Remote Internal Reference Scan")
            Else
              unity_main.m_extRefAutoScan = True
              uniMsg = MLSupport.GSS("OperStatus", "status85", "External reference scan requested remotely")
              Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "External reference scan requested remotely", uniMsg)
              unity_main.m_scanDataType = SDT_PRODEXTREF
              lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status86", "Remote External Reference Scan")
            End If
            
            perform_ref_scan
          Else
            ' Check if product internal reference scan requested by user
            If (unity_main.m_intRefManualScan = True) Then
              unity_main.m_intRefManualScan = False
              unity_main.m_intRefAutoScan = True
              uniMsg = MLSupport.GSS("OperStatus", "status70", "User requested internal reference scan")
              Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "User requested internal reference scan", uniMsg)
        
              If (unity_main.m_intRefPPTScan = True) Then
                lbl_opStatus.Caption = MLSupport.GGS_Params("unity_main.statMsg33", "Manual Internal Reference Qualification Scan (Wavelength Range %1-%2)", CStr(unity_main.m_intRefPPTStartWvln), CStr(unity_main.m_intRefPPTEndWvln))
                unity_main.m_scanDataType = SDT_INTREFPPT
              Else
                unity_main.m_scanDataType = SDT_PRODINTREF
                lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status12", "Manual Internal Reference Scan")
              End If
        
              perform_ref_scan
            Else
              ' Check if product external reference scan requested by user
              If (unity_main.m_extRefManualScan = True) Then
                unity_main.m_extRefManualScan = False
                unity_main.m_extRefAutoScan = True
                uniMsg = MLSupport.GSS("OperStatus", "status98", "User requested external reference scan")
                Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "User requested external reference scan", uniMsg)
        
                If (unity_main.m_extRefPPTScan = True) Then
                  lbl_opStatus.Caption = MLSupport.GGS_Params("unity_main.statMsg35", "Manual External Reference Qualification Scan (Wavelength Range %1-%2)", CStr(unity_main.m_extRefPPTStartWvln), CStr(unity_main.m_extRefPPTEndWvln))
                  unity_main.m_scanDataType = SDT_EXTREFPPT
                Else
                  unity_main.m_scanDataType = SDT_PRODEXTREF
                  lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status87", "Manual External Reference Scan")
                End If
        
                perform_ref_scan
              Else
                ' Check if product sample qualification scan required
                If (unity_main.m_smplPPTScan = True) Then
                  uniMsg = MLSupport.GSS("OperStatus", "status71", "Product sample qualification scan")
                  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Product sample qualification scan", uniMsg)
                  lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status13", "Sample Qualification Scan")
                  unity_main.m_scanDataType = SDT_PRODPPT
                  Call perform_smpl_scan(m_smplPPTScan)  ' Note: do not add unity_main. prefix to passed variable
                Else
                  ' Check if auto product sample scan required
                  If (unity_main.m_smplAutoScan = True) Then
                    uniMsg = MLSupport.GSS("OperStatus", "status72", "Auto product sample scan")
                    Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Auto product sample scan", uniMsg)
                    unity_main.m_scanDataType = SDT_PRODSMPL
                    lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status14", "Auto Sample Scan")
                    Call perform_smpl_scan(m_smplAutoScan)  ' Note: do not add unity_main. prefix to passed variable
                  Else
                    ' Check if auto product sample scan requested remotely
                    If (unity_main.m_remoteSmplScan = True) Then
                      uniMsg = MLSupport.GSS("OperStatus", "status73", "Product sample scan requested remotely")
                      Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Product sample scan requested remotely", uniMsg)
                      unity_main.m_scanDataType = SDT_PRODSMPL
                      lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status15", "Remote Sample Scan")
                      Call perform_smpl_scan(m_remoteSmplScan)  ' Note: do not add unity_main. prefix to passed variable
                    Else
                      ' Check if product sample scan requested by user
                      If (unity_main.m_smplManualScan = True) Then
                        uniMsg = MLSupport.GSS("OperStatus", "status74", "User requested product sample scan")
                        Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "User requested product sample scan", uniMsg)
                        unity_main.m_scanDataType = SDT_PRODSMPL
                        lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status16", "Manual Sample Scan")
                        Call perform_smpl_scan(m_smplManualScan)  ' Note: do not add unity_main. prefix to passed variable
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    End If
  End If
End Sub

#If SSTAR Then
Private Sub tmr_ref_Timer()
  Dim uniMsg As String
  
  Select Case (unity_main.m_scanTmrState)
    Case STS_ABORT                ' scan was aborted due to error or user
      ' Display buttons and wait for user decision
      frm_status.cmd_abortScan.Visible = False
      frm_status.cmd_pauseScan.Visible = False
      frm_status.cmd_resumeScan.Visible = False
      frm_status.cmd_exitScan.Visible = True
      frm_status.cmd_retryScan.Visible = True
      unity_main.m_scanState = SS_STOP
    
    Case STS_SETUP                ' setup and start scan
      If (unity_main.setup_scan = False) Then
        GoTo ReportScanError
      Else
        If (unity_main.m_scanDataType = SDT_INTREFPPT) Then
          uniMsg = MLSupport.GSS("OperStatus", "status75", "Starting internal reference qualification scan")
          Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Starting internal reference qualification scan", uniMsg)
        Else
          If (unity_main.m_scanDataType = SDT_EXTREFPPT) Then
            uniMsg = MLSupport.GSS("OperStatus", "status89", "Starting external reference qualification scan")
            Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Starting external reference qualification scan", uniMsg)
          Else
            If (unity_main.m_scanDataType = SDT_PRODINTREF) Then
              uniMsg = MLSupport.GSS("OperStatus", "status76", "Starting product internal reference scan")
              Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Starting product internal reference scan", uniMsg)
            Else
              uniMsg = MLSupport.GSS("OperStatus", "status97", "Starting product external reference scan")
              Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Starting product external reference scan", uniMsg)
            End If
          End If
        End If
        
        ' Check if external reference qualification or product scan
        If (unity_main.m_scanDataType = SDT_EXTREFPPT) Or (unity_main.m_scanDataType = SDT_PRODEXTREF) Then
          unity_main.setup_ext_ref_pos
        Else
          If (unity_main.start_scan = False) Then
            GoTo ReportScanError
          Else
            unity_main.m_scanTmrState = STS_WAIT_CMP
          End If
        End If
      End If
  
    Case STS_WAIT_POS_EXT_REF     ' waiting for user to position external reference
      Select Case (unity_main.m_extRefPosition)
        Case 0    ' user pressed cancel button
          unity_main.m_ansiErrMsg = "Scan cancel button pressed by user"
          unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status51", "Scan cancel button pressed by user")
          GoTo ReportScanError
        
        Case 1    ' user pressed start button
          If (unity_main.start_scan = False) Then
            GoTo ReportScanError
          Else
            unity_main.m_scanTmrState = STS_WAIT_CMP
          End If
      End Select
        
    Case STS_WAIT_CMP             ' waiting for scan to complete
      If (unity_main.check_scan_cmpl = False) Then
        GoTo ReportScanError
      End If
    
    Case STS_GET_SCAN             ' get scan data
      ' Get reference scan data
      If (unity_main.get_scan_data = True) Then
        ' Check if performing internal reference calibration function
        If (unity_main.m_intRefCalFlg = True) Then
          tmr_ref.enabled = False
          unity_main.m_scanTmrState = STS_COMPLETED
          Exit Sub
        Else
          unity_main.m_scanTmrState = STS_SAVE_SCAN
        End If
      Else
        GoTo ReportScanError
      End If
    
    Case STS_SAVE_SCAN            ' save scan data
      If (unity_main.save_scan_data = True) Then
        ' Check if completed reference qualification scan
        If (unity_main.m_scanDataType = SDT_INTREFPPT) Then
          unity_main.m_intRefPPTScan = False
          unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status17", "Internal Reference Qualification Data Collected")
        Else
          If (unity_main.m_scanDataType = SDT_EXTREFPPT) Then
            unity_main.m_extRefPPTScan = False
            unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status91", "External Reference Qualification Data Collected")
                      
            ' Update any new working reference timeout value
            unity_main.m_extRefTimeoutSecs = unity_main.m_extRefTimeout * 60
            unity_main.m_extRefTimerIgnore = False
                      
            ' Reset external reference PPT variables in case user requested new PPT collection scan
            unity_main.m_extRefPPTAdapterType = unity_main.m_extRefAdapterType
            unity_main.m_extRefPPTAdaptIndx = unity_main.m_extRefAdaptIndx
            unity_main.m_extRefPPTEndWvln = unity_main.m_extRefEndWvln
            unity_main.m_extRefPPTFileName = unity_main.m_extRefFileName
            unity_main.m_extRefPPTFileSetup = False
            unity_main.m_extRefPPTMultiCupIndx = unity_main.m_extRefMultiCupIndx
            unity_main.m_extRefPPTMultiCupType = unity_main.m_extRefMultiCupType
            unity_main.m_extRefPPTNScans = unity_main.m_extRefNScans
            unity_main.m_extRefPPTRotateDir = unity_main.m_extRefRotateDir
            unity_main.m_extRefPPTRotateIndexSteps = unity_main.m_extRefRotateIndexSteps
            unity_main.m_extRefPPTRotateMoveMode = unity_main.m_extRefRotateMoveMode
            unity_main.m_extRefPPTRotateSpeed = unity_main.m_extRefRotateSpeed
            unity_main.m_extRefPPTRotateStepSteps = unity_main.m_extRefRotateStepSteps
            unity_main.m_extRefPPTStartWvln = unity_main.m_extRefStartWvln
            unity_main.m_extRefPPTTrayNum = unity_main.m_extRefTrayNum
          Else
            If (unity_main.m_scanDataType = SDT_PRODINTREF) Then
              unity_main.m_remoteRefScan = False
              unity_main.m_intRefAutoScan = False
              unity_main.m_intRefManualScan = False
              unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status18", "Product Internal Reference Data Collected")
            Else
              unity_main.m_remoteRefScan = False
              unity_main.m_extRefAutoScan = False
              unity_main.m_extRefManualScan = False
              unity_main.m_extRefTimeoutSecs = unity_main.m_extRefTimeout * 60
              unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status92", "Product External Reference Data Collected")
            End If
          End If
        End If
        
        ' Hide "Ref" button if reference from offline file
        If (unity_main.m_bType = "file") Then
          unity_main.cmd_ref.Visible = False
        End If

        unity_main.m_scanTmrState = STS_PLOT_SCAN
        Unload frm_status
      Else
        GoTo AbortScan
      End If
      
    Case STS_PLOT_SCAN            ' plot scan data
      unity_main.plot_spectrum
      tmr_ref.enabled = False
     
      unity_main.m_scanTmrState = STS_COMPLETED
     
      ' Check if performing internal reference on demand
      If (unity_main.m_backFreq = REF_FREQ_ON_DEMAND) Then
        If (unity_main.m_batchRunFlg = False) Then
          ' Restart auto-loop timer
          tmr_all.enabled = True
          uniMsg = MLSupport.GSS("OperStatus", "status77", "Reference sample completed")
          Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Reference sample completed", uniMsg)
        End If
      
        ' Check if product reference required
        If (unity_main.m_intRefAutoScan = True) Or (unity_main.m_extRefAutoScan = True) Then
          unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status5", "Reference Required")
        Else
          If (unity_main.m_remoteRefScan = True) Then
            unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status9", "Reference Requested Remotely")
          Else
            unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status10", "Ready for New Sample")
            cmd_sample.enabled = True
           If cmd_sample.Visible And cmd_sample.enabled Then
             'cmd_sample.SetFocus
            End If
          End If
        End If
        
      End If
  End Select
  
  Exit Sub
  
ReportScanError:
  ' Report error codes
  unity_main.lbl_opStatus.Caption = unity_main.m_uniErrMsg
  Call frm_status.report_error_codes(unity_main.m_ansiErrMsg, unity_main.m_uniErrMsg)

AbortScan:
  unity_main.m_scanTmrState = STS_ABORT
End Sub 'ref_timer
#End If

#If SSTAR Then
Private Sub tmr_sample_Timer()
  Dim rc As Boolean
  Dim uniMsg As String
  
  Select Case (unity_main.m_scanTmrState)
    Case STS_ABORT                ' scan was aborted due to error or user
      ' Display buttons and wait for user decision
      frm_status.cmd_abortScan.Visible = False
      frm_status.cmd_pauseScan.Visible = False
      frm_status.cmd_resumeScan.Visible = False
      frm_status.cmd_exitScan.Visible = True
      frm_status.cmd_retryScan.Visible = True
      unity_main.m_scanState = SS_STOP
      
    Case STS_SETUP                ' setup and start scan
      unity_main.lbl_miltime.Caption = ""
      
      If (unity_main.setup_scan = False) Then
        GoTo ReportScanError
      Else
        If (unity_main.m_scanDataType = SDT_PRODPPT) Then
          uniMsg = MLSupport.GGS_Params("unity_main.statMsg28", "Starting product qualification scan %1 of %2", CStr(unity_main.repcounter), CStr(unity_main.m_smplRepacks))
          Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Starting product qualification scan " & unity_main.repcounter & " of " & unity_main.m_smplRepacks), uniMsg)
        Else
          uniMsg = MLSupport.GGS_Params("unity_main.statMsg29", "Starting product sample scan %1 of %2", CStr(unity_main.repcounter), CStr(unity_main.m_smplRepacks))
          Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Starting product sample scan " & unity_main.repcounter & " of " & unity_main.m_smplRepacks), uniMsg)
        End If
        
        If (unity_main.start_scan = False) Then
          GoTo ReportScanError
        Else
          unity_main.m_scanTmrState = STS_WAIT_CMP
        End If
      End If
  
    Case STS_WAIT_CMP             ' waiting for scan to complete
      If (unity_main.check_scan_cmpl = False) Then
        GoTo ReportScanError
      End If
    
    Case STS_GET_SCAN             ' get scan data
      ' Get sample scan data
      If (unity_main.get_scan_data = True) Then
        ' Check if performing internal reference calibration function
        If (unity_main.m_intRefCalFlg = True) Then
          tmr_sample.enabled = False
          unity_main.m_scanTmrState = STS_COMPLETED
          Exit Sub
        Else
          ' Check if to smooth spectrum data
          If (unity_main.m_enableSmooth = True) Then
            If (unity_main.smooth_scan_data = False) Then
              GoTo ReportScanError
            End If
          End If
        
          unity_main.m_scanTmrState = STS_SAVE_SCAN
        End If
      Else
        GoTo ReportScanError
      End If
    
    Case STS_SAVE_SCAN            ' save scan data
      If (unity_main.save_scan_data = True) Then
        ' Check if completed product qualification or sample scan
        If (unity_main.m_scanDataType = SDT_PRODPPT) Or (unity_main.m_scanDataType = SDT_PRODSMPL) Then
          ' Check if completed product qualification scan
          If (unity_main.m_scanDataType = SDT_PRODPPT) Then
            unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status19", "Product Qualification Sample Collected")
            unity_main.m_scanDataType = SDT_PRODPPTABS
          Else
            unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status20", "Product Sample Collected")
            unity_main.m_scanDataType = SDT_PRODSMPLABS
          End If
        
          ' Calculate and save absorbance data
          If (unity_main.calc_prod_abs = False) Then
            GoTo ReportScanError
          End If
          
          ' Sum absorbance data to be averaged later if performing repacks
          If (unity_main.m_scanDataType = SDT_PRODSMPLABS) And (unity_main.m_smplRepacks > 1) Then
            unity_main.sum_prod_avg_abs
          Else
            ' Check if to standardized the spectrum
            If (unity_main.m_prdModelType = True) Then
              ' Check if to pass treated spectrum
              If (unity_main.m_enableTreatment = True) Then
                rc = unity_main.standardize_spectrum(ProdTreatAbsYVals)
              Else
                rc = unity_main.standardize_spectrum(ProdAbsYVals)
              End If
              
              ' Check if any error standardizing spectrum
              If (rc = False) Then
                GoTo ReportScanError
              End If
            End If
          End If
        Else
          ' Plot spectrum
          unity_main.m_scanTmrState = STS_PLOT_SCAN
          Unload frm_status
        End If
      Else
        GoTo AbortScan
      End If
      
    Case STS_PLOT_SCAN            ' plot scan data
      unity_main.plot_spectrum
      unity_main.m_scanTmrState = STS_PROC_DATA
    
    Case STS_PROC_DATA            ' process absorbance data
      ' Check if completed product sample scan
      If (unity_main.m_scanDataType = SDT_PRODSMPLABS) Then
        unity_main.lbl_date.Caption = CStr(Date)
        unity_main.lbl_time.Caption = CStr(Time)
        unity_main.lbl_miltime.Caption = Now      ' for LIMS output
    
        ' Check if to perform predictions
        If (unity_main.m_makePred = "yes") Then
          unity_main.do_pred
        
          If (unity_main.pukedonpred = True) Then
            unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status21", "Model Failure")
            GoTo MODEL_FAIL
          Else
            ' Check if performing repacks
            If (unity_main.m_smplRepacks > 1) Then
              ' Check if have finished replicate scans
              unity_main.repcounter = unity_main.repcounter + 1
        
              If (unity_main.repcounter <= unity_main.m_smplRepacks) Then
                lbl_opStatus.Caption = MLSupport.GGS_Params("unity_main.statMsg31", "Insert Repack %1 of %2", CStr(unity_main.repcounter), CStr(unity_main.m_smplRepacks))
                tmr_sample.enabled = False
                cmd_sample.enabled = True
                ' Restart auto-loop
                tmr_all.enabled = True
                Exit Sub
              Else
                ' Calculate averaged repack spectrum, save and average predictions
                unity_main.calc_prod_avg_abs
                unity_main.m_scanDataType = SDT_PRODAVGABS
                
                ' Check if to standardized the spectrum
                If (unity_main.m_prdModelType = True) Then
                  ' Check if to pass treated spectrum
                  If (unity_main.m_enableTreatment = True) Then
                    rc = unity_main.standardize_spectrum(ProdTreatAvgAbsYVals)
                  Else
                    rc = unity_main.standardize_spectrum(ProdAvgAbsYVals)
                  End If
              
                  ' Check if any error standardizing spectrum
                  If (rc = False) Then
                    GoTo ReportScanError
                  End If
                End If
                
                unity_main.save_scan_data
                frm_collect.calcavgpreds
                
                ' Check if spectrum data treated
                If (unity_main.m_enableTreatment = True) Then
                  ' Do prediction on treated averaged repack spectrum for alarm purposes
                  ProdTreatAbsYVals = ProdTreatAvgAbsYVals
                Else
                  ' Do prediction on averaged repack spectrum for alarm purposes
                  ProdAbsYVals = ProdAvgAbsYVals
                End If
                
                unity_main.do_pred
                
                If (unity_main.pukedonpred = True) Then
                  unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status21", "Model Failure")
                  GoTo MODEL_FAIL
                End If
              End If
            End If
          
            If (unity_main.m_valueBound > 0) Then
              unity_main.checkbounds
            End If
    
            unity_main.chksigfigs
          
            unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status22", "Checking Alarm Limits")
            frm_collect.checkalarms
          End If
        Else      ' no predictions requested
          unity_main.clear_pred_results
          
          ' Check if performing repacks
          If (unity_main.m_smplRepacks > 1) Then
            ' Check if have finished replicate scans
            unity_main.repcounter = unity_main.repcounter + 1
        
            If (unity_main.repcounter <= unity_main.m_smplRepacks) Then
              lbl_opStatus.Caption = MLSupport.GGS_Params("unity_main.statMsg31", "Insert Repack %1 of %2", CStr(unity_main.repcounter), CStr(unity_main.m_smplRepacks))
              tmr_sample.enabled = False
              cmd_sample.enabled = True
              ' Restart auto-loop
              tmr_all.enabled = True
              Exit Sub
            Else
              unity_main.calc_prod_avg_abs
              unity_main.m_scanDataType = SDT_PRODAVGABS
              unity_main.save_scan_data
            End If
          End If
          
          unity_main.disp_no_pred_results
        End If
        
        ' Save report if requested
        If (unity_main.m_savePredictions = True) Then
          unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status23", "Writing Values to Log File")
          unity_main.write_report_file
        End If

        ' Save csv if requested
        If (unity_main.m_saveCSV = True) Then
          unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status24", "Writing Values to CSV File")
          frm_csvCfg.write_csv_report
        End If
            
        ' Save dynamic report if requested
        If (unity_main.m_saveDynRpt = True) Then
          unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status107", "Writing Values to Dynamic Report File")
          frm_dynRptCfg.write_dynamic_report
        End If
            
        unity_main.add_last50
            
        ' Check if configured for LIMS output
        If (unity_main.m_doLIMS = 1) Then
          ' Check if operator prompt not required
          If (unity_main.pogacceptreject = 0) Then
            frm_POG.writepog
            frm_POG.writeitout
          Else
            frm_sendlims.Show 1
            
            ' Check if operator accepted LIMS output
            If (unity_main.m_acceptLims = True) Then
              frm_POG.writepog
              frm_POG.writeitout
            End If
          End If
        End If

        ' Check if configured for write ticket output
        If (unity_main.m_writeTkt = 1) Then
          frm_ticket.writeticket
        End If
            
        ' Save scan data in different format only if no prediction or prediction worked
        If (unity_main.pukedonpred = False) Then
          unity_main.save_scan_diff_format
        End If
            
        unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status25", "Sample Completed")
            
        If (unity_main.m_intRefPPTScan = True) Or (unity_main.m_extRefPPTScan = True) Then
          unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status6", "Reference Qualification Required")
        Else
          If (unity_main.m_intRefAutoScan = True) Or (unity_main.m_extRefAutoScan = True) Then
            unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status5", "Reference Required")
          Else
            If (unity_main.m_remoteRefScan = True) Then
              unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status9", "Reference Requested Remotely")
            Else
              unity_main.lbl_opStatus.Caption = MLSupport.GSS("OperStatus", "status10", "Ready for New Sample")
            End If
          End If
        End If
        
        ' Show repacks screen button if finished with repacks and predictions have been made for some properties
        If (unity_main.m_smplRepacks > 1) And (unity_main.m_makePred = "yes") And (frmedmod.numprops.Text <> 0) Then
          unity_main.cmd_repacks.Visible = True
        End If
  
        ' Show report screen image (if in supervisor mode)
        If (unity_main.run_min_gui = False) Then
          unity_main.img_report.Visible = True
        End If
      
        ' Check if to display write ticket image
        If (unity_main.m_writeTkt > 0) Then
          unity_main.img_ticket.Visible = True
        End If
      Else
        unity_main.m_smplPPTScan = False
      End If

      uniMsg = MLSupport.GSS("OperStatus", "status78", "Product sample completed")
      Call unity_main.log_scan_status(LOG_DBG_LEVEL2, "Product sample completed", uniMsg)

MODEL_FAIL:
      unity_main.m_scanTmrState = STS_COMPLETED
      unity_main.repcounter = 0
      tmr_sample.enabled = False
      cmd_sample.enabled = True
      If cmd_sample.Visible And cmd_sample.enabled Then
             'cmd_sample.SetFocus
            End If
      If (unity_main.m_batchRunFlg = False) Then
        ' Restart auto-loop
        tmr_all.enabled = True
      End If
  End Select
  
  Exit Sub
  
ReportScanError:
  ' Report error codes
  unity_main.lbl_opStatus.Caption = unity_main.m_uniErrMsg
  Call frm_status.report_error_codes(unity_main.m_ansiErrMsg, unity_main.m_uniErrMsg)

AbortScan:
  unity_main.m_scanTmrState = STS_ABORT
End Sub
#End If

Private Sub tmr_sec1_Timer()
  
  lblSysDateTime.Caption = Now
  
#If ABBFT Then
  If (Abs(unity_main.m_mb3000.m_lastRefTemp - unity_main.m_mb3000.m_currRefTemp) >= unity_main.m_mb3000.m_refTempDiff) Then
    unity_main.lbl_backdate.Caption = MLSupport.GSS("unity_main", "lbl_backdate_ext2", "External Reference has Expired")
    unity_main.m_extRefAutoScan = True
  Else
    unity_main.lbl_backdate.Caption = MLSupport.GGS_Params("unity_main.lbl_backdate_ext2", "External Reference Current Temp: %1; Last Temp: %2", CStr(unity_main.m_mb3000.m_currRefTemp), CStr(unity_main.m_mb3000.m_lastRefTemp))
  End If
#Else
  Select Case (unity_main.m_bType)
    Case "internal"
      ' Check if reference performed on demand
      If (unity_main.m_backFreq = REF_FREQ_ON_DEMAND) Then
        ' Check if reference timeout configured
        If (unity_main.m_intRefTimeout > 0) Then
          ' Display number of minutes left for reference
#If SSRCS Then
         Dim refTimeout As Long
         SSRCSClientError = SSRCSClient.GetRefTimeout(refTimeout)
          
          If (refTimeout > 0) Then
            unity_main.lbl_backdate.Caption = MLSupport.GGS_Params("unity_main.lbl_backdate_int", "Internal Reference Expires in %1 Minute(s)", CStr(refTimeout))
#Else
          If (MS11srv.refTimeout > 0) Then
            unity_main.lbl_backdate.Caption = MLSupport.GGS_Params("unity_main.lbl_backdate_int", "Internal Reference Expires in %1 Minute(s)", CStr(MS11srv.refTimeout))
#End If
          Else
            unity_main.lbl_backdate.Caption = MLSupport.GSS("unity_main", "lbl_backdate_int2", "Internal Reference has Expired")
          End If
        Else
          unity_main.lbl_backdate.Caption = MLSupport.GSS("unity_main", "lbl_backdate_int1", "Internal Reference Will Never Expire")
        End If
      Else
        unity_main.lbl_backdate.Caption = MLSupport.GSS("unity_main", "lbl_backdate_int3", "Internal Reference Performed Every Sample")
      End If
      
    Case "external"
      ' Check if reference timeout configured
      If (unity_main.m_extRefTimeoutSecs > 0) Then
        Dim extRefFileName As String
        Dim fDay As Long
        Dim fMonth As Long
        Dim fYear As Long
        Dim fHour As Long
        Dim fMinute As Long
        Dim fSecond As Long
        Dim fileDate As Date
        Dim refTimeDiff As Long
        
        extRefFileName = (REFERENCES_DIR & unity_main.m_extRefFileName & SPC_FILE_EXT)
        
        ' Check if have external reference spectrum file
        If (CFile.st_FileExist(extRefFileName) = True) Then
          ' Determine when last took reference in seconds
          Call CFile.st_GetFileDate(extRefFileName, esfid_lastwrite, fDay, fMonth, fYear, fHour, fMinute, fSecond)
          fileDate = (fYear & "-" & fMonth & "-" & fDay & " " & fHour & ":" & fMinute & ":" & fSecond)
          refTimeDiff = DateDiff("s", fileDate, lblSysDateTime.Caption)
         
          If (refTimeDiff < 0) Then
            refTimeDiff = 0
          End If
          
          ' Save time difference
          If (unity_main.m_extRefTimeoutSecs > refTimeDiff) Then
            unity_main.m_extRefTimer = unity_main.m_extRefTimeoutSecs - refTimeDiff
          Else
            unity_main.m_extRefTimer = -1
          End If
           
          ' Check if external reference has expired
          If (unity_main.m_extRefTimer < 0) Then
            unity_main.lbl_backdate.Caption = MLSupport.GSS("unity_main", "lbl_backdate_ext2", "External Reference has Expired")
             
            If (unity_main.m_extRefTimerIgnore = False) Then
              unity_main.m_extRefAutoScan = True
            End If
          Else
            unity_main.lbl_backdate.Caption = MLSupport.GGS_Params("unity_main.lbl_backdate_ext", "External Reference Expires in %1 Minute(s)", CStr(CInt(unity_main.m_extRefTimer / 60# + 0.5)))
          End If
        Else
          unity_main.lbl_backdate.Caption = MLSupport.GSS("unity_main", "lbl_backdate_ext2", "External Reference has Expired")
          
          ' Check if no request for external reference qualification scan
          If (unity_main.m_extRefPPTScan = False) Then
            unity_main.m_extRefAutoScan = True
          End If
        End If
      Else
        unity_main.lbl_backdate.Caption = MLSupport.GSS("unity_main", "lbl_backdate_ext1", "External Reference Will Never Expire")
      End If
      
    Case "file"    ' from offline file
      unity_main.lbl_backdate.Caption = MLSupport.GSS("unity_main", "lbl_backdate_ol1", "Offline Reference Will Never Expire")
  End Select
#End If
End Sub

Private Sub tmr_sec30_Timer()

#If SSTAR Then
  If ((unity_main.m_allowIntRefCalAccess = True) And (unity_main.m_intRefVerifyTimer <> 0) And (unity_main.m_intRefVerifyTimeout > 0)) Then
  '  If (unity_main.m_intRefVerifyAccumTime <= (CSng(unity_main.m_intRefVerifyTimeout * SECS_HR))) Then
   '  unity_main.m_intRefVerifyAccumTime = unity_main.m_intRefVerifyAccumTime + 30
   ' End If
 '  If chk_timeout(unity_main.m_intRefVerifyAccumTime, unity_main.m_intRefVerifyTimer) = True Then
    ' Check if time to display reminder to user
    
    If (unity_main.m_dblAccumTime < 0#) Then
      unity_main.m_intRefVerReminderFlg = True
      unity_main.m_dblAccumTime = unity_main.m_dblAccumTime + 1
      Exit Sub
    End If
      
    unity_main.m_dblAccumTime = unity_main.m_dblAccumTime + 1
            
    'If (chk_timeoutMine(unity_main.m_intRefVerifyAccumTime, CSng(unity_main.m_intRefVerifyTimeout) * secs) = True) Then
    If ((unity_main.m_dblAccumTime) > CDbl(unity_main.m_intRefVerifyTimeout * 60#)) Then
      If (m_intRefVerReminderCtr <= 0) Then
        unity_main.m_intRefVerReminderCtr = unity_main.m_intRefVerifyRemindTime      ' 2 counts/minute
        unity_main.m_intRefVerReminderFlg = True
      Else
        unity_main.m_intRefVerReminderCtr = unity_main.m_intRefVerReminderCtr - 1
      End If
    End If
  End If
#End If
End Sub

Private Sub txtsampcomment_DblClick(Button As Integer)
  
  unity_main.formfrom = 1
  unity_main.varfrom = 1
  frm_kybd.lbl_kybd.Caption = unity_main.lbl_sampleComment.Caption
  frm_kybd.txt_kybd.Text = unity_main.txtsampcomment.Text
  frm_kybd.Show 1
End Sub

Private Sub txtsamplename_DblClick(Button As Integer)
  
  unity_main.formfrom = 1
  unity_main.varfrom = 2
  frm_kybd.lbl_kybd.Caption = unity_main.lbl_sampleName.Caption
  frm_kybd.txt_kybd.Text = unity_main.txtsamplename.Text
  frm_kybd.Show 1
End Sub

Private Sub XYPlot1_CursorDisplayed(ByVal state As Boolean)

  If (state = False) Then
    lbl_cursorValue.Caption = ""
  End If
End Sub

Private Sub XYPlot1_CursorMoved(ByVal cursorXVal As Double, cursorYVals As Variant)
  Dim formatStr As String
  Dim cnt As Integer
  
  formatStr = FormatNumber(cursorYVals(0), 5, vbTrue, vbFalse, vbFalse)
 
  If (cursorYVals(0) < 0) Then
    cnt = 8
  Else
    cnt = 7
  End If
 
  lbl_cursorValue.Caption = Left(formatStr, cnt) & " @ " & FormatNumber(cursorXVal, 1, vbTrue, vbFalse, vbFalse)
End Sub
Private Function loadAccumSettings() As Boolean
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim rc As Boolean
  Dim lineCnt As Integer
  Dim inString As String
  Dim xx As String
  Dim strlen As Integer
  Dim tmpStrg As String
  Dim cfgVar As String
  Dim varVal As String
  If (uniFile.OpenFileRead(CFG_DIR & ACCUM_TIME_FILE) = True) Then
   On Error GoTo FILE_ERROR
   fEncoding = uniFile.ReadBOM
  ' Process each remaining line in .ini file
   While Not (uniFile.EOF())
    On Error GoTo FILE_ERROR
    ' Read line from file
    lineCnt = lineCnt + 1
    
    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(inString)
    Else
      rc = uniFile.ReadUnicodeLine(inString)
    End If
      
    If (rc = False) Then GoTo FILE_ERROR
          
    ' Get variable name and its value
    xx = InStr(1, inString, "=")
    
    ' If no "=" char, then older file format used "," filed delimiter
    If ((xx) = 0) Then
      xx = InStr(1, inString, ",")
    End If
    
    strlen = Len(inString)
    tmpStrg = Trim(Mid(inString, 1, xx - 1))
    cfgVar = LCase(tmpStrg)
    varVal = Trim(Mid(inString, xx + 1))
    Dim t As Single
    t = Timer
    
    ' Process value by variable name
    On Error GoTo FILE_ERROR
    Select Case (cfgVar)
     Case "refverifyaccumtime"
      '  m_intRefVerifyAccumTime = CDbl(varVal)
      m_dblAccumTime = CDbl(varVal)
      '  m_intRefVerifyAccumTime = m_intRefVerifyAccumTime * SECS_HR
      End Select
  Wend
  End If
  
FILE_ERROR:
loadAccumSettings = True
End Function
Private Function saveAccumSettings() As Boolean
  Dim uniFile As New clsUniFile
   
  Dim currTime As Single
  Dim accumTime As Single
  Dim deltaTime As Single
  
  'currTime = Timer
  
  'deltaTime = currTime - m_intRefStartTime
  
  'accumTime = unity_main.m_intRefVerifyAccumTime + deltaTime
  accumTime = unity_main.m_dblAccumTime
  If (uniFile.OpenFileWrite(CFG_DIR & ACCUM_TIME_FILE) = True) Then
   
      On Error GoTo FILE_ERROR
      uniFile.WriteBOM fe_UTF16LE
      uniFile.WriteUnicodeLine ("RefVerifyAccumTime=" & (accumTime))
      uniFile.CloseFile
      uniFile.Flush
  End If
FILE_ERROR:
End Function

Private Sub cmd_exit_Click()
  Dim uniMsg As String
  Dim optVal As Integer
 
  Call unity_main.kill_loop(LOG_DBG_LEVEL3, True, "Main screen Exit button selected")
  
  ' Check if run mode disabled
  If (unity_main.m_enableRunMode = False) Then
    uniMsg = MLSupport.GSS("unity_main", "statMsg1", "Are you sure you want to exit the software? Press OK to exit, Cancel to return")
    optVal = CWrap.ShowMessageBoxW(uniMsg, vbOKCancel)
    If (optVal = vbOK) Then
      Call frm_collect.savescansettings(False)
      Call frm_Inst.savemyinsts(False, False)
      Call frm_dynRptCfg.save_cfg(False, False)
      saveAccumSettings
      unity_main.unloadallforms "unity_main"
      Unload Me
      End
    End If
    Exit Sub
  Else
    saveAccumSettings
    ' Check if in run mode
    If (unity_main.run_min_gui = True) Then
      frm_exit.cmd_exit.Visible = False
      frm_exit.cmd_reboot.Visible = True
      frm_exit.cmd_shutdown.Visible = True
    Else
      frm_exit.cmd_exit.Visible = True
      frm_exit.cmd_reboot.Visible = False
      frm_exit.cmd_shutdown.Visible = False
    End If
    
    frm_exit.Show 1
  End If
End Sub



