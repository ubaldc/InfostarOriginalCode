VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_extRef 
   Caption         =   "External Reference Configuration"
   ClientHeight    =   9000
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   11880
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
   Icon            =   "frm_extRef.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin HexUniControls.ctlUniFrameXP frame_main 
      Height          =   7005
      Index           =   1
      Left            =   480
      Top             =   900
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   12356
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_extRef.frx":058A
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483642
      Tip             =   "frm_extRef.frx":05AA
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_extRef.frx":05CA
      Begin HexUniControls.ctlUniFrameXP frame_trayCfg 
         Height          =   6540
         Left            =   360
         Top             =   210
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   11536
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_extRef.frx":05E6
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_extRef.frx":062A
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_extRef.frx":064A
         Begin HexUniControls.ctlUniFrameXP frame_multiCupType 
            Height          =   1935
            Left            =   120
            Top             =   4200
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   3413
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frm_extRef.frx":0666
            Enabled         =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   16711680
            Tip             =   "frm_extRef.frx":06A2
            VistaStyle      =   -1  'True
            UseShadow       =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_extRef.frx":06C2
            Begin HexUniControls.ctlUniRadioXP opt_multiCupType 
               Height          =   450
               Index           =   6
               Left            =   5040
               TabIndex        =   18
               Top             =   1360
               Width           =   4600
               _ExtentX        =   8123
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
               Caption         =   "frm_extRef.frx":06DE
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_extRef.frx":073E
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":075E
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_multiCupType 
               Height          =   450
               Index           =   5
               Left            =   5040
               TabIndex        =   17
               Top             =   860
               Width           =   4600
               _ExtentX        =   8123
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
               Caption         =   "frm_extRef.frx":077A
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_extRef.frx":07E2
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":0802
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_multiCupType 
               Height          =   450
               Index           =   4
               Left            =   5040
               TabIndex        =   16
               Top             =   360
               Width           =   4600
               _ExtentX        =   8123
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
               Caption         =   "frm_extRef.frx":081E
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_extRef.frx":0874
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":0894
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_multiCupType 
               Height          =   450
               Index           =   3
               Left            =   120
               TabIndex        =   15
               Top             =   1360
               Width           =   4600
               _ExtentX        =   8123
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
               Caption         =   "frm_extRef.frx":08B0
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_extRef.frx":0916
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":0936
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_multiCupType 
               Height          =   450
               Index           =   2
               Left            =   120
               TabIndex        =   14
               Top             =   860
               Width           =   4600
               _ExtentX        =   8123
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
               Caption         =   "frm_extRef.frx":0952
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_extRef.frx":09BE
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":09DE
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_multiCupType 
               Height          =   450
               Index           =   1
               Left            =   120
               TabIndex        =   13
               Top             =   360
               Width           =   4600
               _ExtentX        =   8123
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
               Caption         =   "frm_extRef.frx":09FA
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_extRef.frx":0A68
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":0A88
               ShowFocus       =   -1  'True
            End
         End
         Begin HexUniControls.ctlUniFrameXP frame_rotatePlatter 
            Height          =   3840
            Left            =   4800
            Top             =   360
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   6773
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frm_extRef.frx":0AA4
            Enabled         =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   16711680
            Tip             =   "frm_extRef.frx":0AE4
            VistaStyle      =   -1  'True
            UseShadow       =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_extRef.frx":0B04
            Begin HexUniControls.ctlUniFrameXP frame_rotateMode 
               Height          =   2355
               Left            =   60
               Top             =   300
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   4154
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Unicode MS"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frm_extRef.frx":0B20
               Enabled         =   -1  'True
               BackColor       =   -2147483633
               ForeColor       =   16711680
               Tip             =   "frm_extRef.frx":0B5A
               VistaStyle      =   -1  'True
               UseShadow       =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":0B7A
               Begin HexUniControls.ctlUniRadioXP opt_rotateModeIndex 
                  Height          =   450
                  Left            =   120
                  TabIndex        =   22
                  Top             =   1860
                  Width           =   2055
                  _ExtentX        =   3625
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
                  Caption         =   "frm_extRef.frx":0B96
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frm_extRef.frx":0BC4
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frm_extRef.frx":0BE4
                  ShowFocus       =   -1  'True
               End
               Begin HexUniControls.ctlUniRadioXP opt_rotateModeStep 
                  Height          =   450
                  Left            =   120
                  TabIndex        =   21
                  Top             =   1360
                  Width           =   2055
                  _ExtentX        =   3625
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
                  Caption         =   "frm_extRef.frx":0C00
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frm_extRef.frx":0C2E
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frm_extRef.frx":0C4E
                  ShowFocus       =   -1  'True
               End
               Begin HexUniControls.ctlUniRadioXP opt_rotateModeCont 
                  Height          =   450
                  Left            =   120
                  TabIndex        =   20
                  Top             =   860
                  Width           =   2055
                  _ExtentX        =   3625
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
                  Caption         =   "frm_extRef.frx":0C6A
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frm_extRef.frx":0C9E
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frm_extRef.frx":0CBE
                  ShowFocus       =   -1  'True
               End
               Begin HexUniControls.ctlUniRadioXP opt_rotateModeNone 
                  Height          =   450
                  Left            =   120
                  TabIndex        =   19
                  Top             =   360
                  Width           =   2055
                  _ExtentX        =   3625
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
                  Caption         =   "frm_extRef.frx":0CDA
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frm_extRef.frx":0D02
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frm_extRef.frx":0D22
                  ShowFocus       =   -1  'True
               End
            End
            Begin HexUniControls.ctlUniTextBoxXP txt_cmplRevSteps 
               Height          =   375
               Left            =   3840
               TabIndex        =   34
               Top             =   3360
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   661
               BackColor       =   -2147483633
               ForeColor       =   -2147483640
               Enabled         =   0   'False
               Locked          =   0   'False
               Text            =   "frm_extRef.frx":0D3E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxLength       =   0
               MultiLine       =   0   'False
               Alignment       =   1
               ScrollBars      =   0
               PasswordChar    =   ""
               TrapTab         =   0   'False
               EnableContextMenu=   -1  'True
               RaiseChangeEvent=   -1  'True
               Tip             =   "frm_extRef.frx":0D62
               NoHideSel       =   0   'False
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":0D82
            End
            Begin HexUniControls.ctlNumIncXP numInc_rotateIndexSteps 
               Height          =   600
               Left            =   3840
               TabIndex        =   27
               Top             =   2640
               Width           =   1110
               _ExtentX        =   1958
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
               Max             =   1500
               AllowSpace      =   0   'False
               BorderColor     =   -1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ButtonBackColor =   -2147483633
               ButtonForeColor =   0
               ButtonStyle     =   4
               ButtonWidth     =   20
               Tip             =   ""
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":0D9E
               TrapTabKey      =   0   'False
            End
            Begin HexUniControls.ctlUniLabel lbl_cmplRevSteps 
               Height          =   450
               Left            =   120
               Top             =   3360
               Width           =   3600
               _ExtentX        =   6350
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
               Caption         =   "frm_extRef.frx":0DBA
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   1
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frm_extRef.frx":0E22
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":0E42
            End
            Begin HexUniControls.ctlNumIncXP numInc_rotateStepSteps 
               Height          =   600
               Left            =   3840
               TabIndex        =   26
               Top             =   2640
               Width           =   1110
               _ExtentX        =   1958
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
               Text            =   "0"
               Min             =   1
               Max             =   100
               AllowSpace      =   0   'False
               BorderColor     =   -1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ButtonBackColor =   -2147483633
               ButtonForeColor =   0
               ButtonStyle     =   4
               ButtonWidth     =   20
               Tip             =   ""
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":0E5E
               TrapTabKey      =   0   'False
            End
            Begin HexUniControls.ctlNumIncXP numInc_rotateSpeed 
               Height          =   600
               Left            =   3840
               TabIndex        =   25
               Top             =   2640
               Width           =   1110
               _ExtentX        =   1958
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
               Max             =   100
               AllowSpace      =   0   'False
               BorderColor     =   -1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ButtonBackColor =   -2147483633
               ButtonForeColor =   0
               ButtonStyle     =   4
               ButtonWidth     =   20
               Tip             =   ""
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":0E7A
               TrapTabKey      =   0   'False
            End
            Begin HexUniControls.ctlUniFrameXP frame_rotateDir 
               Height          =   1395
               Left            =   2400
               Top             =   300
               Width           =   2710
               _ExtentX        =   4789
               _ExtentY        =   2461
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Unicode MS"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frm_extRef.frx":0E96
               Enabled         =   -1  'True
               BackColor       =   -2147483633
               ForeColor       =   16711680
               Tip             =   "frm_extRef.frx":0EC8
               VistaStyle      =   -1  'True
               UseShadow       =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":0EE8
               Begin HexUniControls.ctlUniRadioXP opt_rotateDirCCW 
                  Height          =   450
                  Left            =   120
                  TabIndex        =   24
                  Top             =   860
                  Width           =   2530
                  _ExtentX        =   4471
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
                  Caption         =   "frm_extRef.frx":0F04
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frm_extRef.frx":0F46
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frm_extRef.frx":0F66
                  ShowFocus       =   -1  'True
               End
               Begin HexUniControls.ctlUniRadioXP opt_rotateDirCW 
                  Height          =   450
                  Left            =   120
                  TabIndex        =   23
                  Top             =   360
                  Width           =   2530
                  _ExtentX        =   4471
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
                  Caption         =   "frm_extRef.frx":0F82
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frm_extRef.frx":0FB4
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frm_extRef.frx":0FD4
                  ShowFocus       =   -1  'True
               End
            End
            Begin HexUniControls.ctlUniLabel lbl_rotateIndexSteps 
               Height          =   600
               Left            =   120
               Top             =   2640
               Width           =   3600
               _ExtentX        =   6350
               _ExtentY        =   1058
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Unicode MS"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frm_extRef.frx":0FF0
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   1
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frm_extRef.frx":1058
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":1078
            End
            Begin HexUniControls.ctlUniLabel lbl_rotateStepSteps 
               Height          =   600
               Left            =   120
               Top             =   2640
               Width           =   3600
               _ExtentX        =   6350
               _ExtentY        =   1058
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Unicode MS"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frm_extRef.frx":1094
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   1
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frm_extRef.frx":10FC
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":111C
            End
            Begin HexUniControls.ctlUniLabel lbl_rotateSpeed 
               Height          =   600
               Left            =   120
               Top             =   2640
               Width           =   3600
               _ExtentX        =   6350
               _ExtentY        =   1058
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Unicode MS"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "frm_extRef.frx":1138
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   1
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frm_extRef.frx":1178
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":1198
            End
         End
         Begin HexUniControls.ctlUniFrameXP frame_adpaterType 
            Height          =   3840
            Left            =   120
            Top             =   360
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   6773
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frm_extRef.frx":11B4
            Enabled         =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   16711680
            Tip             =   "frm_extRef.frx":11EC
            VistaStyle      =   -1  'True
            UseShadow       =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_extRef.frx":120C
            Begin HexUniControls.ctlUniRadioXP opt_adapterType 
               Height          =   450
               Index           =   5
               Left            =   120
               TabIndex        =   12
               Top             =   2360
               Width           =   4335
               _ExtentX        =   7646
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
               Caption         =   "frm_extRef.frx":1228
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_extRef.frx":126A
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":128A
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_adapterType 
               Height          =   450
               Index           =   4
               Left            =   120
               TabIndex        =   11
               Top             =   1860
               Width           =   4335
               _ExtentX        =   7646
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
               Caption         =   "frm_extRef.frx":12A6
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_extRef.frx":1306
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":1326
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_adapterType 
               Height          =   450
               Index           =   3
               Left            =   120
               TabIndex        =   10
               Top             =   1360
               Width           =   4335
               _ExtentX        =   7646
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
               Caption         =   "frm_extRef.frx":1342
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_extRef.frx":1382
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":13A2
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_adapterType 
               Height          =   450
               Index           =   2
               Left            =   120
               TabIndex        =   9
               Top             =   860
               Width           =   4335
               _ExtentX        =   7646
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
               Caption         =   "frm_extRef.frx":13BE
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_extRef.frx":13F6
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":1416
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_adapterType 
               Height          =   450
               Index           =   1
               Left            =   120
               TabIndex        =   8
               Top             =   360
               Width           =   4335
               _ExtentX        =   7646
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
               Caption         =   "frm_extRef.frx":1432
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_extRef.frx":146C
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_extRef.frx":148C
               ShowFocus       =   -1  'True
            End
         End
      End
   End
   Begin HexUniControls.ctlUniFrameXP frame_main 
      Height          =   7005
      Index           =   0
      Left            =   480
      Top             =   900
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   12356
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_extRef.frx":14A8
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483642
      Tip             =   "frm_extRef.frx":14C8
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_extRef.frx":14E8
      Begin HexUniControls.ctlUniFrameXP Frame5 
         Height          =   2730
         Left            =   480
         Top             =   3660
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   4815
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_extRef.frx":1504
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_extRef.frx":1544
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_extRef.frx":1564
         Begin HexUniControls.ctlUniLabel lbl_maxWvln 
            Height          =   285
            Left            =   1200
            Top             =   2160
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frm_extRef.frx":1580
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_extRef.frx":15AE
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_extRef.frx":15CE
         End
         Begin HexUniControls.ctlUniLabel lbl_minWvln 
            Height          =   285
            Left            =   1200
            Top             =   1755
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frm_extRef.frx":15EA
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_extRef.frx":1618
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_extRef.frx":1638
         End
         Begin HexUniControls.ctlUniLabel lbl_endWvln 
            Height          =   285
            Left            =   1200
            Top             =   1365
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frm_extRef.frx":1654
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_extRef.frx":1680
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_extRef.frx":16A0
         End
         Begin HexUniControls.ctlUniLabel lbl_startWvln 
            Height          =   285
            Left            =   1200
            Top             =   960
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frm_extRef.frx":16BC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_extRef.frx":16EC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_extRef.frx":170C
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_maxWvln 
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   2160
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   661
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "frm_extRef.frx":1728
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   0
            MultiLine       =   0   'False
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frm_extRef.frx":1748
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_extRef.frx":1768
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_minWvln 
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   1760
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   661
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "frm_extRef.frx":1784
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   0
            MultiLine       =   0   'False
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frm_extRef.frx":17A4
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_extRef.frx":17C4
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_dfltEndWvln 
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   1360
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   661
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "frm_extRef.frx":17E0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   0
            MultiLine       =   0   'False
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frm_extRef.frx":1800
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_extRef.frx":1820
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_endWvln 
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   1360
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   661
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_extRef.frx":183C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   0
            MultiLine       =   0   'False
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frm_extRef.frx":185C
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_extRef.frx":187C
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_dfltStartWvln 
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   960
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   661
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "frm_extRef.frx":1898
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   0
            MultiLine       =   0   'False
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frm_extRef.frx":18B8
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_extRef.frx":18D8
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_startWvln 
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   661
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_extRef.frx":18F4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   0
            MultiLine       =   0   'False
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frm_extRef.frx":1914
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_extRef.frx":1934
         End
         Begin HexUniControls.ctlUniCheckXP chk_dfltWavelens 
            Height          =   525
            Left            =   120
            TabIndex        =   5
            Top             =   350
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   926
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frm_extRef.frx":1950
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_extRef.frx":1998
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_extRef.frx":19B8
            ShowFocus       =   -1  'True
         End
      End
      Begin HexUniControls.ctlUniFrameXP Frame1 
         Height          =   3015
         Left            =   480
         Top             =   420
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   5318
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_extRef.frx":19D4
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_extRef.frx":1A18
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_extRef.frx":1A38
         Begin HexUniControls.ctlUniLabel lbl_refNScans 
            Height          =   615
            Left            =   1440
            Top             =   1920
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   1085
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frm_extRef.frx":1A54
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_extRef.frx":1AAE
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_extRef.frx":1ACE
         End
         Begin HexUniControls.ctlNumIncXP numInc_refNScans 
            Height          =   615
            Left            =   120
            TabIndex        =   4
            Top             =   1920
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1085
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
            Text            =   "24"
            Min             =   1
            Max             =   512
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
            MouseIcon       =   "frm_extRef.frx":1AEA
            TrapTabKey      =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_refPPT 
            Height          =   615
            Left            =   1440
            Top             =   1200
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   1085
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frm_extRef.frx":1B06
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_extRef.frx":1B84
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_extRef.frx":1BA4
         End
         Begin HexUniControls.ctlNumIncXP numInc_refPPT 
            Height          =   615
            Left            =   120
            TabIndex        =   3
            Top             =   1200
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1085
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
            Text            =   "300"
            Min             =   0
            Max             =   300
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
            MouseIcon       =   "frm_extRef.frx":1BC0
            TrapTabKey      =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_refTimeout 
            Height          =   615
            Left            =   1440
            Top             =   480
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   1085
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frm_extRef.frx":1BDC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_extRef.frx":1C42
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_extRef.frx":1C62
         End
         Begin HexUniControls.ctlNumIncXP numInc_refTimeout 
            Height          =   615
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1085
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
            Text            =   "2880"
            Min             =   0
            Max             =   2880
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
            MouseIcon       =   "frm_extRef.frx":1C7E
            TrapTabKey      =   0   'False
         End
      End
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_extRefName 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   50
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_extRef.frx":1C9A
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
      Tip             =   "frm_extRef.frx":1CBA
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_extRef.frx":1CDA
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   11400
      Top             =   1080
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9000
      FormDesignWidth =   11880
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   8520
      TabIndex        =   28
      Top             =   8040
      Width           =   2000
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_extRef.frx":1CF6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_extRef.frx":1D22
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_extRef.frx":1D42
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   650
      Left            =   5880
      TabIndex        =   29
      Top             =   8040
      Width           =   2000
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_extRef.frx":1D5E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_extRef.frx":1D96
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_extRef.frx":1E14
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   375
      Left            =   120
      Top             =   50
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_extRef.frx":1E30
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_extRef.frx":1E58
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_extRef.frx":1E78
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_extRefDesc 
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   50
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_extRef.frx":1E94
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
      Tip             =   "frm_extRef.frx":1EB4
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_extRef.frx":1ED4
   End
   Begin HexUniControls.ctlUniLabel Label2 
      Height          =   375
      Left            =   4440
      Top             =   50
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_extRef.frx":1EF0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_extRef.frx":1F1E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_extRef.frx":1F3E
   End
   Begin HexUniControls.ctlUniTabbedXP tab_frame 
      Height          =   315
      Left            =   600
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   600
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   16711680
      BackColorTab    =   -2147483633
      ForeColorTab    =   8947848
      ShowFocus       =   0   'False
      Tip             =   "frm_extRef.frx":1F5A
      ButtonStyle     =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_extRef.frx":1F7A
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   11400
      Top             =   2280
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
      Left            =   11400
      Top             =   1800
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_extRef.frx":1F96
   End
End
Attribute VB_Name = "frm_extRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_extRefAdapterType As String
Public m_extRefAdaptIndx As Integer
Public m_extRefDesc As String
Public m_extRefEndWvln As Double
Public m_extRefFileName As String
Public m_extRefMultiCupIndx As Integer
Public m_extRefMultiCupType As String
Public m_extRefName As String
Public m_extRefNScans As Integer
Public m_extRefPPT As Integer
Public m_extRefRotateDir As TRAY_ROTATE_DIRS
Public m_extRefRotateIndexSteps As Integer
Public m_extRefRotateMoveMode As TRAY_ROTATE_MOVEMENTS
Public m_extRefRotateSpeed As Integer
Public m_extRefRotateStepSteps As Integer
Public m_extRefStartWvln As Double
Public m_extRefTimeout As Integer
Public m_extRefTrayNum As Integer

Private m_dfltSettings As Boolean
Private m_ignoreEvent As Boolean
Private m_badIniVal As Boolean

Public Sub init_ext_ref_settings()
  Dim ii As Integer

  m_dfltSettings = True
  
  ' Setup to display first tab frame
  show_frame 0
  
  txt_extRefName.Text = ""
  txt_extRefDesc.Text = ""
  
  chk_dfltWavelens.Value = 1
  txt_dfltStartWvln.Text = ProdDfltData.startWvln
  txt_dfltStartWvln.Visible = True
  txt_dfltEndWvln.Text = ProdDfltData.endWvln
  txt_dfltEndWvln.Visible = True
    
  txt_startWvln.Text = ProdDfltData.startWvln
  txt_startWvln.Visible = False
  txt_endWvln.Text = ProdDfltData.endWvln
  txt_endWvln.Visible = False
  txt_minWvln.Text = MS11CfgData.minWvln
  txt_minWvln.Visible = False
  lbl_minWvln.Visible = False
  txt_maxWvln.Text = MS11CfgData.maxWvln
  txt_maxWvln.Visible = False
  lbl_maxWvln.Visible = False

  numInc_refNScans.Text = unity_main.m_intRefNScans
  numInc_refPPT.Text = unity_main.m_intRefPPT
  numInc_refTimeout.Text = unity_main.m_intRefTimeout

  m_extRefAdapterType = ProdDfltData.adapterType
  m_extRefMultiCupType = ProdDfltData.multiCupType
  m_extRefRotateDir = ProdDfltData.rotateDir
  m_extRefRotateIndexSteps = ProdDfltData.rotateIndexSteps
  m_extRefRotateMoveMode = ProdDfltData.rotateMoveMode
  m_extRefRotateSpeed = ProdDfltData.rotateSpeed
  m_extRefRotateStepSteps = ProdDfltData.rotateStepSteps
  
  setup_ext_ref_tray_cfg
  m_dfltSettings = False
End Sub

Function load_ext_ref_cfg_file(extRefFile As String, prodSelect As Boolean) As Boolean
  Dim tmpFile As String
  Dim lineCnt As Integer
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim inString As String
  Dim varStr As Variant
  Dim rc As Boolean
  Dim errMsg As String
  Dim uniMsg As String

  m_badIniVal = False
  load_ext_ref_cfg_file = False
  m_extRefFileName = EXT_REFS_CFG_DIR & extRefFile & CFG_FILE_EXT
  
  If (uniFile.st_FileExist(m_extRefFileName) = False) Then
    unity_main.errorstring = (m_extRefFileName & " file not found")
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg6", "%1 file not found", m_extRefFileName)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    GoTo LOAD_ERROR
  End If
  
  'If it made it here, the config file exists, now copy to temp file and load it
  On Error GoTo BAD_FILE
  tmpFile = (EXT_REFS_CFG_DIR & TMP_LOAD_EXT_REF_CFG_FILE)
  uniFile.st_CopyFile m_extRefFileName, tmpFile
  
  If (uniFile.OpenFileRead(tmpFile) = False) Then GoTo BAD_FILE
  
  fEncoding = uniFile.ReadBOM
  
  lineCnt = 0
  lineCnt = lineCnt + 1
    
  If (fEncoding = fe_ANSI) Then
    rc = uniFile.ReadAnsiLine(inString)
  Else
    rc = uniFile.ReadUnicodeLine(inString)
  End If
      
  If (rc = False) Then GoTo BAD_FILE

  While Not (uniFile.EOF())
    Select Case (inString)
      Case "[signature settings]"
        Call unity_main.load_file_signature_vals(m_extRefFileName, uniFile, fEncoding, lineCnt, m_badIniVal)
        inString = unity_main.m_iniString
       
      Case "[external ref settings]"
        Call load_ext_ref_file_vals(m_extRefFileName, uniFile, fEncoding, lineCnt)
        inString = unity_main.m_iniString
        
        If (unity_main.process_file_signature_vals(m_extRefFileName, False, m_badIniVal) = False) Then
          GoTo LOAD_ERROR
        End If
        
        Call process_ext_ref_file_vars(m_extRefFileName, uniFile, fEncoding, lineCnt, prodSelect)
        GoTo FILE_PROCESSED
        
      Case Else
        GoTo BAD_FILE
    End Select
  Wend
  
FILE_PROCESSED:
  ' Update operational parameters if loading selected product
  If (prodSelect = True) Then
    unity_main.m_extRefAdapterType = m_extRefAdapterType
    unity_main.m_extRefAdaptIndx = m_extRefAdaptIndx
    unity_main.m_extRefEndWvln = m_extRefEndWvln
    unity_main.m_extRefMultiCupIndx = m_extRefMultiCupIndx
    unity_main.m_extRefMultiCupType = m_extRefMultiCupType
    unity_main.m_extRefNScans = m_extRefNScans
    unity_main.m_extRefPPT = m_extRefPPT
    unity_main.m_extRefRotateDir = m_extRefRotateDir
    unity_main.m_extRefRotateIndexSteps = m_extRefRotateIndexSteps
    unity_main.m_extRefRotateMoveMode = m_extRefRotateMoveMode
    unity_main.m_extRefRotateSpeed = m_extRefRotateSpeed
    unity_main.m_extRefRotateStepSteps = m_extRefRotateStepSteps
    unity_main.m_extRefStartWvln = m_extRefStartWvln
    unity_main.m_extRefTimeout = m_extRefTimeout
    unity_main.m_extRefTrayNum = m_extRefTrayNum
    
    ' Update working external refernce timeout if less than current timer
    ' and greater than new timeout value
    If ((unity_main.m_extRefTimer = 0) Or ((m_extRefTimeout * 60) < unity_main.m_extRefTimer)) And _
       ((unity_main.m_extRefTimeoutSecs = 0) Or ((unity_main.m_extRefTimeoutSecs / 60) > m_extRefTimeout)) Then
      unity_main.m_extRefTimeoutSecs = m_extRefTimeout * 60
    End If
  End If
        
  ' Check if ini file had bad value
  If (m_badIniVal = True) Then
    unity_main.errorstring = (m_extRefFileName & " had incompatible value(s). Updated file with proper/default value(s).")
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg7", "%1 had incompatible value(s). Updated file with proper/default value(s)", m_extRefFileName)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg3", "%1. Refer to system log for more details", uniMsg), vbOKOnly
    frm_extRef.save_ext_ref_settings
  End If
  
  uniFile.CloseFile
  uniFile.st_RmFile tmpFile
  load_ext_ref_cfg_file = True
  Exit Function
  
BAD_FILE:
  If (lineCnt = 0) Then
    errMsg = (m_extRefFileName & " file open error." & Error$)
    uniMsg = MLSupport.GGS_Params("fileErrMsg2", "%1 file open error. %2", m_extRefFileName, Error$)
  Else
    errMsg = (m_extRefFileName & " file has error on line " & CStr(lineCnt) & ". " & Error$)
    uniMsg = MLSupport.GGS_Params("fileErrMsg3", "%1 file has error on line %2. %3", m_extRefFileName, CStr(lineCnt), Error$)
  End If
  
  unity_main.errorstring = errMsg
  unity_main.write_error
  
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  
LOAD_ERROR:
  uniFile.CloseFile
  unity_main.passstring = ""
  
  If (tmpFile <> "") And (uniFile.st_FileExist(tmpFile) = True) Then
    uniFile.st_RmFile tmpFile
  End If
End Function

Sub load_ext_ref_file_vals(ByVal fileName As String, ByVal uniFile As clsUniFile, ByVal fEncoding As eFileEncoding, ByRef lineCnt As Integer)
  Dim xx As Variant
  Dim inString As String
  Dim tmpStrg As String
  Dim cfgVar As String
  Dim varVal As String
  Dim strlen As Integer
  Dim rc As Boolean

  unity_main.m_iniString = ""

  ' Process each line in external ref settings section
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
        m_extRefAdapterType = varVal
      Case "desc"
        m_extRefDesc = varVal
      Case "endwvln"
        m_extRefEndWvln = CDbl(varVal)
      Case "multicuptype"
        m_extRefMultiCupType = varVal
      Case "name"
        m_extRefName = varVal
      Case "refnscans"
        m_extRefNScans = CInt(varVal)
      Case "refppt"
        m_extRefPPT = CInt(varVal)
      Case "reftimeout"
        m_extRefTimeout = CInt(varVal)
      Case "rotatedir"
        m_extRefRotateDir = CInt(varVal)
      Case "rotateindexsteps"
        m_extRefRotateIndexSteps = CInt(varVal)
      Case "rotatemovemode"
        m_extRefRotateMoveMode = CInt(varVal)
      Case "rotatespeed"
        m_extRefRotateSpeed = CInt(varVal)
      Case "rotatestepsteps"
        m_extRefRotateStepSteps = CInt(varVal)
      Case "startwvln"
        m_extRefStartWvln = CDbl(varVal)
    End Select
  Wend
  
  Exit Sub
  
BAD_INI_VALUE:
    unity_main.errorstring = (fileName & " line " & CStr(lineCnt) & " had incompatible value. " & Trim(tmpStrg) & " - " & varVal & "; will use default value")
    unity_main.write_error
    m_badIniVal = True
    Resume Next
  
FILE_ERROR:
  unity_main.m_iniString = ""
End Sub

Public Sub process_ext_ref_file_vars(fileName As String, ByVal uniFile As clsUniFile, ByVal fEncoding As eFileEncoding, ByRef lineCnt As Integer, prodSelect As Boolean)

  ' Setup number of reference scans, PPT and timeout
  numInc_refNScans.Text = m_extRefNScans
  numInc_refPPT.Text = m_extRefPPT
  numInc_refTimeout.Text = m_extRefTimeout

  ' Setup product & default starting/ending wavelengths
  txt_dfltEndWvln.Text = ProdDfltData.endWvln
  txt_dfltStartWvln.Text = ProdDfltData.startWvln
  txt_endWvln.Text = m_extRefEndWvln
  txt_startWvln.Text = m_extRefStartWvln
  
  ' Setup instrument min/max wavelengths
  txt_minWvln.Text = unity_main.m_minWvln
  txt_maxWvln.Text = unity_main.m_maxWvln

  ' Check if wavelengths match product defaults
  If (ProdDfltData.endWvln = m_extRefEndWvln) And (ProdDfltData.startWvln = m_extRefStartWvln) Then
    chk_dfltWavelens.Value = 1
    txt_dfltStartWvln.Visible = True
    txt_dfltEndWvln.Visible = True
    txt_startWvln.Visible = False
    txt_endWvln.Visible = False
    txt_minWvln.Visible = False
    lbl_minWvln.Visible = False
    txt_maxWvln.Visible = False
    lbl_maxWvln.Visible = False
  Else
    chk_dfltWavelens.Value = 0
    txt_dfltStartWvln.Visible = False
    txt_dfltEndWvln.Visible = False
    txt_startWvln.Visible = True
    txt_endWvln.Visible = True
    txt_minWvln.Visible = True
    lbl_minWvln.Visible = True
    txt_maxWvln.Visible = True
    lbl_maxWvln.Visible = True
  End If

  txt_extRefName.Text = m_extRefName
  txt_extRefDesc.Text = m_extRefDesc

  ' Check and setup tray configuration parameters
  setup_ext_ref_tray_cfg
End Sub

Public Function check_ext_ref_settings() As Boolean
  Dim startWvln As Double
  Dim endWvln As Double
  Dim rc As Boolean
  Dim userReq As Integer

  check_ext_ref_settings = False

  If (txt_extRefName.Text = "") Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_extRef", "errMsg5", "Please enter a external reference name"), vbExclamation
    Exit Function
  End If

  ' Check values if using custom scan wavelength range
  If (chk_dfltWavelens.Value = 0) Then
    txt_startWvln.Text = Trim(txt_startWvln.Text)
    txt_endWvln.Text = Trim(txt_endWvln.Text)

    On Error GoTo BAD_VALUE
    startWvln = CDbl(txt_startWvln.Text)
    endWvln = CDbl(txt_endWvln.Text)
    
    If (startWvln < unity_main.m_minWvln) Or (startWvln > unity_main.m_maxWvln) Then
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_collect.errMsg1", "Please enter a Scan Starting Wavelength value in Scan Settings Tab between %1 and %2", CStr(unity_main.m_minWvln), CStr(unity_main.m_maxWvln)), vbExclamation
      Exit Function
    Else
      If (endWvln < unity_main.m_minWvln) Or (endWvln > unity_main.m_maxWvln) Then
        CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_collect.errMsg2", "Please enter a Scan Ending Wavelength value in Scan Settings Tab between %1 and %2", CStr(unity_main.m_minWvln), CStr(unity_main.m_maxWvln)), vbExclamation
        Exit Function
      Else
        If (endWvln <= startWvln) Then
          CWrap.ShowMessageBoxW MLSupport.GSS("frm_extRef", "errMsg1", "Please enter a Scan Ending Wavelength value in Scan Settings Tab greater than the Starting Wavelength value"), vbExclamation
          Exit Function
        Else
          ' Check wavelengths with instrument's table
#If SSRCS Then
          SSRCSClientError = unity_main.SSRCSClient.ChkWvlnRange(startWvln, endWvln, rc)
#Else
          rc = unity_main.MS11srv.ChkWvlnRange(startWvln, endWvln)
#End If

          ' Check if wavelengths changed by instrument, if so ask user if okay
          If (rc = False) Then
            userReq = CWrap.ShowMessageBoxW(MLSupport.GGS_Params("frm_collect.errMsg3", "Starting/Ending Wavelengths in Scan Settings were changed to %1 - %2. Do you want use these values?", CStr(startWvln), CStr(endWvln)), vbYesNo)
            
            If (userReq = vbNo) Then
              Exit Function
            Else
              txt_startWvln.Text = startWvln
              txt_endWvln.Text = endWvln
            End If
          End If
        End If
      End If
    End If
  End If

  ' Check if no multi-cup type selected for multicup adapter
  If (m_extRefAdapterType = CFG_MULTI_CUP_AT) And (m_extRefMultiCupType = CFG_NONE_MCT) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_extRef", "errMsg2", "Please select a Multi-Cup Type for Multi-Cup Adapter in Tray Settings Tab"), vbExclamation
    Exit Function
  End If
  
  If (m_extRefRotateMoveMode <> TRM_NONE) And (m_extRefRotateDir = TRD_NONE) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_extRef", "errMsg3", "Please select a Direction for Rotating Platter in Tray Settings Tab"), vbExclamation
    Exit Function
  End If
  
  check_ext_ref_settings = True
  Exit Function
  
BAD_VALUE:
  CWrap.ShowMessageBoxW MLSupport.GSS("frm_extRef", "errMsg4", "Please enter a valid number for Scan Starting/Ending Wavelength"), vbOKOnly
End Function

Public Sub save_ext_ref_settings()
  Dim nn As Integer
  Dim tmpFile As String
  Dim uniFile As New clsUniFile
  Dim errMsg As String
  Dim uniMsg As String

  ' Save number of reference scans, PPT and timeout
  m_extRefNScans = Trim(numInc_refNScans.Text)
  m_extRefPPT = Trim(numInc_refPPT.Text)
  m_extRefTimeout = Trim(numInc_refTimeout.Text)

  ' Determine if to use product default or custom wavelengths
  If (chk_dfltWavelens.Value = 0) Then
    m_extRefEndWvln = frm_extRef.txt_endWvln.Text
    m_extRefStartWvln = frm_extRef.txt_startWvln.Text
  Else
    m_extRefEndWvln = frm_extRef.txt_dfltEndWvln.Text
    m_extRefStartWvln = frm_extRef.txt_dfltStartWvln.Text
  End If
  
  ' Save multi-cup type selection
  m_extRefMultiCupType = CFG_NONE_MCT
  
  For nn = 1 To MAX_MULTI_CUP_TYPES
    If (opt_multiCupType(nn).Value = True) Then
      m_extRefMultiCupType = MS11MultiCupInfo(nn).cfgName
      GoTo LEAVE_LOOP
    End If
  Next nn
  
LEAVE_LOOP:
  ' Save reference name and description
  m_extRefName = txt_extRefName.Text
  m_extRefDesc = txt_extRefDesc.Text

  ' Save rotation direction selection
  If (opt_rotateDirCW.Value = True) Then
    m_extRefRotateDir = TRD_CW
  Else
    If (opt_rotateDirCCW.Value = True) Then
      m_extRefRotateDir = TRD_CCW
    Else
      m_extRefRotateDir = TRD_NONE
    End If
  End If
  
  ' Save rotation mode selection and values
  If (frm_extRef.opt_rotateModeNone.Value = True) Then
    ' Setup for no rotation
    m_extRefRotateMoveMode = TRM_NONE
    m_extRefRotateIndexSteps = frm_extRef.numInc_rotateIndexSteps.Text
    m_extRefRotateSpeed = frm_extRef.numInc_rotateSpeed.Text
    m_extRefRotateStepSteps = frm_extRef.numInc_rotateStepSteps.Text
  Else
    If (frm_extRef.opt_rotateModeCont.Value = True) Then
      ' Setup for continuous rotation
      m_extRefRotateMoveMode = TRM_CONT
      m_extRefRotateIndexSteps = 0
      m_extRefRotateSpeed = frm_extRef.numInc_rotateSpeed.Text
      m_extRefRotateStepSteps = 0
    Else
      If (frm_extRef.opt_rotateModeStep.Value = True) Then
        ' Setup for steps for stepped rotation
        m_extRefRotateMoveMode = TRM_STEP
        m_extRefRotateIndexSteps = 0
        m_extRefRotateSpeed = 0
        m_extRefRotateStepSteps = frm_extRef.numInc_rotateStepSteps.Text
      Else
        If (frm_extRef.opt_rotateModeIndex.Value = True) Then
          ' Setup for steps for indexed rotation
          m_extRefRotateMoveMode = TRM_INDEX
          m_extRefRotateIndexSteps = frm_extRef.numInc_rotateIndexSteps.Text
          m_extRefRotateSpeed = 0
          m_extRefRotateStepSteps = 0
        End If
      End If
    End If
  End If

  ' Build ini file name
  On Error GoTo FILE_ERROR
  m_extRefFileName = (EXT_REFS_CFG_DIR & m_extRefStartWvln & "-" & m_extRefEndWvln & EXT_REF_SCAN_FILE & m_extRefName & CFG_FILE_EXT)
  tmpFile = (EXT_REFS_CFG_DIR & TMP_SAVE_EXT_REF_CFG_FILE)
  
  ' Create and write out to temporary ini file
  If (uniFile.OpenFileWrite(tmpFile) = True) Then
    On Error GoTo FILE_ERROR
    uniFile.WriteBOM fe_UTF16LE
    uniFile.WriteUnicodeLine "[signature settings]"
    uniFile.WriteUnicodeLine ("DevID=" & MS11CfgData.devID)
    uniFile.WriteUnicodeLine ("SmplTable=" & unity_main.m_smplTable)
    uniFile.WriteUnicodeLine ("ScanMode=" & unity_main.m_sysScanMode)
    uniFile.WriteUnicodeLine ("Version=" & INFOSTAR_VER)
    uniFile.WriteUnicodeLine "[external ref settings]"
    uniFile.WriteUnicodeLine ("AdapterType=" & m_extRefAdapterType)
    uniFile.WriteUnicodeLine ("Desc=" & Trim(m_extRefDesc))
    uniFile.WriteUnicodeLine ("EndWvln=" & m_extRefEndWvln)
    uniFile.WriteUnicodeLine ("MultiCupType=" & m_extRefMultiCupType)
    uniFile.WriteUnicodeLine ("Name=" & m_extRefName)
    uniFile.WriteUnicodeLine ("RefNScans=" & m_extRefNScans)
    uniFile.WriteUnicodeLine ("RefPPT=" & m_extRefPPT)
    uniFile.WriteUnicodeLine ("RefTimeout=" & m_extRefTimeout)
    uniFile.WriteUnicodeLine ("RotateDir=" & m_extRefRotateDir)
    uniFile.WriteUnicodeLine ("RotateIndexSteps=" & m_extRefRotateIndexSteps)
    uniFile.WriteUnicodeLine ("RotateMoveMode=" & m_extRefRotateMoveMode)
    uniFile.WriteUnicodeLine ("RotateSpeed=" & m_extRefRotateSpeed)
    uniFile.WriteUnicodeLine ("RotateStepSteps=" & m_extRefRotateStepSteps)
    uniFile.WriteUnicodeLine ("StartWvln=" & m_extRefStartWvln)
    uniFile.Flush
    uniFile.CloseFile
        
    If (uniFile.st_FileExist(m_extRefFileName) = True) Then
      uniFile.st_SetFileAttr m_extRefFileName, vbNormal
      uniFile.st_RmFile m_extRefFileName
    End If
  
    uniFile.st_CopyFile tmpFile, m_extRefFileName
    uniFile.st_RmFile tmpFile
  
    frm_collect.build_ref_name_list "external"
  Else
FILE_ERROR:
    uniFile.CloseFile
  
    If (tmpFile <> "") And (uniFile.st_FileExist(tmpFile) = True) Then
      uniFile.st_RmFile tmpFile
    End If
 
    errMsg = (m_extRefFileName & " file write error. " & Error$)
    unity_main.errorstring = errMsg
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", m_extRefFileName, Error$)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  End If
End Sub

Public Sub setup_ext_ref_tray_cfg()
  
  m_ignoreEvent = True

  ' Setup adapter type and get tray number
  setup_adapter_tray_num
  
  ' Setup tray rotation direction
  Select Case (m_extRefRotateDir)
    Case TRD_CW         ' Clockwise
      opt_rotateDirCW.Value = True
    Case TRD_NONE       ' no rotation
      opt_rotateDirCW.Value = False
      opt_rotateDirCCW.Value = False
    Case TRD_CCW        ' Counter Clockwise
      opt_rotateDirCCW.Value = True
    Case Else           ' invalid value
      If (m_dfltSettings = False) Then
        unity_main.errorstring = (m_extRefFileName & " had incompatible value. RotateDir was " & m_extRefRotateDir & "; updated to " & ProdDfltData.rotateDir)
        unity_main.write_error
        m_badIniVal = True
      End If
        
      m_extRefRotateDir = ProdDfltData.rotateDir
      m_extRefRotateMoveMode = ProdDfltData.rotateMoveMode
  End Select
  
  ' Setup tray movement
  Select Case (m_extRefRotateMoveMode)
    Case TRM_NONE                ' no movement
    Case TRM_CONT                ' continuous movement
      ' Check for invalid velocity value
      If (m_extRefRotateSpeed < MS11TrayInfoData(m_extRefTrayNum).minVel) Or (m_extRefRotateSpeed > MS11TrayInfoData(m_extRefTrayNum).maxVel) Then
        If (m_dfltSettings = False) Then
          unity_main.errorstring = (m_extRefFileName & " had incompatible value. RotateSpeed was " & m_extRefRotateSpeed & "; updated to " & MS11DfltTrayCfgData(m_extRefTrayNum).velCont)
          unity_main.write_error
          m_badIniVal = True
        End If
          
        unity_main.m_extRefRotateSpeed = MS11DfltTrayCfgData(m_extRefTrayNum).velCont
      End If
    Case TRM_STEP                ' stepped movement
      ' Check for invalid number of steps for stepped samples
      If ((m_extRefRotateStepSteps < 1) And (MS11TrayInfoData(m_extRefTrayNum).maxStps4scn <> 0)) Or (m_extRefRotateStepSteps > MS11TrayInfoData(m_extRefTrayNum).maxStps4scn) Then
        If (m_dfltSettings = False) Then
          unity_main.errorstring = (m_extRefFileName & " had incompatible value. RotateStepSteps was " & m_extRefRotateStepSteps & "; updated to " & MS11DfltTrayCfgData(m_extRefTrayNum).stps4scn)
          unity_main.write_error
          m_badIniVal = True
        End If
          
        m_extRefRotateStepSteps = MS11DfltTrayCfgData(m_extRefTrayNum).stps4scn
      End If
    Case TRM_INDEX               ' indexed movement
      ' Check for invalid number of steps for indexed samples
      If ((m_extRefRotateIndexSteps < 1) And (MS11TrayInfoData(m_extRefTrayNum).maxStps4IX <> 0)) Or (m_extRefRotateIndexSteps > MS11TrayInfoData(m_extRefTrayNum).maxStps4IX) Then
        If (m_dfltSettings = False) Then
          unity_main.errorstring = (m_extRefFileName & " had incompatible value. RotateIndexSteps was " & m_extRefRotateIndexSteps & "; updated to " & MS11DfltTrayCfgData(m_extRefTrayNum).stps4IX)
          unity_main.write_error
          m_badIniVal = True
        End If
          
        m_extRefRotateIndexSteps = MS11DfltTrayCfgData(m_extRefTrayNum).stps4IX
      End If
    Case Else                       ' invalid value
      If (m_dfltSettings = False) Then
        unity_main.errorstring = (m_extRefFileName & " had incompatible value. RotateMoveMode was " & m_extRefRotateMoveMode & "; updated to " & ProdDfltData.rotateMoveMode)
        unity_main.write_error
        m_badIniVal = True
      End If
      
      m_extRefRotateMoveMode = ProdDfltData.rotateMoveMode
      m_extRefRotateDir = ProdDfltData.rotateDir

  End Select
  
  numInc_rotateIndexSteps.Text = m_extRefRotateIndexSteps
  numInc_rotateSpeed.Text = m_extRefRotateSpeed
  numInc_rotateStepSteps.Text = m_extRefRotateStepSteps

  m_ignoreEvent = False
End Sub

Public Sub setup_adapter_tray_num()
  Dim adaptIndx As Integer

  ' Get tray number for instrument's adapter type
PROCESS_ADAPTERTYPE:
  For adaptIndx = 1 To MAX_ADAPTER_TYPES
    If (m_extRefAdapterType = MS11AdapterInfo(adaptIndx).cfgName) Then
      m_extRefTrayNum = MS11AdapterInfo(adaptIndx).trayNum
      m_extRefAdaptIndx = adaptIndx
      GoTo LEAVE_LOOP
    End If
  Next adaptIndx

  ' Flag invalid adapter type
  If (m_dfltSettings = False) Then
    unity_main.errorstring = (m_extRefFileName & " had incompatible value. AdapterType was " & m_extRefAdapterType & "; updated to " & ProdDfltData.adapterType)
    unity_main.write_error
    m_badIniVal = True
  End If
    
  m_extRefAdapterType = ProdDfltData.adapterType
  GoTo PROCESS_ADAPTERTYPE

LEAVE_LOOP:
  Select Case (MS11CfgData.devID)
    Case DTID_DRAWER0            ' SS2200/SS2400 standard drawer system
      opt_adapterType(adaptIndx).Value = True
      disable_multicup_types
      setup_rotate_platter
      
    Case DTID_TOPWIND0           ' Top window w/out internal reflectance
      opt_adapterType(adaptIndx).Value = True
      disable_multicup_types
      setup_rotate_platter

    Case DTID_DRAWER1            ' SS2200/SS2400 drawer w/out stepper system
      opt_adapterType(adaptIndx).Value = True
      disable_multicup_types
      setup_rotate_platter

    Case DTID_TOPWIND1           ' Top window with internal reflectance
      Select Case (m_extRefAdapterType)
        Case CFG_SAMP_WINDOW_AT
          opt_adapterType(adaptIndx).Value = True
          disable_multicup_types
          setup_rotate_platter
          
        Case CFG_IRIS_AT
          opt_adapterType(adaptIndx).Value = True
          disable_multicup_types
          setup_rotate_platter
          
        Case CFG_ISI_RING_AT
          opt_adapterType(adaptIndx).Value = True
          disable_multicup_types
          setup_rotate_platter
          
        Case CFG_SINGLE_POS_ROTATE_AT
          opt_adapterType(adaptIndx).Value = True
          disable_multicup_types
          setup_rotate_platter
          
        Case CFG_MULTI_CUP_AT
          opt_adapterType(adaptIndx).Value = True
          frame_multiCupType.Visible = True
          setup_multi_cup_type
      End Select
      
    Case Else                   ' Unknown dev ID
      disable_multicup_types
      frame_rotatePlatter.Visible = False
  End Select
End Sub

Public Sub disable_multicup_types()
  Dim nn As Integer
  
  frame_multiCupType.Visible = False
  m_extRefMultiCupType = CFG_NONE_MCT
  
  For nn = 1 To MAX_MULTI_CUP_TYPES
    opt_multiCupType(nn).Value = False
  Next nn
End Sub

Public Sub setup_multi_cup_type()
  Dim multiCupIndx As Integer

  ' Display possible multi-cup selections supported by instrument
  For multiCupIndx = 1 To MAX_MULTI_CUP_TYPES
    opt_multiCupType(multiCupIndx).Value = False

    If (MS11MultiCupInfo(multiCupIndx).trayNum > 0) Then
      If (MS11TrayInfoData(MS11MultiCupInfo(multiCupIndx).trayNum).trayID >= TTID_48POS) Then
        opt_multiCupType(multiCupIndx).Visible = True
      Else
        opt_multiCupType(multiCupIndx).Visible = False
      End If
    Else
      opt_multiCupType(multiCupIndx).Visible = False
    End If
  Next multiCupIndx

  ' Get tray number for instrument's multi-cup type
PROCESS_MULTICUPTYPE:
  For multiCupIndx = 0 To MAX_MULTI_CUP_TYPES
    If (m_extRefMultiCupType = MS11MultiCupInfo(multiCupIndx).cfgName) Then
      m_extRefMultiCupIndx = multiCupIndx
      m_extRefTrayNum = MS11MultiCupInfo(multiCupIndx).trayNum
      If (multiCupIndx <> 0) Then
        opt_multiCupType(multiCupIndx).Value = True
      End If
      GoTo LEAVE_LOOP
    End If
  Next multiCupIndx
  
  ' Flag invalid multi-cup type
  If (m_dfltSettings = False) Then
    unity_main.errorstring = (m_extRefFileName & " had incompatible value. MultiCupType was " & m_extRefMultiCupType & "; updated to " & ProdDfltData.multiCupType)
    unity_main.write_error
    m_badIniVal = True
  End If
    
  m_extRefMultiCupType = ProdDfltData.multiCupType
  GoTo PROCESS_MULTICUPTYPE
 
LEAVE_LOOP:
  setup_rotate_platter
End Sub

Public Sub setup_rotate_platter()
  Dim indx As Integer
  
  ' Setup plater info only for valid tray number
  If (m_extRefTrayNum > 0) And (m_extRefTrayNum <= MS11CfgData.nTrays) Then
    frame_rotatePlatter.Visible = True
    indx = m_extRefTrayNum
  
    ' Check if tray supports continuous movement
    If (MS11TrayInfoData(indx).maxVel > 0) Then
      ' Setup continuous movement cfg info
      opt_rotateModeCont.Visible = True
      numInc_rotateSpeed.Min = MS11TrayInfoData(indx).minVel
      numInc_rotateSpeed.Max = MS11TrayInfoData(indx).maxVel
    Else
      opt_rotateModeCont.Visible = False
      numInc_rotateSpeed.Min = 0
      numInc_rotateSpeed.Max = 0
    End If
    
    ' Check if tray supports stepped movement
    If (MS11TrayInfoData(indx).maxStps4scn > 0) Then
      ' Setup stepped movement cfg info
      opt_rotateModeStep.Visible = True
      numInc_rotateStepSteps.Min = 1
      numInc_rotateStepSteps.Max = MS11TrayInfoData(indx).maxStps4scn
      txt_cmplRevSteps.Text = MS11TrayInfoData(indx).nstps4rev
    Else
      opt_rotateModeStep.Visible = False
      numInc_rotateStepSteps.Min = 0
      numInc_rotateStepSteps.Max = 0
    End If
    
    ' Check if tray supports indexed movement
    If (MS11TrayInfoData(indx).maxStps4IX > 0) Then
      ' Setup indexed movement cfg info
      opt_rotateModeIndex.Visible = True
      numInc_rotateIndexSteps.Min = 1
      numInc_rotateIndexSteps.Max = MS11TrayInfoData(indx).maxStps4IX
      txt_cmplRevSteps.Text = MS11TrayInfoData(indx).nstps4rev
    Else
      opt_rotateModeIndex.Visible = False
      numInc_rotateIndexSteps.Min = 0
      numInc_rotateIndexSteps.Max = 0
    End If
    
    opt_rotateModeNone.Visible = True
  Else
    frame_rotatePlatter.Visible = False
  End If

   setup_rotate_move_mode
End Sub
  
Public Sub setup_rotate_move_mode()
  
  Select Case (m_extRefRotateMoveMode)
    Case TRM_NONE        ' None
      m_extRefRotateDir = TRD_NONE
      opt_rotateModeNone.Value = True
      frame_rotateDir.Visible = False
      opt_rotateDirCW.Value = False
      opt_rotateDirCCW.Value = False
      numInc_rotateSpeed.Visible = False
      lbl_rotateSpeed.Visible = False
      numInc_rotateStepSteps.Visible = False
      lbl_rotateStepSteps.Visible = False
      numInc_rotateIndexSteps.Visible = False
      lbl_rotateIndexSteps.Visible = False
      txt_cmplRevSteps.Visible = False
      lbl_cmplRevSteps.Visible = False
    
    Case TRM_CONT        ' Continuous
      If (numInc_rotateSpeed.Max > 0) Then
        opt_rotateModeCont.Value = True
        frame_rotateDir.Visible = True
        numInc_rotateSpeed.Visible = True
        lbl_rotateSpeed.Visible = True
        numInc_rotateStepSteps.Visible = False
        lbl_rotateStepSteps.Visible = False
        numInc_rotateIndexSteps.Visible = False
        lbl_rotateIndexSteps.Visible = False
        txt_cmplRevSteps.Visible = False
        lbl_cmplRevSteps.Visible = False
        
        ' Set value to default if have not been configured before
        If (m_extRefRotateSpeed = 0) Then
          numInc_rotateSpeed.Text = MS11DfltTrayCfgData(m_extRefTrayNum).velCont
        End If
      Else
        If (m_dfltSettings = False) Then
          unity_main.errorstring = (m_extRefFileName & " had incompatible value. RotateMoveMode was " & m_extRefRotateMoveMode & "; updated to " & ProdDfltData.rotateMoveMode)
          unity_main.write_error
          m_badIniVal = True
        End If
          
        m_extRefRotateMoveMode = ProdDfltData.rotateMoveMode
        m_extRefRotateDir = ProdDfltData.rotateDir
        setup_rotate_move_mode
      End If
    
    Case TRM_STEP        ' Stepped
      If (numInc_rotateStepSteps.Max > 0) Then
        opt_rotateModeStep.Value = True
        frame_rotateDir.Visible = True
        numInc_rotateSpeed.Visible = False
        lbl_rotateSpeed.Visible = False
        numInc_rotateStepSteps.Visible = True
        lbl_rotateStepSteps.Visible = True
        numInc_rotateIndexSteps.Visible = False
        lbl_rotateIndexSteps.Visible = False
        txt_cmplRevSteps.Visible = True
        lbl_cmplRevSteps.Visible = True
        
        ' Set value to default if have not been configured before
        If (m_extRefRotateStepSteps = 0) Then
          numInc_rotateStepSteps.Text = MS11DfltTrayCfgData(m_extRefTrayNum).stps4scn
        End If
      Else
        If (m_dfltSettings = False) Then
          unity_main.errorstring = (m_extRefFileName & " had incompatible value. RotateMoveMode was " & m_extRefRotateMoveMode & "; updated to " & ProdDfltData.rotateMoveMode)
          unity_main.write_error
          m_badIniVal = True
        End If
          
        m_extRefRotateMoveMode = ProdDfltData.rotateMoveMode
        m_extRefRotateDir = ProdDfltData.rotateDir
        setup_rotate_move_mode
      End If
    
    Case TRM_INDEX       ' Indexed
      If (numInc_rotateIndexSteps.Max > 0) Then
        opt_rotateModeIndex.Value = True
        frame_rotateDir.Visible = True
        numInc_rotateSpeed.Visible = False
        lbl_rotateSpeed.Visible = False
        numInc_rotateStepSteps.Visible = False
        lbl_rotateStepSteps.Visible = False
        numInc_rotateIndexSteps.Visible = True
        lbl_rotateIndexSteps.Visible = True
        txt_cmplRevSteps.Visible = True
        lbl_cmplRevSteps.Visible = True
        
        ' Set value to default if have not been configured before
        If (m_extRefRotateIndexSteps = 0) Then
          numInc_rotateIndexSteps.Text = MS11DfltTrayCfgData(m_extRefTrayNum).stps4IX
        End If
      Else
        If (m_dfltSettings = False) Then
          unity_main.errorstring = (m_extRefFileName & " had incompatible value. RotateMoveMode was " & m_extRefRotateMoveMode & "; updated to " & ProdDfltData.rotateMoveMode)
          unity_main.write_error
          m_badIniVal = True
        End If
          
        m_extRefRotateMoveMode = ProdDfltData.rotateMoveMode
        m_extRefRotateDir = ProdDfltData.rotateDir
        setup_rotate_move_mode
      End If
      
    Case Else               ' Unknown
      If (m_dfltSettings = False) Then
        unity_main.errorstring = (m_extRefFileName & " had incompatible value. RotateMoveMode was " & m_extRefRotateMoveMode & "; updated to " & ProdDfltData.rotateMoveMode)
        unity_main.write_error
        m_badIniVal = True
      End If
        
      m_extRefRotateMoveMode = ProdDfltData.rotateMoveMode
      m_extRefRotateDir = ProdDfltData.rotateDir
      setup_rotate_move_mode
  End Select
End Sub

Function setup_ext_ref(extRefFileName As String) As Boolean
  Dim rc As Boolean
  Dim spcFilename As String
  Dim spcIO As GSpcIOLib.GSPCio
  Dim numSubfiles As Long
  Dim errStrg As String
  Dim subFileIndx As Long
  Dim varXVals As Variant
  Dim varYVals As Variant
   
  ' Load reference qualification file
  spcFilename = (REFERENCES_DIR & extRefFileName & SPC_FILE_EXT)
  rc = LoadSpcFile(spcFilename, spcIO, numSubfiles, errStrg)
  
  If (rc = True) Then
    ' Get reference spectrum data
    subFileIndx = 0
    rc = GetSpcFileSpectrumData(spcIO, subFileIndx, varXVals, varYVals, errStrg)
    
    If (rc = True) Then
      Dim numPts As Integer
      Dim nn As Integer
      
      numPts = (spcIO.LastPoint - spcIO.FirstPoint) / MS11CfgData.wvlnIncr
      ReDim ProdRefXVals(numPts)
      ReDim ProdRefYVals(numPts)
      
      For nn = 0 To numPts
        ProdRefXVals(nn) = varXVals(nn)
        ProdRefYVals(nn) = varYVals(nn)
      Next nn
        
      ' clear any previous errors
      Clear_MS11_Error_Codes
        
      ' Setup reference spectrum data for product scan
#If SSRCS Then
      SSRCSClientError = unity_main.SSRCSClient.SetRefScan((numPts + 1), ProdRefYVals(0))
      
      If (SSRCSClientError <> 0) Then
        rc = False
#Else
      rc = unity_main.MS11srv.SetRefScan(ProdRefYVals(0))
      
      If (rc = False) Then
#End If
        Call Get_MS11_Errorcodes_Msg(errStrg)
        unity_main.m_ansiErrMsg = "Error setting reference scan data"
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status95", "Error setting external reference scan data")
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
  
  If (rc = False) Then
    unity_main.errorstring = unity_main.m_ansiErrMsg
    unity_main.write_error
  End If
  
  setup_ext_ref = rc
End Function

Private Sub chk_dfltWavelens_Click()

  If (chk_dfltWavelens.Value = 0) Then
    txt_dfltStartWvln.Visible = False
    txt_dfltEndWvln.Visible = False
    txt_startWvln.Visible = True
    txt_endWvln.Visible = True
    txt_minWvln.Visible = True
    lbl_minWvln.Visible = True
    txt_maxWvln.Visible = True
    lbl_maxWvln.Visible = True
  Else
    txt_dfltStartWvln.Visible = True
    txt_dfltEndWvln.Visible = True
    txt_startWvln.Visible = False
    txt_endWvln.Visible = False
    txt_minWvln.Visible = False
    lbl_minWvln.Visible = False
    txt_maxWvln.Visible = False
    lbl_maxWvln.Visible = False
  End If
End Sub

Private Sub cmd_cancel_Click()

  unity_main.errorstring = "External Reference Configuration screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  frm_extRef.Visible = False
End Sub

Private Sub cmd_save_Click()
  
  ' Check if valid configuration
  If (check_ext_ref_settings = True) Then
    unity_main.errorstring = "External Reference Configuration screen Save Changes button selected"
    unity_main.write_error (LOG_DBG_LEVEL3)
    
    frm_extRef.save_ext_ref_settings
    frm_extRef.Visible = False
    
    unity_main.errorstring = ("User saved new settings for external reference configuration file: " & m_extRefFileName)
    unity_main.write_error (LOG_DBG_LEVEL1)
    
    Call unity_main.load_prod_file("", True)
    unity_main.txtsamplename.Text = ""
    unity_main.txtsampcomment.Text = ""
    unity_main.repcounter = 0
  End If
End Sub

Private Sub Form_Load()
  Dim ii As Integer

  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
  ' Setup tab headers
  tab_frame.AddTab MLSupport.GSS("frm_extRef", "TabCaption0", "Scan Settings")
  tab_frame.AddTab MLSupport.GSS("frm_extRef", "TabCaption1", "Tray Settings")
  
  ' Setup to display first tab frame
  show_frame 0
End Sub

Private Sub numInc_refTimeout_DblClick()
  
  unity_main.formfrom = 16
  unity_main.varfrom = 1
  frm_numpad.lbl_num.Caption = lbl_refTimeout.Caption
  frm_numpad.txt_num.Text = numInc_refTimeout.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_refNScans_DblClick()
  
  unity_main.formfrom = 16
  unity_main.varfrom = 3
  frm_numpad.lbl_num.Caption = lbl_refNScans.Caption
  frm_numpad.txt_num.Text = numInc_refNScans.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_refPPT_DblClick()
  
  unity_main.formfrom = 16
  unity_main.varfrom = 2
  frm_numpad.lbl_num.Caption = lbl_refPPT.Caption
  frm_numpad.txt_num.Text = numInc_refPPT.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_rotateIndexSteps_DblClick()
  
  unity_main.formfrom = 16
  unity_main.varfrom = 6
  frm_numpad.lbl_num.Caption = lbl_rotateIndexSteps.Caption
  frm_numpad.txt_num.Text = numInc_rotateIndexSteps.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_rotateSpeed_DblClick()
  
  unity_main.formfrom = 16
  unity_main.varfrom = 7
  frm_numpad.lbl_num.Caption = lbl_rotateSpeed.Caption
  frm_numpad.txt_num.Text = numInc_rotateSpeed.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_rotateStepSteps_DblClick()
  
  unity_main.formfrom = 16
  unity_main.varfrom = 8
  frm_numpad.lbl_num.Caption = lbl_rotateStepSteps.Caption
  frm_numpad.txt_num.Text = numInc_rotateStepSteps.Text
  frm_numpad.Show 1
End Sub

Private Sub opt_adapterType_Click(Index As Integer)
  
  If (m_ignoreEvent = False) Then
    m_ignoreEvent = True
    m_extRefAdapterType = MS11AdapterInfo(Index).cfgName
    m_extRefRotateMoveMode = TRM_NONE
    setup_adapter_tray_num
    m_ignoreEvent = False
  End If
End Sub

Private Sub opt_multiCupType_Click(Index As Integer)
  
  If (m_ignoreEvent = False) Then
    m_ignoreEvent = True
    m_extRefMultiCupType = MS11MultiCupInfo(Index).cfgName
    m_extRefRotateMoveMode = TRM_NONE
    setup_multi_cup_type
    m_ignoreEvent = False
  End If
End Sub

Private Sub opt_rotateDirCCW_Click()
  
  If (m_ignoreEvent = False) Then
    m_extRefRotateDir = TRD_CCW
  End If
End Sub

Private Sub opt_rotateDirCW_Click()
  
  If (m_ignoreEvent = False) Then
    m_extRefRotateDir = TRD_CW
  End If
End Sub

Private Sub opt_rotateModeCont_Click()
  
  If (m_ignoreEvent = False) Then
    m_ignoreEvent = True
    m_extRefRotateMoveMode = TRM_CONT
    setup_rotate_move_mode
    m_ignoreEvent = False
  End If
End Sub

Private Sub opt_rotateModeIndex_Click()
  
  If (m_ignoreEvent = False) Then
    m_ignoreEvent = True
    m_extRefRotateMoveMode = TRM_INDEX
    setup_rotate_move_mode
    m_ignoreEvent = False
  End If
End Sub

Private Sub opt_rotateModeNone_Click()
  
  If (m_ignoreEvent = False) Then
    m_ignoreEvent = True
    m_extRefRotateMoveMode = TRM_NONE
    m_extRefRotateDir = TRD_NONE
    setup_rotate_move_mode
    m_ignoreEvent = False
  End If
End Sub

Private Sub opt_rotateModeStep_Click()
  
  If (m_ignoreEvent = False) Then
    m_ignoreEvent = True
    m_extRefRotateMoveMode = TRM_STEP
    setup_rotate_move_mode
    m_ignoreEvent = False
  End If
End Sub

Private Sub show_frame(n As Integer)
  Dim ii As Integer
  
  For ii = 0 To frame_main.Count - 1
    If ii = n Then
      frame_main(ii).Visible = True
    Else
      frame_main(ii).Visible = False
    End If
  Next
End Sub

Private Sub tab_frame_TabChanged(Index As Integer)

  show_frame (Index - 1)
End Sub

Private Sub txt_endWvln_DblCLick(Button As Integer)
  
  unity_main.formfrom = 16
  unity_main.varfrom = 5
  frm_numpad.lbl_num.Caption = lbl_endWvln.Caption & " " & MLSupport.GSS("frm_extRef", "lbl_num1", "Wavelength")
  frm_numpad.txt_num.Text = txt_endWvln.Text
  frm_numpad.Show 1
End Sub

Private Sub txt_extRefDesc_DblCLick(Button As Integer)

  unity_main.formfrom = 16
  unity_main.varfrom = 2
  frm_kybd.lbl_kybd.Caption = Label2.Caption
  frm_kybd.txt_kybd.Text = txt_extRefDesc.Text
  frm_kybd.Show 1
End Sub

Private Sub txt_extRefName_DblCLick(Button As Integer)

  unity_main.formfrom = 16
  unity_main.varfrom = 1
  frm_kybd.lbl_kybd.Caption = Label1.Caption
  frm_kybd.txt_kybd.Text = txt_extRefName.Text
  frm_kybd.Show 1
End Sub

Private Sub txt_startWvln_DblCLick(Button As Integer)
  
  unity_main.formfrom = 16
  unity_main.varfrom = 4
  frm_numpad.lbl_num.Caption = lbl_startWvln.Caption & " " & MLSupport.GSS("frm_extRef", "lbl_num1", "Wavelength")
  frm_numpad.txt_num.Text = txt_startWvln.Text
  frm_numpad.Show 1
End Sub






