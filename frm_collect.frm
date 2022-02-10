VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "resize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_collect 
   Caption         =   "Product Configuration"
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin HexUniControls.ctlUniFrameXP frame_main 
      Height          =   7005
      Index           =   0
      Left            =   480
      Top             =   795
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
      Caption         =   "frm_collect.frx":0000
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483642
      Tip             =   "frm_collect.frx":0020
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_collect.frx":0040
      Begin HexUniControls.ctlUniFrameXP frame_prediction 
         Height          =   3015
         Left            =   3360
         Top             =   2880
         Width           =   4215
         _ExtentX        =   7435
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
         Caption         =   "frm_collect.frx":005C
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":0090
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":00B0
         Begin HexUniControls.ctlUniFrameXP frame_wavenumIndices 
            Height          =   1455
            Left            =   120
            Top             =   1440
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   2566
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frm_collect.frx":00CC
            Enabled         =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   16711680
            Tip             =   "frm_collect.frx":0110
            VistaStyle      =   -1  'True
            UseShadow       =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":0130
            Begin HexUniControls.ctlUniLabel lbl_startWavenum 
               Height          =   285
               Left            =   120
               Top             =   480
               Width           =   1005
               _ExtentX        =   1773
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
               Caption         =   "frm_collect.frx":014C
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frm_collect.frx":017C
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":019C
            End
            Begin HexUniControls.ctlUniTextBoxXP txt_startWavenumIndx 
               Height          =   375
               Left            =   1200
               TabIndex        =   84
               Top             =   480
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frm_collect.frx":01B8
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
               Tip             =   "frm_collect.frx":01D8
               NoHideSel       =   0   'False
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":01F8
            End
            Begin HexUniControls.ctlUniLabel lbl_endWavenum 
               Height          =   285
               Left            =   120
               Top             =   960
               Width           =   1005
               _ExtentX        =   1773
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
               Caption         =   "frm_collect.frx":0214
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frm_collect.frx":0240
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":0260
            End
            Begin HexUniControls.ctlUniTextBoxXP txt_endWavenumIndx 
               Height          =   375
               Left            =   1200
               TabIndex        =   85
               Top             =   960
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Enabled         =   -1  'True
               Locked          =   0   'False
               Text            =   "frm_collect.frx":027C
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
               Tip             =   "frm_collect.frx":029C
               NoHideSel       =   0   'False
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":02BC
            End
            Begin HexUniControls.ctlUniTextBoxXP txt_startWavenum 
               Height          =   375
               Left            =   2040
               TabIndex        =   83
               Top             =   480
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
               BackColor       =   -2147483633
               ForeColor       =   -2147483640
               Enabled         =   0   'False
               Locked          =   0   'False
               Text            =   "frm_collect.frx":02D8
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
               MultiLine       =   0   'False
               Alignment       =   1
               ScrollBars      =   0
               PasswordChar    =   ""
               TrapTab         =   0   'False
               EnableContextMenu=   -1  'True
               RaiseChangeEvent=   -1  'True
               Tip             =   "frm_collect.frx":02F8
               NoHideSel       =   0   'False
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":0318
            End
            Begin HexUniControls.ctlUniTextBoxXP txt_endWavenum 
               Height          =   375
               Left            =   2040
               TabIndex        =   86
               Top             =   960
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
               BackColor       =   -2147483633
               ForeColor       =   -2147483640
               Enabled         =   0   'False
               Locked          =   0   'False
               Text            =   "frm_collect.frx":0334
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
               MultiLine       =   0   'False
               Alignment       =   1
               ScrollBars      =   0
               PasswordChar    =   ""
               TrapTab         =   0   'False
               EnableContextMenu=   -1  'True
               RaiseChangeEvent=   -1  'True
               Tip             =   "frm_collect.frx":0354
               NoHideSel       =   0   'False
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":0374
            End
            Begin HexUniControls.ctlUniLabel Label23 
               Height          =   375
               Left            =   3660
               Top             =   840
               Width           =   255
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
               Caption         =   "frm_collect.frx":0390
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frm_collect.frx":03B4
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":03D4
            End
            Begin HexUniControls.ctlUniLabel Label22 
               Height          =   375
               Left            =   3660
               Top             =   360
               Width           =   255
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
               Caption         =   "frm_collect.frx":03F0
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frm_collect.frx":0414
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":0434
            End
            Begin HexUniControls.ctlUniLabel Label21 
               Height          =   255
               Left            =   3360
               Top             =   960
               Width           =   375
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
               Caption         =   "frm_collect.frx":0450
               BackColor       =   -2147483633
               ForeColor       =   0
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frm_collect.frx":0474
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":0494
            End
            Begin HexUniControls.ctlUniLabel Label20 
               Height          =   255
               Left            =   3360
               Top             =   480
               Width           =   375
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
               Caption         =   "frm_collect.frx":04B0
               BackColor       =   -2147483633
               ForeColor       =   0
               Alignment       =   0
               VAlignment      =   0
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frm_collect.frx":04D4
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":04F4
            End
         End
         Begin HexUniControls.ctlUniLabel lbl_smplNumPts 
            Height          =   285
            Left            =   120
            Top             =   480
            Width           =   2205
            _ExtentX        =   3889
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
            Caption         =   "frm_collect.frx":0510
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":0550
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":0570
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_smplNumPts 
            Height          =   375
            Left            =   2400
            TabIndex        =   87
            Top             =   480
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   661
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "frm_collect.frx":058C
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
            MultiLine       =   0   'False
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frm_collect.frx":05AC
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":05CC
         End
         Begin HexUniControls.ctlUniLabel lbl_waveNumIncr 
            Height          =   285
            Left            =   120
            Top             =   960
            Width           =   2205
            _ExtentX        =   3889
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
            Caption         =   "frm_collect.frx":05E8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":0630
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":0650
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_waveNumIncr 
            Height          =   375
            Left            =   2400
            TabIndex        =   92
            Top             =   960
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   661
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "frm_collect.frx":066C
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
            MultiLine       =   0   'False
            Alignment       =   1
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frm_collect.frx":068C
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":06AC
         End
         Begin HexUniControls.ctlUniLabel Label19 
            Height          =   375
            Left            =   3720
            Top             =   840
            Width           =   255
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
            Caption         =   "frm_collect.frx":06C8
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":06EC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":070C
         End
         Begin HexUniControls.ctlUniLabel Label18 
            Height          =   255
            Left            =   3360
            Top             =   960
            Width           =   375
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
            Caption         =   "frm_collect.frx":0728
            BackColor       =   -2147483633
            ForeColor       =   0
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":074C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":076C
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_coAdditon 
         Height          =   2415
         Left            =   6840
         Top             =   285
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4260
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":0788
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":07BA
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":07DA
         Begin HexUniControls.ctlUniLabel lbl_numMeasures 
            Height          =   285
            Left            =   120
            Top             =   480
            Width           =   1965
            _ExtentX        =   3466
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
            Caption         =   "frm_collect.frx":07F6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":0838
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":0858
         End
         Begin HexUniControls.ctlUniLabel lbl_numSamples 
            Height          =   285
            Left            =   120
            Top             =   960
            Width           =   1965
            _ExtentX        =   3466
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
            Caption         =   "frm_collect.frx":0874
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":08B6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":08D6
         End
         Begin HexUniControls.ctlUniLabel lbl_delayStart 
            Height          =   285
            Left            =   120
            Top             =   1440
            Width           =   1965
            _ExtentX        =   3466
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
            Caption         =   "frm_collect.frx":08F2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":0932
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":0952
         End
         Begin HexUniControls.ctlUniLabel lbl_delayMeasure 
            Height          =   285
            Left            =   120
            Top             =   1920
            Width           =   1965
            _ExtentX        =   3466
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
            Caption         =   "frm_collect.frx":096E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":09B0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":09D0
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_numMeasures 
            Height          =   375
            Left            =   2160
            TabIndex        =   88
            Top             =   480
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   661
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "frm_collect.frx":09EC
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
            MultiLine       =   0   'False
            Alignment       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frm_collect.frx":0A0E
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":0A2E
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_delayStart 
            Height          =   375
            Left            =   2160
            TabIndex        =   89
            Top             =   1440
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   661
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_collect.frx":0A4A
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
            Tip             =   "frm_collect.frx":0A6A
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":0A8A
         End
         Begin HexUniControls.ctlNumIncXP numInc_numSamples 
            Height          =   495
            Left            =   2160
            TabIndex        =   90
            Top             =   900
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
            Text            =   "10"
            Min             =   1
            Max             =   25
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
            MouseIcon       =   "frm_collect.frx":0AA6
            TrapTabKey      =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_delayMeasure 
            Height          =   375
            Left            =   2160
            TabIndex        =   91
            Top             =   1860
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   661
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_collect.frx":0AC2
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
            Tip             =   "frm_collect.frx":0AE2
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":0B02
         End
         Begin HexUniControls.ctlUniLabel Label11 
            Height          =   345
            Left            =   3000
            Top             =   1440
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frm_collect.frx":0B1E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":0B46
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":0B66
         End
         Begin HexUniControls.ctlUniLabel Label12 
            Height          =   345
            Left            =   3000
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frm_collect.frx":0B82
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":0BAA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":0BCA
         End
      End
      Begin HexUniControls.ctlUniFrameXP frame_scanning 
         Height          =   1935
         Left            =   3360
         Top             =   285
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3413
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":0BE6
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":0C16
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":0C36
         Begin VB.ListBox lst_gain 
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frm_collect.frx":0C52
            Left            =   1680
            List            =   "frm_collect.frx":0C6B
            TabIndex        =   95
            Top             =   1440
            Width           =   855
         End
         Begin VB.ListBox lst_speed 
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frm_collect.frx":0C8B
            Left            =   1680
            List            =   "frm_collect.frx":0CA4
            TabIndex        =   94
            Top             =   960
            Width           =   855
         End
         Begin VB.ListBox lst_resolution 
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frm_collect.frx":0CD9
            Left            =   1680
            List            =   "frm_collect.frx":0CF2
            TabIndex        =   93
            Top             =   480
            Width           =   855
         End
         Begin HexUniControls.ctlUniLabel Label6 
            Height          =   285
            Left            =   120
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
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
            Caption         =   "frm_collect.frx":0D12
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":0D46
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":0D66
         End
         Begin HexUniControls.ctlUniLabel Label7 
            Height          =   285
            Left            =   120
            Top             =   960
            Width           =   1365
            _ExtentX        =   2408
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
            Caption         =   "frm_collect.frx":0D82
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":0DB4
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":0DD4
         End
         Begin HexUniControls.ctlUniLabel Label8 
            Height          =   285
            Left            =   120
            Top             =   1440
            Width           =   1365
            _ExtentX        =   2408
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
            Caption         =   "frm_collect.frx":0DF0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   1
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":0E24
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":0E44
         End
         Begin HexUniControls.ctlUniLabel Label10 
            Height          =   375
            Left            =   2640
            Top             =   960
            Width           =   495
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
            Caption         =   "frm_collect.frx":0E60
            BackColor       =   -2147483633
            ForeColor       =   0
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":0E88
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":0EA8
         End
         Begin HexUniControls.ctlUniLabel Label27 
            Height          =   255
            Left            =   2940
            Top             =   360
            Width           =   255
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
            Caption         =   "frm_collect.frx":0EC4
            BackColor       =   -2147483633
            ForeColor       =   0
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":0EE8
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":0F08
         End
         Begin HexUniControls.ctlUniLabel Label9 
            Height          =   255
            Left            =   2640
            Top             =   480
            Width           =   375
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
            Caption         =   "frm_collect.frx":0F24
            BackColor       =   -2147483633
            ForeColor       =   0
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":0F48
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":0F68
         End
      End
      Begin HexUniControls.ctlUniTextBoxXP txtreps 
         Height          =   375
         Left            =   10320
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   6540
         Visible         =   0   'False
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   661
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   -1  'True
         Text            =   "frm_collect.frx":0F84
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
         Tip             =   "frm_collect.frx":0FA6
         NoHideSel       =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":0FC6
      End
      Begin HexUniControls.ctlUniCheckXP chk_useExtRefTrayCfg 
         Height          =   495
         Left            =   3360
         TabIndex        =   16
         Top             =   5340
         Width           =   5655
         _ExtentX        =   9975
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
         Caption         =   "frm_collect.frx":0FE2
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frm_collect.frx":105A
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":107A
         ShowFocus       =   -1  'True
      End
      Begin HexUniControls.ctlUniFrameXP Frame5 
         Height          =   2730
         Left            =   3360
         Top             =   2580
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
         Caption         =   "frm_collect.frx":1096
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":10D6
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":10F6
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
            Caption         =   "frm_collect.frx":1112
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":1140
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":1160
         End
         Begin HexUniControls.ctlUniLabel lbl_minWvln 
            Height          =   285
            Left            =   1200
            Top             =   1760
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
            Caption         =   "frm_collect.frx":117C
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":11AA
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":11CA
         End
         Begin HexUniControls.ctlUniLabel lbl_endWvln 
            Height          =   285
            Left            =   1200
            Top             =   1360
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
            Caption         =   "frm_collect.frx":11E6
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":1212
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":1232
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
            Caption         =   "frm_collect.frx":124E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":127E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":129E
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_maxWvln 
            Height          =   375
            Left            =   120
            TabIndex        =   76
            Top             =   2160
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   661
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "frm_collect.frx":12BA
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
            Tip             =   "frm_collect.frx":12DA
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":12FA
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_minWvln 
            Height          =   375
            Left            =   120
            TabIndex        =   75
            Top             =   1760
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   661
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "frm_collect.frx":1316
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
            Tip             =   "frm_collect.frx":1336
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":1356
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_dfltEndWvln 
            Height          =   375
            Left            =   120
            TabIndex        =   74
            Top             =   1360
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   661
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "frm_collect.frx":1372
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
            Tip             =   "frm_collect.frx":1392
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":13B2
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_endWvln 
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   1360
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   661
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_collect.frx":13CE
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
            Tip             =   "frm_collect.frx":13EE
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":140E
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_dfltStartWvln 
            Height          =   375
            Left            =   120
            TabIndex        =   73
            Top             =   960
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   661
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "frm_collect.frx":142A
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
            Tip             =   "frm_collect.frx":144A
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":146A
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_startWvln 
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   661
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_collect.frx":1486
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
            Tip             =   "frm_collect.frx":14A6
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":14C6
         End
         Begin HexUniControls.ctlUniCheckXP chk_dfltWavelens 
            Height          =   525
            Left            =   120
            TabIndex        =   13
            Top             =   345
            Width           =   3060
            _ExtentX        =   5398
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
            Caption         =   "frm_collect.frx":14E2
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":152A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":154A
            ShowFocus       =   -1  'True
         End
      End
      Begin HexUniControls.ctlUniFrameXP Frame9 
         Height          =   2055
         Left            =   6960
         Top             =   285
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3625
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":1566
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":1594
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":15B4
         Begin HexUniControls.ctlUniCheckXP chk_savereps 
            Height          =   615
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Width           =   3195
            _ExtentX        =   5636
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
            Caption         =   "frm_collect.frx":15D0
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":162C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":164C
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniLabel lbl_numRepacks 
            Height          =   855
            Left            =   1440
            Top             =   360
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   1508
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frm_collect.frx":1668
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":16C0
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":16E0
         End
         Begin HexUniControls.ctlNumIncXP numInc_numRepacks 
            Height          =   615
            Left            =   120
            TabIndex        =   17
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
            Text            =   "100"
            Min             =   1
            Max             =   100
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
            MouseIcon       =   "frm_collect.frx":16FC
            TrapTabKey      =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameXP Frame15 
         Height          =   2055
         Left            =   3360
         Top             =   285
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3625
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":1718
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":1756
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":1776
         Begin HexUniControls.ctlUniLabel lbl_numSmplPPT 
            Height          =   615
            Left            =   1440
            Top             =   1080
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
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
            Caption         =   "frm_collect.frx":1792
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":17EC
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":180C
         End
         Begin HexUniControls.ctlNumIncXP numInc_smplPPT 
            Height          =   615
            Left            =   120
            TabIndex        =   12
            Top             =   1080
            Visible         =   0   'False
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
            Text            =   "0"
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
            MouseIcon       =   "frm_collect.frx":1828
            TrapTabKey      =   0   'False
         End
         Begin HexUniControls.ctlUniLabel lbl_smplNScans 
            Height          =   615
            Left            =   1440
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
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
            Caption         =   "frm_collect.frx":1844
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":1898
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":18B8
         End
         Begin HexUniControls.ctlNumIncXP numInc_smplNScans 
            Height          =   615
            Left            =   120
            TabIndex        =   11
            Top             =   360
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
            Text            =   "0"
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
            MouseIcon       =   "frm_collect.frx":18D4
            TrapTabKey      =   0   'False
         End
      End
      Begin HexUniControls.ctlUniFrameXP Frame13 
         Height          =   1485
         Left            =   120
         Top             =   4680
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   2619
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":18F0
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":1936
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":1956
         Begin HexUniControls.ctlUniRadioXP opt_backall 
            Height          =   450
            Left            =   120
            TabIndex        =   10
            Top             =   860
            Width           =   2700
            _ExtentX        =   4763
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
            Caption         =   "frm_collect.frx":1972
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":19AA
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":19CA
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP opt_backdemand 
            Height          =   450
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   2700
            _ExtentX        =   4763
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
            Caption         =   "frm_collect.frx":19E6
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":1A18
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":1A38
            ShowFocus       =   -1  'True
         End
      End
      Begin HexUniControls.ctlUniFrameXP Frame2 
         Height          =   2595
         Left            =   120
         Top             =   1920
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   4577
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":1A54
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":1A90
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":1AB0
         Begin HexUniControls.ctlUniFileBoxXP file_refNames 
            Height          =   255
            Left            =   720
            TabIndex        =   72
            Top             =   2520
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            SelBackColor    =   -2147483635
            SelForeColor    =   -2147483634
            RoundedBorders  =   0   'False
            SelectorStyle   =   -1
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
            Tip             =   "frm_collect.frx":1ACC
            Path            =   ""
            Pattern         =   "*.*"
            PatternAlsoForDirs=   0   'False
            ReadOnly        =   -1  'True
            System          =   0   'False
            Hidden          =   0   'False
            PermitNavigation=   -1  'True
            MultiSelect     =   0
            HScroll         =   0   'False
            ShowFullPath    =   0   'False
            DisplayMode     =   1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":1AEC
         End
         Begin HexUniControls.ctlUniComboBoxXP combo_olRefFileName 
            Height          =   450
            Left            =   120
            TabIndex        =   8
            Top             =   1860
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   794
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
            Tip             =   "frm_collect.frx":1B08
            Sorted          =   0   'False
            HScroll         =   -1  'True
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
            MouseIcon       =   "frm_collect.frx":1B28
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniComboBoxXP combo_extRefFileName 
            Height          =   450
            Left            =   120
            TabIndex        =   7
            Top             =   1860
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   794
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
            Tip             =   "frm_collect.frx":1B44
            Sorted          =   0   'False
            HScroll         =   -1  'True
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
            MouseIcon       =   "frm_collect.frx":1B64
            DropDownOnTextClick=   -1  'True
            DropDownWidth   =   -1
            ManualStart     =   0   'False
         End
         Begin HexUniControls.ctlUniRadioXP optbgfile 
            Height          =   450
            Left            =   120
            TabIndex        =   6
            Top             =   1360
            Width           =   2700
            _ExtentX        =   4763
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
            Caption         =   "frm_collect.frx":1B80
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":1BC2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":1BE2
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP optbgexternal 
            Height          =   450
            Left            =   120
            TabIndex        =   5
            Top             =   860
            Width           =   2700
            _ExtentX        =   4763
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
            Caption         =   "frm_collect.frx":1BFE
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":1C42
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":1C62
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP optbginternal 
            Height          =   450
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   2700
            _ExtentX        =   4763
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
            Caption         =   "frm_collect.frx":1C7E
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":1CC2
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":1CE2
            ShowFocus       =   -1  'True
         End
      End
      Begin HexUniControls.ctlUniFrameXP Frame1 
         Height          =   1485
         Left            =   120
         Top             =   285
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   2619
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":1CFE
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":1D38
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":1D58
         Begin HexUniControls.ctlUniRadioXP optbk 
            Height          =   450
            Left            =   120
            TabIndex        =   3
            Top             =   860
            Visible         =   0   'False
            Width           =   2700
            _ExtentX        =   4763
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
            Caption         =   "frm_collect.frx":1D74
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":1DA8
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":1DC8
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP optabs 
            Height          =   450
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   2700
            _ExtentX        =   4763
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
            Caption         =   "frm_collect.frx":1DE4
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":1E18
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":1E38
            ShowFocus       =   -1  'True
         End
      End
   End
   Begin HexUniControls.ctlUniFrameXP frame_main 
      Height          =   7005
      Index           =   4
      Left            =   480
      Top             =   795
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
      Caption         =   "frm_collect.frx":1E54
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483642
      Tip             =   "frm_collect.frx":1E74
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_collect.frx":1E94
      Begin HexUniControls.ctlUniFrameXP Frame17 
         Height          =   1995
         Left            =   360
         Top             =   2880
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":1EB0
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":1EFA
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":1F1A
         Begin HexUniControls.ctlUniCheckXP chk_hideValCol 
            Height          =   450
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   3855
            _ExtentX        =   6800
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
            Caption         =   "frm_collect.frx":1F36
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":1F78
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":1F98
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_noPredVal 
            Height          =   450
            Left            =   120
            TabIndex        =   82
            Top             =   1320
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_collect.frx":1FB4
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
            Tip             =   "frm_collect.frx":1FD8
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":1FF8
         End
         Begin HexUniControls.ctlUniLabel lbl_noPredVal 
            Height          =   345
            Left            =   120
            Top             =   960
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frm_collect.frx":2014
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":206C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":208C
         End
      End
      Begin HexUniControls.ctlUniFrameXP Frame10 
         Height          =   1485
         Left            =   4680
         Top             =   285
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2619
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":20A8
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":20DE
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":20FE
         Begin HexUniControls.ctlUniCheckXP chk_clrUserInputs 
            Height          =   450
            Left            =   120
            TabIndex        =   80
            Top             =   840
            Width           =   3855
            _ExtentX        =   6800
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
            Caption         =   "frm_collect.frx":211A
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":2178
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":2198
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniCheckXP chk_userInputs 
            Height          =   450
            Left            =   120
            TabIndex        =   70
            Top             =   360
            Width           =   3855
            _ExtentX        =   6800
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
            Caption         =   "frm_collect.frx":21B4
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":21F8
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":2218
            ShowFocus       =   -1  'True
         End
      End
      Begin HexUniControls.ctlUniFrameXP Frame12 
         Height          =   2445
         Left            =   360
         Top             =   285
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4313
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":2234
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":2280
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":22A0
         Begin HexUniControls.ctlUniLabel lbl_noOLVal 
            Height          =   345
            Left            =   120
            Top             =   1440
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frm_collect.frx":22BC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":230E
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":232E
         End
         Begin HexUniControls.ctlUniRadioXP opt_ollts 
            Height          =   450
            Left            =   120
            TabIndex        =   68
            Top             =   860
            Width           =   3855
            _ExtentX        =   6800
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
            Caption         =   "frm_collect.frx":234A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":23A4
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":23C4
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP opt_olval 
            Height          =   450
            Left            =   120
            TabIndex        =   67
            Top             =   360
            Width           =   3855
            _ExtentX        =   6800
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
            Caption         =   "frm_collect.frx":23E0
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":241A
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":243A
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_noOLVal 
            Height          =   450
            Left            =   120
            TabIndex        =   81
            Top             =   1800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_collect.frx":2456
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
            Tip             =   "frm_collect.frx":2478
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":2498
         End
      End
   End
   Begin HexUniControls.ctlUniFrameXP frame_main 
      Height          =   7005
      Index           =   1
      Left            =   480
      Top             =   795
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
      Caption         =   "frm_collect.frx":24B4
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483642
      Tip             =   "frm_collect.frx":24D4
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_collect.frx":24F4
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
         Caption         =   "frm_collect.frx":2510
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":2554
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":2574
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
            Caption         =   "frm_collect.frx":2590
            Enabled         =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   16711680
            Tip             =   "frm_collect.frx":25CC
            VistaStyle      =   -1  'True
            UseShadow       =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":25EC
            Begin HexUniControls.ctlUniRadioXP opt_multiCupType 
               Height          =   450
               Index           =   6
               Left            =   5040
               TabIndex        =   29
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
               Caption         =   "frm_collect.frx":2608
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_collect.frx":2668
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":2688
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_multiCupType 
               Height          =   450
               Index           =   5
               Left            =   5040
               TabIndex        =   28
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
               Caption         =   "frm_collect.frx":26A4
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_collect.frx":270C
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":272C
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_multiCupType 
               Height          =   450
               Index           =   4
               Left            =   5040
               TabIndex        =   27
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
               Caption         =   "frm_collect.frx":2748
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_collect.frx":279E
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":27BE
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_multiCupType 
               Height          =   450
               Index           =   3
               Left            =   120
               TabIndex        =   26
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
               Caption         =   "frm_collect.frx":27DA
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_collect.frx":2840
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":2860
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_multiCupType 
               Height          =   450
               Index           =   2
               Left            =   120
               TabIndex        =   25
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
               Caption         =   "frm_collect.frx":287C
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_collect.frx":28E8
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":2908
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_multiCupType 
               Height          =   450
               Index           =   1
               Left            =   120
               TabIndex        =   24
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
               Caption         =   "frm_collect.frx":2924
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_collect.frx":2992
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":29B2
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
            Caption         =   "frm_collect.frx":29CE
            Enabled         =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   16711680
            Tip             =   "frm_collect.frx":2A0E
            VistaStyle      =   -1  'True
            UseShadow       =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":2A2E
            Begin HexUniControls.ctlUniTextBoxXP txt_cmplRevSteps 
               Height          =   375
               Left            =   3840
               TabIndex        =   78
               Top             =   3360
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   661
               BackColor       =   -2147483633
               ForeColor       =   -2147483640
               Enabled         =   0   'False
               Locked          =   0   'False
               Text            =   "frm_collect.frx":2A4A
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
               Tip             =   "frm_collect.frx":2A6E
               NoHideSel       =   0   'False
               RightToLeft     =   0   'False
               ManualStart     =   0   'False
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":2A8E
            End
            Begin HexUniControls.ctlNumIncXP numInc_rotateIndexSteps 
               Height          =   600
               Left            =   3840
               TabIndex        =   38
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
               MouseIcon       =   "frm_collect.frx":2AAA
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
               Caption         =   "frm_collect.frx":2AC6
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   1
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frm_collect.frx":2B2E
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":2B4E
            End
            Begin HexUniControls.ctlNumIncXP numInc_rotateStepSteps 
               Height          =   600
               Left            =   3840
               TabIndex        =   37
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
               MouseIcon       =   "frm_collect.frx":2B6A
               TrapTabKey      =   0   'False
            End
            Begin HexUniControls.ctlNumIncXP numInc_rotateSpeed 
               Height          =   600
               Left            =   3840
               TabIndex        =   36
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
               MouseIcon       =   "frm_collect.frx":2B86
               TrapTabKey      =   0   'False
            End
            Begin HexUniControls.ctlUniFrameXP frame_rotateDir 
               Height          =   1395
               Left            =   2400
               Top             =   240
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
               Caption         =   "frm_collect.frx":2BA2
               Enabled         =   -1  'True
               BackColor       =   -2147483633
               ForeColor       =   16711680
               Tip             =   "frm_collect.frx":2BD4
               VistaStyle      =   -1  'True
               UseShadow       =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":2BF4
               Begin HexUniControls.ctlUniRadioXP opt_rotateDirCCW 
                  Height          =   450
                  Left            =   120
                  TabIndex        =   35
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
                  Caption         =   "frm_collect.frx":2C10
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frm_collect.frx":2C52
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frm_collect.frx":2C72
                  ShowFocus       =   -1  'True
               End
               Begin HexUniControls.ctlUniRadioXP opt_rotateDirCW 
                  Height          =   450
                  Left            =   120
                  TabIndex        =   34
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
                  Caption         =   "frm_collect.frx":2C8E
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frm_collect.frx":2CC0
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frm_collect.frx":2CE0
                  ShowFocus       =   -1  'True
               End
            End
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
               Caption         =   "frm_collect.frx":2CFC
               Enabled         =   -1  'True
               BackColor       =   -2147483633
               ForeColor       =   16711680
               Tip             =   "frm_collect.frx":2D36
               VistaStyle      =   -1  'True
               UseShadow       =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":2D56
               Begin HexUniControls.ctlUniRadioXP opt_rotateModeIndex 
                  Height          =   450
                  Left            =   120
                  TabIndex        =   33
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
                  Caption         =   "frm_collect.frx":2D72
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frm_collect.frx":2DA0
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frm_collect.frx":2DC0
                  ShowFocus       =   -1  'True
               End
               Begin HexUniControls.ctlUniRadioXP opt_rotateModeStep 
                  Height          =   450
                  Left            =   120
                  TabIndex        =   32
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
                  Caption         =   "frm_collect.frx":2DDC
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frm_collect.frx":2E0A
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frm_collect.frx":2E2A
                  ShowFocus       =   -1  'True
               End
               Begin HexUniControls.ctlUniRadioXP opt_rotateModeCont 
                  Height          =   450
                  Left            =   120
                  TabIndex        =   31
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
                  Caption         =   "frm_collect.frx":2E46
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frm_collect.frx":2E7A
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frm_collect.frx":2E9A
                  ShowFocus       =   -1  'True
               End
               Begin HexUniControls.ctlUniRadioXP opt_rotateModeNone 
                  Height          =   450
                  Left            =   120
                  TabIndex        =   30
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
                  Caption         =   "frm_collect.frx":2EB6
                  Enabled         =   -1  'True
                  Align           =   0
                  RadioBackColor  =   -2147483643
                  RadioForeColor  =   -2147483640
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  Pressed         =   0   'False
                  Tip             =   "frm_collect.frx":2EDE
                  Style           =   -1
                  MousePointer    =   0
                  MouseIcon       =   "frm_collect.frx":2EFE
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
               Caption         =   "frm_collect.frx":2F1A
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   1
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frm_collect.frx":2F82
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":2FA2
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
               Caption         =   "frm_collect.frx":2FBE
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   1
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frm_collect.frx":3026
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":3046
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
               Caption         =   "frm_collect.frx":3062
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Alignment       =   1
               VAlignment      =   1
               BackStyle       =   1
               BorderStyle     =   0
               AutoSize        =   0   'False
               Tip             =   "frm_collect.frx":30A2
               Style           =   0
               Enabled         =   -1  'True
               Margin          =   0
               RoundedBorders  =   0   'False
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":30C2
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
            Caption         =   "frm_collect.frx":30DE
            Enabled         =   -1  'True
            BackColor       =   -2147483633
            ForeColor       =   16711680
            Tip             =   "frm_collect.frx":3116
            VistaStyle      =   -1  'True
            UseShadow       =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":3136
            Begin HexUniControls.ctlUniRadioXP opt_adapterType 
               Height          =   450
               Index           =   5
               Left            =   120
               TabIndex        =   23
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
               Caption         =   "frm_collect.frx":3152
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_collect.frx":3194
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":31B4
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_adapterType 
               Height          =   450
               Index           =   4
               Left            =   120
               TabIndex        =   22
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
               Caption         =   "frm_collect.frx":31D0
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_collect.frx":3230
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":3250
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_adapterType 
               Height          =   450
               Index           =   3
               Left            =   120
               TabIndex        =   21
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
               Caption         =   "frm_collect.frx":326C
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_collect.frx":32AC
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":32CC
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_adapterType 
               Height          =   450
               Index           =   2
               Left            =   120
               TabIndex        =   20
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
               Caption         =   "frm_collect.frx":32E8
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_collect.frx":3320
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":3340
               ShowFocus       =   -1  'True
            End
            Begin HexUniControls.ctlUniRadioXP opt_adapterType 
               Height          =   450
               Index           =   1
               Left            =   120
               TabIndex        =   19
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
               Caption         =   "frm_collect.frx":335C
               Enabled         =   -1  'True
               Align           =   0
               RadioBackColor  =   -2147483643
               RadioForeColor  =   -2147483640
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               Pressed         =   0   'False
               Tip             =   "frm_collect.frx":3396
               Style           =   -1
               MousePointer    =   0
               MouseIcon       =   "frm_collect.frx":33B6
               ShowFocus       =   -1  'True
            End
         End
      End
   End
   Begin HexUniControls.ctlUniFrameXP frame_main 
      Height          =   7005
      Index           =   2
      Left            =   480
      Top             =   795
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
      Caption         =   "frm_collect.frx":33D2
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483642
      Tip             =   "frm_collect.frx":33F2
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_collect.frx":3412
      Begin HexUniControls.ctlUniFrameXP Frame4 
         Height          =   1815
         Left            =   5880
         Top             =   285
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   3201
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":342E
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":3468
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":3488
         Begin HexUniControls.ctlUniTextBoxXP txt_caldir 
            Height          =   450
            Left            =   120
            TabIndex        =   48
            Top             =   1200
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_collect.frx":34A4
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
            Tip             =   "frm_collect.frx":34C4
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":34E4
         End
         Begin HexUniControls.ctlUniLabel lbl_caldir 
            Height          =   300
            Left            =   120
            Top             =   840
            Width           =   4455
            _ExtentX        =   7858
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
            Caption         =   "frm_collect.frx":3500
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":3544
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":3564
         End
         Begin HexUniControls.ctlUniRadioXP optsavescanno 
            Height          =   450
            Left            =   1695
            TabIndex        =   47
            Top             =   360
            Width           =   1995
            _ExtentX        =   3519
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
            Caption         =   "frm_collect.frx":3580
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":35A4
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":35C4
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP optsavescanyes 
            Height          =   450
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   1400
            _ExtentX        =   2461
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
            Caption         =   "frm_collect.frx":35E0
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":3606
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":3626
            ShowFocus       =   -1  'True
         End
      End
      Begin HexUniControls.ctlUniFrameXP Frame8 
         Height          =   4935
         Left            =   360
         Top             =   285
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   8705
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":3642
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":368E
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":36AE
         Begin HexUniControls.ctlUniCheckXP chk_clrManualName 
            Height          =   450
            Left            =   375
            TabIndex        =   40
            Top             =   840
            Width           =   4800
            _ExtentX        =   8467
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
            Caption         =   "frm_collect.frx":36CA
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":3728
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":3748
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniTextBoxXP txt_scanname 
            Height          =   450
            Left            =   120
            TabIndex        =   79
            Top             =   4080
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   794
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            Enabled         =   0   'False
            Locked          =   0   'False
            Text            =   "frm_collect.frx":3764
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
            MultiLine       =   0   'False
            Alignment       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frm_collect.frx":3784
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":37A4
         End
         Begin HexUniControls.ctlUniLabel Label3 
            Height          =   300
            Left            =   120
            Top             =   3720
            Width           =   5115
            _ExtentX        =   9022
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
            Caption         =   "frm_collect.frx":37C0
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":3800
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":3820
         End
         Begin HexUniControls.ctlNumIncXP numInc_nameCounter 
            Height          =   600
            Left            =   3590
            TabIndex        =   45
            Top             =   2040
            Width           =   1695
            _ExtentX        =   2990
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
            Text            =   "999999"
            Min             =   0
            Max             =   999999
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
            MouseIcon       =   "frm_collect.frx":383C
            TrapTabKey      =   0   'False
         End
         Begin HexUniControls.ctlNumIncXP numInc_dateCounter 
            Height          =   600
            Left            =   3590
            TabIndex        =   44
            Top             =   1320
            Width           =   1695
            _ExtentX        =   2990
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
            Text            =   "999999"
            Min             =   1
            Max             =   999999
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
            MouseIcon       =   "frm_collect.frx":3858
            TrapTabKey      =   0   'False
         End
         Begin HexUniControls.ctlUniTextBoxXP txtsampnamebase 
            Height          =   450
            Left            =   120
            TabIndex        =   43
            Top             =   3120
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_collect.frx":3874
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
            Tip             =   "frm_collect.frx":38A0
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":38C0
         End
         Begin HexUniControls.ctlUniLabel Label4 
            Height          =   300
            Left            =   120
            Top             =   2760
            Width           =   2895
            _ExtentX        =   5106
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
            Caption         =   "frm_collect.frx":38DC
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":391C
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":393C
         End
         Begin HexUniControls.ctlUniRadioXP optnamecounter 
            Height          =   450
            Left            =   120
            TabIndex        =   42
            Top             =   2040
            Width           =   3375
            _ExtentX        =   5953
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
            Caption         =   "frm_collect.frx":3958
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":39AA
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":39CA
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP opt_namedate 
            Height          =   450
            Left            =   120
            TabIndex        =   41
            Top             =   1320
            Width           =   3375
            _ExtentX        =   5953
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
            Caption         =   "frm_collect.frx":39E6
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":3A2E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":3A4E
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP optnamemanual 
            Height          =   450
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   3375
            _ExtentX        =   5953
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
            Caption         =   "frm_collect.frx":3A6A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":3ABE
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":3ADE
            ShowFocus       =   -1  'True
         End
      End
   End
   Begin HexUniControls.ctlUniFrameXP frame_main 
      Height          =   7005
      Index           =   3
      Left            =   480
      Top             =   795
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
      Caption         =   "frm_collect.frx":3AFA
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   -2147483642
      Tip             =   "frm_collect.frx":3B1A
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_collect.frx":3B3A
      Begin HexUniControls.ctlUniFrameXP Frame18 
         Height          =   915
         Left            =   5520
         Top             =   3240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1614
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":3B56
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":3BA4
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":3BC4
         Begin HexUniControls.ctlUniRadioXP opt_DynRptNo 
            Height          =   450
            Left            =   1800
            TabIndex        =   66
            Top             =   360
            Width           =   1395
            _ExtentX        =   2461
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
            Caption         =   "frm_collect.frx":3BE0
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":3C04
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":3C24
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP opt_dynRptYes 
            Height          =   450
            Left            =   120
            TabIndex        =   65
            Top             =   360
            Width           =   1395
            _ExtentX        =   2461
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
            Caption         =   "frm_collect.frx":3C40
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":3C66
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":3C86
            ShowFocus       =   -1  'True
         End
      End
      Begin HexUniControls.ctlUniFrameXP Frame6 
         Height          =   1965
         Left            =   5520
         Top             =   4440
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3466
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":3CA2
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":3CE0
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":3D00
         Begin HexUniControls.ctlUniRadioXP opt_tktdemand 
            Height          =   450
            Left            =   120
            TabIndex        =   64
            Top             =   1320
            Width           =   3000
            _ExtentX        =   5292
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
            Caption         =   "frm_collect.frx":3D1C
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":3D4E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":3D6E
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP opt_tktall 
            Height          =   450
            Left            =   120
            TabIndex        =   63
            Top             =   840
            Width           =   3000
            _ExtentX        =   5292
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
            Caption         =   "frm_collect.frx":3D8A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":3DB6
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":3DD6
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP opt_tktno 
            Height          =   450
            Left            =   120
            TabIndex        =   62
            Top             =   360
            Width           =   3000
            _ExtentX        =   5292
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
            Caption         =   "frm_collect.frx":3DF2
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":3E16
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":3E36
            ShowFocus       =   -1  'True
         End
      End
      Begin HexUniControls.ctlUniFrameXP Frame11 
         Height          =   1485
         Left            =   5520
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   2619
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":3E52
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":3E9E
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":3EBE
         Begin HexUniControls.ctlUniRadioXP opt_boundh 
            Height          =   450
            Left            =   2400
            TabIndex        =   60
            Top             =   360
            Width           =   2115
            _ExtentX        =   3731
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
            Caption         =   "frm_collect.frx":3EDA
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":3F02
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":3F22
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP opt_boundl 
            Height          =   450
            Left            =   120
            TabIndex        =   59
            Top             =   840
            Width           =   2115
            _ExtentX        =   3731
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
            Caption         =   "frm_collect.frx":3F3E
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":3F64
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":3F84
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP opt_boundhl 
            Height          =   450
            Left            =   2400
            TabIndex        =   61
            Top             =   840
            Width           =   2115
            _ExtentX        =   3731
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
            Caption         =   "frm_collect.frx":3FA0
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":3FD4
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":3FF4
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP opt_boundno 
            Height          =   450
            Left            =   120
            TabIndex        =   58
            Top             =   360
            Width           =   2115
            _ExtentX        =   3731
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
            Caption         =   "frm_collect.frx":4010
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":4034
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":4054
            ShowFocus       =   -1  'True
         End
      End
      Begin HexUniControls.ctlUniFrameXP Frame14 
         Height          =   1965
         Left            =   240
         Top             =   4440
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3466
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":4070
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":40B6
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":40D6
         Begin HexUniControls.ctlUniTextBoxXP txt_csvfilename 
            Height          =   450
            Left            =   120
            TabIndex        =   56
            Top             =   1200
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_collect.frx":40F2
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
            Tip             =   "frm_collect.frx":4112
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":4132
         End
         Begin HexUniControls.ctlUniLabel Label1 
            Height          =   300
            Left            =   120
            Top             =   840
            Width           =   4575
            _ExtentX        =   8070
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
            Caption         =   "frm_collect.frx":414E
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":41A6
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":41C6
         End
         Begin HexUniControls.ctlUniRadioXP optcsvno 
            Height          =   450
            Left            =   2640
            TabIndex        =   55
            Top             =   360
            Width           =   1395
            _ExtentX        =   2461
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
            Caption         =   "frm_collect.frx":41E2
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":4206
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":4226
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP optcsvyes 
            Height          =   450
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Width           =   1395
            _ExtentX        =   2461
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
            Caption         =   "frm_collect.frx":4242
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":4268
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":4288
            ShowFocus       =   -1  'True
         End
      End
      Begin HexUniControls.ctlUniFrameXP Frame16 
         Height          =   915
         Left            =   5520
         Top             =   2160
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   1614
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":42A4
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":42DA
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":42FA
         Begin HexUniControls.ctlUniCheckXP chk_lims 
            Height          =   450
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   4440
            _ExtentX        =   7832
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
            Caption         =   "frm_collect.frx":4316
            Enabled         =   -1  'True
            Align           =   0
            CheckBackColor  =   -2147483643
            CheckForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":4366
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":4386
            ShowFocus       =   -1  'True
         End
      End
      Begin HexUniControls.ctlUniFrameXP Frame7 
         Height          =   1965
         Left            =   240
         Top             =   2160
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3466
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":43A2
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":43EE
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":440E
         Begin HexUniControls.ctlUniTextBoxXP txtpredfile 
            Height          =   450
            Left            =   120
            TabIndex        =   53
            Top             =   1200
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   794
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frm_collect.frx":442A
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
            Tip             =   "frm_collect.frx":444A
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":446A
         End
         Begin HexUniControls.ctlUniLabel Label5 
            Height          =   300
            Left            =   120
            Top             =   840
            Width           =   2895
            _ExtentX        =   5106
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
            Caption         =   "frm_collect.frx":4486
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frm_collect.frx":44DE
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":44FE
         End
         Begin HexUniControls.ctlUniRadioXP optsavepredsno 
            Height          =   450
            Left            =   2640
            TabIndex        =   52
            Top             =   360
            Width           =   1395
            _ExtentX        =   2461
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
            Caption         =   "frm_collect.frx":451A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":453E
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":455E
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP optsavepredsyes 
            Height          =   450
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   1400
            _ExtentX        =   2461
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
            Caption         =   "frm_collect.frx":457A
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":45A0
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":45C0
            ShowFocus       =   -1  'True
         End
      End
      Begin HexUniControls.ctlUniFrameXP Frame3 
         Height          =   1485
         Left            =   240
         Top             =   360
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   2619
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_collect.frx":45DC
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   16711680
         Tip             =   "frm_collect.frx":461C
         VistaStyle      =   -1  'True
         UseShadow       =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_collect.frx":463C
         Begin HexUniControls.ctlUniRadioXP optpredno 
            Height          =   450
            Left            =   120
            TabIndex        =   50
            Top             =   860
            Width           =   1400
            _ExtentX        =   2461
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
            Caption         =   "frm_collect.frx":4658
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":467C
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":469C
            ShowFocus       =   -1  'True
         End
         Begin HexUniControls.ctlUniRadioXP optpredyes 
            Height          =   450
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   1400
            _ExtentX        =   2461
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
            Caption         =   "frm_collect.frx":46B8
            Enabled         =   -1  'True
            Align           =   0
            RadioBackColor  =   -2147483643
            RadioForeColor  =   -2147483640
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Pressed         =   0   'False
            Tip             =   "frm_collect.frx":46DE
            Style           =   -1
            MousePointer    =   0
            MouseIcon       =   "frm_collect.frx":46FE
            ShowFocus       =   -1  'True
         End
      End
   End
   Begin HexUniControls.ctlUniTabbedXP tab_frame 
      Height          =   320
      Left            =   600
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   495
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
      Tip             =   "frm_collect.frx":471A
      ButtonStyle     =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_collect.frx":473A
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   11280
      Top             =   120
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
   Begin HexUniControls.ctlUniLabel lblprod 
      Height          =   375
      Left            =   2760
      Top             =   120
      Width           =   3615
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
      Caption         =   "frm_collect.frx":4756
      BackColor       =   -2147483633
      ForeColor       =   255
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_collect.frx":4776
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_collect.frx":4796
   End
   Begin HexUniControls.ctlUniLabel Label2 
      Height          =   375
      Left            =   360
      Top             =   120
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
      Caption         =   "frm_collect.frx":47B2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_collect.frx":47F0
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_collect.frx":4810
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   645
      Left            =   8520
      TabIndex        =   0
      Top             =   8040
      Width           =   1995
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
      Caption         =   "frm_collect.frx":482C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_collect.frx":4858
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_collect.frx":4878
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   645
      Left            =   5880
      TabIndex        =   1
      Top             =   8040
      Width           =   1995
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
      Caption         =   "frm_collect.frx":4894
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_collect.frx":48CC
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_collect.frx":48EC
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   11400
      Top             =   720
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
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_collect.frx":4908
   End
End
Attribute VB_Name = "frm_collect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Product settings
Public m_adapterType As String
Public m_adaptIndx As Integer
Public m_alarmMD As Integer
Public m_alarmND As Integer
Public m_alarmRR As Integer
Public m_alarmProp As Integer
Public m_backFreq As REF_FREQS
Public m_bType As String
Public m_clrManualName As Boolean
Public m_clrUserInputs As Boolean
Public m_dayCounter As Long                 ' used in date-counter naming
Public m_delayMeasure As Single             ' delay time between measurements
Public m_delayStart As Single               ' delay time before acquisiiton start
Public m_doLIMS As Integer
Public m_endWavenum As Double               ' ending wavenumber (cm-1)
Public m_endWavenumIndx As Long             ' ending wavenumber index
Public m_extRefFileName As String
Public m_gainIndx As Integer                ' gain index
Public m_hideValCol As Boolean
Public m_makePred As String
Public m_multiCupIndx As Integer
Public m_multiCupType As String
Public m_nameBase As String
Public m_nameCounter As Long
Public m_nameScanType As String
Public m_noOLVal As String
Public m_noPredVal As String
Public m_numMeasures As Long                ' number of measurements
Public m_numModelVars As Integer            ' num of model variables
Public m_numSamples As Long                 ' number of samples
Public m_olFormat As Boolean                ' if true=pics if false = numbers
Public m_olRefFileName As String
Public m_repsAvg As Integer
Public m_resolution As Integer              ' scanning resolution (cm -1)
Public m_resolutionIndx As Integer          ' scanning resolution index
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
Public m_smplEndWvln As Double
Public m_smplNScans As Integer
Public m_smplNumPts As Long                 ' sampling number of points
Public m_smplPPT As Integer
Public m_smplStartWvln As Double
Public m_sNameMode As Integer               ' 0 = no save, 1 = enter, 2 = comment, 3 = counter,
                                            ' 4 = date_counter, 5 = Stat name each, 6 = stat counter
Public m_speedIndx As Integer               ' scan arm speed index
Public m_startWavenum As Double             ' starting wavenumber (cm -1)
Public m_startWavenumIndx As Long           ' starting wavenumber index
Public m_sType As String
Public m_trayNum As Integer
Public m_useExtRefTrayCfg As Boolean
Public m_useMIV As Boolean
Public m_valueBound As Integer              ' bounds reported values to max/min 0=act value 1=bound min 2 = bound max, 3 = both
Public m_waveNumIncr As Double              ' sampling wavenumber intervals (cm -1)
Public m_writeTkt As Integer                ' 0 = no, 1 = always, 2 = on demand

Public m_ignoreEvent As Boolean

Public Sub show_frame(frameOffset As Integer)
  Dim ii As Integer
  
#If ABBFT Then
  ' Note: frame_main(1) not used
  Select Case (frameOffset)
    Case 0
      frame_main(0).Visible = True
      frame_main(1).Visible = False
      frame_main(2).Visible = False
      frame_main(3).Visible = False
      frame_main(4).Visible = False
    Case 1
      frame_main(2).Visible = True
      frame_main(0).Visible = False
      frame_main(1).Visible = False
      frame_main(3).Visible = False
      frame_main(4).Visible = False
    Case 2
      frame_main(3).Visible = True
      frame_main(0).Visible = False
      frame_main(1).Visible = False
      frame_main(2).Visible = False
      frame_main(4).Visible = False
    Case 3
      frame_main(4).Visible = True
      frame_main(0).Visible = False
      frame_main(1).Visible = False
      frame_main(2).Visible = False
      frame_main(3).Visible = False
  End Select
#Else
  For ii = 0 To frame_main.Count - 1
    If ii = frameOffset Then
      frame_main(ii).Visible = True
    Else
      frame_main(ii).Visible = False
    End If
  Next
#End If
End Sub

#If ABBFT Then
Public Sub calc_wavenum_incr(resolution As Integer, numPts As Long, waveNumIncr As Double)
  
  numPts = 32768 / resolution
  waveNumIncr = CDbl(MB3000_MAX_WAVENUM / numPts)
End Sub
#End If

#If ABBFT Then
Public Function calc_wavenum(wavenumIndx As Long, wavenum As Double) As Boolean
  Dim numPts As Long
  
  calc_wavenum = True
  numPts = txt_smplNumPts.Text
  
  If (wavenumIndx < 0) Then
    wavenumIndx = 0
    calc_wavenum = False
  Else
    If (wavenumIndx >= numPts) Then
      wavenumIndx = numPts - 1
      calc_wavenum = False
    End If
  End If
  
  wavenum = wavenumIndx * CDbl(txt_waveNumIncr.Text)
End Function
#End If

Public Function check_scan_settings() As Boolean
  Dim startWvIndx As Integer
  Dim endWvIndx As Integer
  Dim startWvln As Double
  Dim endWvln As Double
  Dim rc As Boolean
  Dim userReq As Integer

  check_scan_settings = False

#If ABBFT Then
  If (lst_resolution.ListIndex = -1) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_collect", "errMsg8", "Please select a Resolution in Scan Settings Tab"), vbExclamation
    Exit Function
  End If
  
  If (lst_speed.ListIndex = -1) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_collect", "errMsg9", "Please select an Arm Speed in Scan Settings Tab"), vbExclamation
    Exit Function
  End If
  
  If (lst_gain.ListIndex = -1) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_collect", "errMsg10", "Please select a Gain Value in Scan Settings Tab"), vbExclamation
    Exit Function
  End If
  
  If (IsNumeric(txt_numMeasures.Text) = False) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_collect", "errMsg11", "Please enter a valid # of Measurements in Scan Settings Tab"), vbExclamation
    Exit Function
  End If
  
  If (IsNumeric(txt_delayStart.Text) = False) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_collect", "errMsg12", "Please enter a valid Scan Start Delay in Scan Settings Tab"), vbExclamation
    Exit Function
  End If
  
  If (IsNumeric(txt_delayMeasure.Text) = False) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_collect", "errMsg13", "Please enter a valid Measurement Delay in Scan Settings Tab"), vbExclamation
    Exit Function
  End If
  
  startWvIndx = CInt(txt_startWavenumIndx.Text)
  endWvIndx = CInt(txt_endWavenumIndx.Text)
  
  If (startWvIndx >= endWvIndx) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_collect", "errMsg14", "Please enter valid Starting/Ending Wavenumber Indices in Scan Settings Tab"), vbExclamation
    Exit Function
  End If
#Else
  If (optbgexternal.Value = True) Then
    If (combo_extRefFileName.Text = "") Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_collect", "errMsg5", "Please select a External Reference file name in Scan Settings Tab"), vbExclamation
      Exit Function
    End If
  Else
    If (optbgfile.Value = True) Then
      If (combo_olRefFileName.Text = "") Then
        CWrap.ShowMessageBoxW MLSupport.GSS("frm_collect", "errMsg6", "Please select a Offline Reference file name in Scan Settings Tab"), vbExclamation
        Exit Function
      End If
    End If
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
          CWrap.ShowMessageBoxW MLSupport.GSS("frm_collect", "errMsg1", "Please enter a Scan Ending Wavelength value in Scan Settings Tab greater than the Starting Wavelength value"), vbExclamation
          Exit Function
        Else
          ' Check wavelengths with instrument's table
#If SSRCS Then
          SSRCSClientError = unity_main.SSRCSClient.chkWvlnRange(startWvln, endWvln, rc)
#Else
          rc = unity_main.MS11srv.chkWvlnRange(startWvln, endWvln)
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
  If (frm_collect.m_adapterType = CFG_MULTI_CUP_AT) And (frm_collect.m_multiCupType = CFG_NONE_MCT) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_collect", "errMsg2", "Please select a Multi-Cup Type for Multi-Cup Adapter in Tray Settings Tab"), vbExclamation
    Exit Function
  End If
  
  If (frm_collect.m_rotateMoveMode <> TRM_NONE) And (frm_collect.m_rotateDir = TRD_NONE) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_collect", "errMsg3", "Please select a Direction for Rotating Platter in Tray Settings Tab"), vbExclamation
    Exit Function
  End If
  
  If (frm_collect.optsavescanyes.Value = True) Then
    If (Trim(frm_collect.txt_caldir.Text) = "") Then
      CWrap.ShowMessageBoxW MLSupport.GSS("frm_applCfg", "errMsg7", "Please enter a Spectrum Directory in Sample Naming Tab"), vbExclamation
      Exit Function
    End If
  End If
#End If

  check_scan_settings = True
  Exit Function
  
BAD_VALUE:
  CWrap.ShowMessageBoxW MLSupport.GSS("frm_collect", "errMsg4", "Please enter a valid number for Scan Starting/Ending Wavelength"), vbOKOnly
End Function

Public Function check_ref_file_wvlns(refStartWvln As Double, refEndWvln As Double, refFileName As String) As Boolean

  check_ref_file_wvlns = False
  
  ' Check reference wavelengths against instrument
  If (refStartWvln < unity_main.m_minWvln) Or (refStartWvln > unity_main.m_maxWvln) Then
    unity_main.m_ansiErrMsg = ("Reference file " & refFileName & " starting wavelength " & refStartWvln & " outside of instrument wavelength range " & unity_main.m_minWvln & " - " & unity_main.m_maxWvln)
    unity_main.m_uniErrMsg = MLSupport.GGS_Params("frm_collect.errMsg5", "Reference file %1 starting wavelength %2 outside of instrument wavelength range %3 - %4", refFileName, CStr(refStartWvln), CStr(unity_main.m_minWvln), CStr(unity_main.m_maxWvln))
    GoTo BAD_WVLNS
  Else
    If (refEndWvln < unity_main.m_minWvln) Or (refEndWvln > unity_main.m_maxWvln) Then
      unity_main.m_ansiErrMsg = ("Reference file " & refFileName & " ending wavelength " & refEndWvln & " outside of instrument wavelength range " & unity_main.m_minWvln & " - " & unity_main.m_maxWvln)
      unity_main.m_uniErrMsg = MLSupport.GGS_Params("frm_collect.errMsg6", "Reference file %1 ending wavelength %2 outside of instrument wavelength range %3 - %4", refFileName, CStr(refEndWvln), CStr(unity_main.m_minWvln), CStr(unity_main.m_maxWvln))
      GoTo BAD_WVLNS
    End If
  End If
  
  check_ref_file_wvlns = True
  Exit Function
  
BAD_WVLNS:
  unity_main.errorstring = unity_main.m_ansiErrMsg
  unity_main.write_error
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", unity_main.m_uniErrMsg), vbCritical
End Function

#If SSTAR Then
Public Sub update_prod_ext_ref_wvlns(fileName As String)

  ' Load external reference cfg file
  If (frm_extRef.load_ext_ref_cfg_file(fileName, False) = True) Then
    ' Check reference wavelengths against instrument
    If (check_ref_file_wvlns(frm_extRef.m_extRefStartWvln, frm_extRef.m_extRefEndWvln, frm_extRef.m_extRefFileName) = False) Then
      GoTo REF_FILE_ERR
    End If
    
    frm_collect.chk_useExtRefTrayCfg.Value = 0
    frm_collect.chk_useExtRefTrayCfg.Value = 1
    
    ' Setup product starting/ending wavelengths to match external reference
    frm_collect.txt_endWvln.Text = frm_extRef.m_extRefEndWvln
    frm_collect.txt_startWvln.Text = frm_extRef.m_extRefStartWvln

    ' Check if wavelengths match product defaults
    If (ProdDfltData.endWvln = frm_extRef.m_extRefEndWvln) And (ProdDfltData.startWvln = frm_extRef.m_extRefStartWvln) Then
      frm_collect.chk_dfltWavelens.Value = 1
    Else
      frm_collect.chk_dfltWavelens.Value = 0
    End If
  Else
REF_FILE_ERR:
    frm_collect.combo_extRefFileName.Text = ""
  End If
End Sub
#End If

#If SSTAR Then
Public Sub update_prod_ol_ref_wvlns(fileName As String)

  If (frm_olRef.get_ol_ref_wvlns(fileName) = True) Then
    ' Check reference wavelengths against instrument
    If (check_ref_file_wvlns(frm_olRef.m_olRefStartWvln, frm_olRef.m_olRefEndWvln, frm_olRef.m_olRefFileName) = False) Then GoTo REF_FILE_ERR
    
    ' Setup product starting/ending wavelengths to match external reference
    frm_collect.txt_endWvln.Text = frm_olRef.m_olRefEndWvln
    frm_collect.txt_startWvln.Text = frm_olRef.m_olRefStartWvln

    ' Check if wavelengths match product defaults
    If (ProdDfltData.endWvln = frm_olRef.m_olRefEndWvln) And (ProdDfltData.startWvln = frm_olRef.m_olRefStartWvln) Then
      frm_collect.chk_dfltWavelens.Value = 1
    Else
      frm_collect.chk_dfltWavelens.Value = 0
    End If
  Else
REF_FILE_ERR:
    frm_collect.combo_olRefFileName.Text = ""
  End If
End Sub
#End If

#If SSTAR Then
Public Sub setup_prod_tray_cfg()
  Dim fileName As String

  m_ignoreEvent = True
  fileName = (PRODUCTS_CFG_DIR & unity_main.current_ini)

  ' Setup adapter type and get tray number
  setup_adapter_tray_num
  
  ' Setup tray rotation direction
  Select Case (frm_collect.m_rotateDir)
    Case TRD_CW         ' Clockwise
      opt_rotateDirCW.Value = True
    Case TRD_NONE       ' no rotation
      opt_rotateDirCW.Value = False
      opt_rotateDirCCW.Value = False
    Case TRD_CCW        ' Counter Clockwise
      opt_rotateDirCCW.Value = True
    Case Else           ' invalid value
      unity_main.errorstring = (fileName & " had incompatible value. RotateDir was " & frm_collect.m_rotateDir & "; updated to " & ProdDfltData.rotateDir)
      unity_main.write_error
      frm_collect.m_rotateDir = ProdDfltData.rotateDir
      frm_collect.m_rotateMoveMode = ProdDfltData.rotateMoveMode
      unity_main.m_badIniVal = True
  End Select
  
  ' Setup tray movement
  Select Case (frm_collect.m_rotateMoveMode)
    Case TRM_NONE                ' no movement
    Case TRM_CONT                ' continuous movement
      ' Check for invalid velocity value
      If (frm_collect.m_rotateSpeed < MS11TrayInfoData(frm_collect.m_trayNum).minVel) Or (frm_collect.m_rotateSpeed > MS11TrayInfoData(frm_collect.m_trayNum).maxVel) Then
        unity_main.errorstring = (fileName & " had incompatible value. RotateSpeed was " & frm_collect.m_rotateSpeed & "; updated to " & MS11DfltTrayCfgData(frm_collect.m_trayNum).velCont)
        unity_main.write_error
        frm_collect.m_rotateSpeed = MS11DfltTrayCfgData(frm_collect.m_trayNum).velCont
        unity_main.m_badIniVal = True
      End If
    Case TRM_STEP                ' stepped movement
      ' Check for invalid number of steps for stepped samples
      If ((frm_collect.m_rotateStepSteps < 1) And (MS11TrayInfoData(frm_collect.m_trayNum).maxStps4scn <> 0)) Or (frm_collect.m_rotateStepSteps > MS11TrayInfoData(frm_collect.m_trayNum).maxStps4scn) Then
        unity_main.errorstring = (fileName & " had incompatible value. RotateStepSteps was " & frm_collect.m_rotateStepSteps & "; updated to " & MS11DfltTrayCfgData(frm_collect.m_trayNum).stps4scn)
        unity_main.write_error
        frm_collect.m_rotateStepSteps = MS11DfltTrayCfgData(frm_collect.m_trayNum).stps4scn
        unity_main.m_badIniVal = True
      End If
    Case TRM_INDEX               ' indexed movement
      ' Check for invalid number of steps for indexed samples
      If ((frm_collect.m_rotateIndexSteps < 1) And (MS11TrayInfoData(frm_collect.m_trayNum).maxStps4IX <> 0)) Or (frm_collect.m_rotateIndexSteps > MS11TrayInfoData(frm_collect.m_trayNum).maxStps4IX) Then
        unity_main.errorstring = (fileName & " had incompatible value. RotateIndexSteps was " & frm_collect.m_rotateIndexSteps & "; updated to " & MS11DfltTrayCfgData(frm_collect.m_trayNum).stps4IX)
        unity_main.write_error
        frm_collect.m_rotateIndexSteps = MS11DfltTrayCfgData(frm_collect.m_trayNum).stps4IX
        unity_main.m_badIniVal = True
      End If
    Case Else                       ' invalid value
      unity_main.errorstring = (fileName & " had incompatible value. RotateMoveMode was " & frm_collect.m_rotateMoveMode & "; updated to " & ProdDfltData.rotateMoveMode)
      unity_main.write_error
      frm_collect.m_rotateMoveMode = ProdDfltData.rotateMoveMode
      frm_collect.m_rotateDir = ProdDfltData.rotateDir
      unity_main.m_badIniVal = True
  End Select
  
  numInc_rotateIndexSteps.Text = frm_collect.m_rotateIndexSteps
  numInc_rotateSpeed.Text = frm_collect.m_rotateSpeed
  numInc_rotateStepSteps.Text = frm_collect.m_rotateStepSteps

  m_ignoreEvent = False
End Sub
#End If

#If SSTAR Then
Public Sub setup_adapter_tray_num()
  Dim adaptIndx As Integer
  Dim fileName As String

  fileName = (PRODUCTS_CFG_DIR & unity_main.current_ini)
  
  ' Get tray number for instrument's adapter type
PROCESS_ADAPTERTYPE:
  For adaptIndx = 1 To MAX_ADAPTER_TYPES
    If (frm_collect.m_adapterType = MS11AdapterInfo(adaptIndx).cfgName) Then
      frm_collect.m_adaptIndx = adaptIndx
      frm_collect.m_trayNum = MS11AdapterInfo(adaptIndx).trayNum
      GoTo LEAVE_LOOP
    End If
  Next adaptIndx

  ' Flag invalid adapter type
  unity_main.errorstring = (fileName & " had incompatible value. AdapterType was " & frm_collect.m_adapterType & "; updated to " & ProdDfltData.adapterType)
  unity_main.write_error
  frm_collect.m_adapterType = ProdDfltData.adapterType
  unity_main.m_badIniVal = True
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
      Select Case (frm_collect.m_adapterType)
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
#End If

#If SSTAR Then
Public Sub disable_multicup_types()
  Dim nn As Integer
  
  frame_multiCupType.Visible = False
  frm_collect.m_multiCupType = CFG_NONE_MCT
  
  For nn = 1 To MAX_MULTI_CUP_TYPES
    opt_multiCupType(nn).Value = False
  Next nn
End Sub
#End If

#If SSTAR Then
Public Sub setup_multi_cup_type()
  Dim multiCupIndx As Integer
  Dim fileName As String

  fileName = (PRODUCTS_CFG_DIR & unity_main.current_ini)

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
    If (frm_collect.m_multiCupType = MS11MultiCupInfo(multiCupIndx).cfgName) Then
      frm_collect.m_multiCupIndx = multiCupIndx
      frm_collect.m_trayNum = MS11MultiCupInfo(multiCupIndx).trayNum
      If (multiCupIndx <> 0) Then
        opt_multiCupType(multiCupIndx).Value = True
      End If
      GoTo LEAVE_LOOP
    End If
  Next multiCupIndx
  
  ' Flag invalid multi-cup type
  unity_main.errorstring = (fileName & " had incompatible value. MultiCupType was " & frm_collect.m_multiCupType & "; updated to " & ProdDfltData.multiCupType)
  unity_main.write_error
  frm_collect.m_multiCupType = ProdDfltData.multiCupType
  unity_main.m_badIniVal = True
  GoTo PROCESS_MULTICUPTYPE
 
LEAVE_LOOP:
  setup_rotate_platter
End Sub
#End If

#If SSTAR Then
Public Sub setup_rotate_platter()
  Dim indx As Integer
  
  ' Setup plater info only for valid tray number
  If (frm_collect.m_trayNum > 0) And (frm_collect.m_trayNum <= MS11CfgData.nTrays) Then
    frame_rotatePlatter.Visible = True
    indx = frm_collect.m_trayNum
  
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
#End If
  
#If SSTAR Then
Public Sub setup_rotate_move_mode()
  Dim fileName As String

  fileName = (PRODUCTS_CFG_DIR & unity_main.current_ini)
  
  Select Case (frm_collect.m_rotateMoveMode)
    Case TRM_NONE        ' None
      frm_collect.m_rotateDir = TRD_NONE
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
        If (frm_collect.m_rotateSpeed = 0) Then
          numInc_rotateSpeed.Text = MS11DfltTrayCfgData(frm_collect.m_trayNum).velCont
        End If
      Else
        unity_main.errorstring = (fileName & " had incompatible value. RotateMoveMode was " & frm_collect.m_rotateMoveMode & "; updated to " & ProdDfltData.rotateMoveMode)
        unity_main.write_error
        frm_collect.m_rotateMoveMode = ProdDfltData.rotateMoveMode
        frm_collect.m_rotateDir = ProdDfltData.rotateDir
        unity_main.m_badIniVal = True
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
        If (frm_collect.m_rotateStepSteps = 0) Then
          numInc_rotateStepSteps.Text = MS11DfltTrayCfgData(frm_collect.m_trayNum).stps4scn
        End If
      Else
        unity_main.errorstring = (fileName & " had incompatible value. RotateMoveMode was " & frm_collect.m_rotateMoveMode & "; updated to " & ProdDfltData.rotateMoveMode)
        unity_main.write_error
        frm_collect.m_rotateMoveMode = ProdDfltData.rotateMoveMode
        frm_collect.m_rotateDir = ProdDfltData.rotateDir
        unity_main.m_badIniVal = True
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
        If (frm_collect.m_rotateIndexSteps = 0) Then
          numInc_rotateIndexSteps.Text = MS11DfltTrayCfgData(frm_collect.m_trayNum).stps4IX
        End If
      Else
        unity_main.errorstring = (fileName & " had incompatible value. RotateMoveMode was " & frm_collect.m_rotateMoveMode & "; updated to " & ProdDfltData.rotateMoveMode)
        unity_main.write_error
        frm_collect.m_rotateMoveMode = ProdDfltData.rotateMoveMode
        frm_collect.m_rotateDir = ProdDfltData.rotateDir
        unity_main.m_badIniVal = True
        setup_rotate_move_mode
      End If
      
    Case Else               ' Unknown
      unity_main.errorstring = (fileName & " had incompatible value. RotateMoveMode was " & frm_collect.m_rotateMoveMode & "; updated to " & ProdDfltData.rotateMoveMode)
      unity_main.write_error
      frm_collect.m_rotateMoveMode = ProdDfltData.rotateMoveMode
      frm_collect.m_rotateDir = ProdDfltData.rotateDir
      unity_main.m_badIniVal = True
      setup_rotate_move_mode
  End Select
End Sub
#End If

Public Function get_samp_name() As Boolean
  Dim fileName As String
  Dim dirname As String
  Dim fileFound As Boolean
  Dim ctr As Long
  Dim ii As Integer
  Dim scanMsg As String
  Dim uniMsg As String
  
  unity_main.gotscanname = False

  ' Check if using global sample naming convention
  If (unity_main.m_enableGlobalName = True) Then
    get_global_samp_name
  End If

  Select Case (unity_main.m_sNameMode)
    Case 1                  ' enter manual name
      ' Check if no batch scan being performed
      If (unity_main.m_batchRunFlg = True) Then
        unity_main.gotscanname = True
      Else
        frm_scanname.Show 1
  
        'Check if using global sample naming convention
        If (unity_main.m_enableGlobalName = True) Then
          If (unity_main.gotscanname = True) Then
            ' Save entry if using base counter only and valid number
            If (unity_main.m_globalNameMode = "Ctr") And (IsNumeric(unity_main.txtsamplename.Text) = True) Then
              ctr = unity_main.txtsamplename.Text
              ctr = ctr + 1
        
              If (ctr > frm_Inst.numInc_globalBaseCtr.Max) Then
                ctr = frm_Inst.numInc_globalBaseCtr.Min
              End If
        
              frm_Inst.numInc_globalBaseCtr.Text = ctr
            End If
          Else
            rollback_global_ctr
          End If
        End If
      End If
      
    Case 3                  ' auto name based on base + counter
      ' Check if not using global sample naming convention
      If (unity_main.m_enableGlobalName = False) Then
        If (unity_main.m_saveIt = "save") Then
          If (unity_main.firstrep = True) Then
hadit:
            fileName = unity_main.m_nameBase & frm_collect.numInc_nameCounter.Text & SPC_FILE_EXT
            dirname = unity_main.m_saveDir
            fileFound = CFile.st_FileExist(dirname & fileName)
        
            If (fileFound = True) Then
              ctr = frm_collect.numInc_nameCounter.Text + 1
        
              If (ctr > frm_collect.numInc_nameCounter.Max) Then
                ctr = frm_collect.numInc_nameCounter.Min
              End If
        
              frm_collect.numInc_nameCounter.Text = ctr
              GoTo hadit
            End If
        
            unity_main.txtsamplename.Text = unity_main.m_nameBase & frm_collect.numInc_nameCounter.Text
          End If
        Else
          ctr = frm_collect.numInc_nameCounter.Text + 1
        
          If (ctr > frm_collect.numInc_nameCounter.Max) Then
            ctr = frm_collect.numInc_nameCounter.Min
          End If
        
          frm_collect.numInc_nameCounter.Text = ctr
          unity_main.txtsamplename.Text = unity_main.m_nameBase & ctr
        End If
      End If

      ' Check if to use manual data entry and no batch scan
      If (unity_main.m_useMIV = True) And (unity_main.m_batchRunFlg = False) Then
        frm_scanname.txt_fname.Text = unity_main.txtsamplename.Text
        frm_scanname.Show 1
          
        ' Roll back counter if sample canceled
        If (unity_main.gotscanname = False) Then
          ' Check if using global sample naming convention
          If (unity_main.m_enableGlobalName = True) Then
            rollback_global_ctr
          Else
            If (unity_main.m_saveIt = "nosave") Then
              ctr = frm_collect.numInc_nameCounter.Text - 1
        
              If (ctr < frm_collect.numInc_nameCounter.Min) Then
                ctr = frm_collect.numInc_nameCounter.Max
              End If
        
              frm_collect.numInc_nameCounter.Text = ctr
            End If
          End If
        End If
      Else
        unity_main.gotscanname = True
      End If

    Case 4                  ' auto name based on date + counter
      ' Check if not using global sample naming convention
      If (unity_main.m_enableGlobalName = False) Then
        frm_collect.makedatename
      End If
      
      ' Check if to use manual data entry and no batch scan
      If (unity_main.m_useMIV = True) And (unity_main.m_batchRunFlg = False) Then
        frm_scanname.txt_fname.Text = unity_main.txtsamplename.Text
        frm_scanname.Show 1
          
        ' Roll back counter if sample canceled
        If (unity_main.gotscanname = False) Then
          ' Check if using global sample naming convention
          If (unity_main.m_enableGlobalName = True) Then
            rollback_global_ctr
          Else
            If (unity_main.m_saveIt = "nosave") Then
              ctr = frm_collect.numInc_dateCounter.Text - 1
        
              If (ctr < frm_collect.numInc_dateCounter.Min) Then
                ctr = frm_collect.numInc_dateCounter.Max
              End If
        
              frm_collect.numInc_dateCounter.Text = ctr
            End If
          End If
        End If
      Else
        unity_main.gotscanname = True
      End If
    
    Case Else
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_collect.errMsg4", "Unsupported Sample Naming Mode %1", CStr(unity_main.m_sNameMode)), vbCritical
  End Select
  
  get_samp_name = unity_main.gotscanname
  
  If (get_samp_name = True) Then
    scanMsg = ("Sample name = " & unity_main.txtsamplename.Text)
    uniMsg = MLSupport.GGS_Params("frm_collect.scanMsg1", "Sample name: %1", unity_main.txtsamplename.Text)
    Call unity_main.log_scan_status(LOG_DBG_LEVEL2, scanMsg, uniMsg)
    
    If (unity_main.m_saveIt = "save") Then
      scanMsg = "Sample spectrum to be saved to file"
      uniMsg = MLSupport.GSS("frm_collect", "statMsg1", "Sample spectrum to be saved to file")
    Else
      scanMsg = "Sample spectrum will not be saved to file"
      uniMsg = MLSupport.GSS("frm_collect", "statMsg2", "Sample spectrum will not be saved to file")
    End If
    
    Call unity_main.log_scan_status(LOG_DBG_LEVEL2, scanMsg, uniMsg)
  End If
End Function

Public Sub get_global_samp_name()
  Dim fileName As String
  Dim dirname As String
  Dim fileFound As Boolean
  Dim ctr As Long
  Dim datex As String
  Dim buildName As String
  
  Select Case (unity_main.m_globalNameMode)
    Case "Base"           ' base name + counter
      frm_Inst.m_saveGlobalIni = True     ' flag to save ini file

      If (unity_main.m_saveIt = "save") Then
        If (unity_main.firstrep = True) Then
INC_BASE_CTR:
          ctr = frm_Inst.numInc_globalBaseCtr.Text
          fileName = unity_main.m_globalNameBase & ctr & SPC_FILE_EXT
          dirname = unity_main.m_saveDir
          fileFound = CFile.st_FileExist(dirname & fileName)
        
          ' Check if file exists
          If (fileFound = True) Then
            ctr = ctr + 1
        
            If (ctr > frm_Inst.numInc_globalBaseCtr.Max) Then
              ctr = frm_Inst.numInc_globalBaseCtr.Min
            End If
        
            frm_Inst.numInc_globalBaseCtr.Text = ctr
            GoTo INC_BASE_CTR
          End If
      
          unity_main.txtsamplename.Text = unity_main.m_globalNameBase & frm_Inst.numInc_globalBaseCtr.Text
          
          ' Setup for next sample
          ctr = ctr + 1
      
          If (ctr > frm_Inst.numInc_globalBaseCtr.Max) Then
            frm_Inst.numInc_globalBaseCtr.Text = frm_Inst.numInc_globalBaseCtr.Min
          End If
          
          frm_Inst.numInc_globalBaseCtr.Text = ctr
        End If
      Else
        ctr = frm_Inst.numInc_globalBaseCtr.Text
      
        ' Setup for next sample
        If (ctr + 1 > frm_Inst.numInc_globalBaseCtr.Max) Then
          frm_Inst.numInc_globalBaseCtr.Text = frm_Inst.numInc_globalBaseCtr.Min
        Else
          frm_Inst.numInc_globalBaseCtr.Text = ctr + 1
        End If
        
        unity_main.txtsamplename.Text = unity_main.m_globalNameBase & ctr
      End If
      
    Case "Ctr"           ' base name counter only
      frm_Inst.m_saveGlobalIni = True     ' flag to save ini file

      If (unity_main.m_saveIt = "save") Then
        If (unity_main.firstrep = True) Then
INC_BASE_CTR2:
          ctr = frm_Inst.numInc_globalBaseCtr.Text
          fileName = ctr & SPC_FILE_EXT
          dirname = unity_main.m_saveDir
          fileFound = CFile.st_FileExist(dirname & fileName)
        
          ' Check if file exists
          If (fileFound = True) Then
            ctr = ctr + 1
        
            If (ctr > frm_Inst.numInc_globalBaseCtr.Max) Then
              ctr = frm_Inst.numInc_globalBaseCtr.Min
            End If
        
            frm_Inst.numInc_globalBaseCtr.Text = ctr
            GoTo INC_BASE_CTR2
          End If
      
          unity_main.txtsamplename.Text = frm_Inst.numInc_globalBaseCtr.Text
          
          ' Setup for next sample
          ctr = ctr + 1
      
          If (ctr > frm_Inst.numInc_globalBaseCtr.Max) Then
            frm_Inst.numInc_globalBaseCtr.Text = frm_Inst.numInc_globalBaseCtr.Min
          End If
          
          frm_Inst.numInc_globalBaseCtr.Text = ctr
        End If
      Else
        ctr = frm_Inst.numInc_globalBaseCtr.Text
      
        ' Setup for next sample
        If (ctr + 1 > frm_Inst.numInc_globalBaseCtr.Max) Then
          frm_Inst.numInc_globalBaseCtr.Text = frm_Inst.numInc_globalBaseCtr.Min
        Else
          frm_Inst.numInc_globalBaseCtr.Text = ctr + 1
        End If
        
        unity_main.txtsamplename.Text = ctr
      End If
      
    Case "Date"           ' date name + counter
      frm_Inst.m_saveGlobalIni = True     ' flag to save ini file
      datex = Date
  
      ' Reset date counter if new day
      If (unity_main.m_globalDate <> datex) Then
        unity_main.m_globalDate = datex
        frm_Inst.numInc_globalDateCtr.Text = frm_Inst.numInc_globalDateCtr.Min
      End If
  
INC_DATE_CTR:
      Call rebuild_date(unity_main.m_globalDate, buildName)
      buildName = buildName & "_"

      ' Check if spectrum file is to be saved
      If (unity_main.m_saveIt = "save") Then
        ctr = frm_Inst.numInc_globalDateCtr.Text
        fileName = buildName & ctr & SPC_FILE_EXT
        dirname = unity_main.m_saveDir
        fileFound = CFile.st_FileExist(dirname & fileName)
        
        ' Check if file exists
        If (fileFound = True) Then
          ctr = ctr + 1
        
          If (ctr > frm_Inst.numInc_globalDateCtr.Max) Then
            ctr = frm_Inst.numInc_globalDateCtr.Min
          End If
        
          frm_Inst.numInc_globalDateCtr.Text = ctr
          GoTo INC_DATE_CTR
        End If
        
        unity_main.txtsamplename.Text = buildName & frm_Inst.numInc_globalDateCtr.Text
        
        ' Setup for next sample
        ctr = ctr + 1
        
        If (ctr > frm_Inst.numInc_globalDateCtr.Max) Then
          ctr = frm_Inst.numInc_globalDateCtr.Min
        End If
        
        frm_Inst.numInc_globalDateCtr.Text = ctr
      Else
        ctr = frm_Inst.numInc_globalDateCtr.Text
      
        ' Setup for next sample
        If (ctr + 1 > frm_Inst.numInc_globalDateCtr.Max) Then
          frm_Inst.numInc_globalDateCtr.Text = frm_Inst.numInc_globalDateCtr.Min
        Else
          frm_Inst.numInc_globalDateCtr.Text = ctr + 1
        End If
        
        unity_main.txtsamplename.Text = buildName & ctr
      End If
    End Select
End Sub

Public Sub roll_back_name_ctr()
  Dim ctr As Long
  
  ' Check if using global sample naming convention
  If (unity_main.m_enableGlobalName = True) Then
    frm_collect.rollback_global_ctr
  Else
    If (LCase(unity_main.m_saveIt) = "nosave") Then
      ' Roll back counter since sample canceled
      Select Case (unity_main.m_sNameMode)
        Case 3                  ' auto name based on base + counter
          ctr = frm_collect.numInc_nameCounter.Text - 1
        
          If (ctr < frm_collect.numInc_nameCounter.Min) Then
            ctr = frm_collect.numInc_nameCounter.Max
          End If
        
          frm_collect.numInc_nameCounter.Text = ctr

        Case 4                  ' auto name based on date + counter
          ctr = frm_collect.numInc_dateCounter.Text - 1
        
          If (ctr < frm_collect.numInc_dateCounter.Min) Then
            ctr = frm_collect.numInc_dateCounter.Max
          End If
        
          frm_collect.numInc_dateCounter.Text = ctr
      End Select
    End If
  End If
End Sub

Public Sub rollback_global_ctr()
  Dim ctr As Long
  
  ' Check if base name + counter or counter only
  If (unity_main.m_globalNameMode = "Base") Or (unity_main.m_globalNameMode = "Ctr") Then
    ctr = frm_Inst.numInc_globalBaseCtr.Text - 1
        
    If (ctr < frm_Inst.numInc_globalBaseCtr.Min) Then
      ctr = frm_Inst.numInc_globalBaseCtr.Max
    End If
      
    frm_Inst.numInc_globalBaseCtr.Text = ctr
  Else      ' date name + counter
    ctr = frm_Inst.numInc_globalDateCtr.Text - 1
        
    If (ctr < frm_Inst.numInc_globalDateCtr.Min) Then
      ctr = frm_Inst.numInc_globalDateCtr.Max
    End If
        
    frm_Inst.numInc_globalDateCtr.Text = ctr
  End If
End Sub

Sub checkalarms()
  Dim cvalue As Double
  Dim cond1, cond2, cond3, cond4 As Double
  Dim picType As String ' 1=green, 2=yellow, 3=red
  Dim ii As Integer
  Dim rowHt As Integer
  Dim rratio As Double
  Dim res1, res2 As Double
  Dim pFExpCodes As Long
  Dim ndPFExpCode As Integer
  Dim tempstring As String
  Dim scanMsg As String
  Dim uniMsg As String
  Dim labData() As Single
  Dim nn As Integer
  Dim expCodeFlg As Boolean

  scanMsg = "Checking property alarm limits"
  uniMsg = MLSupport.GSS("frm_collect", "statMsg3", "Checking property alarm limits")
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, scanMsg, uniMsg)
  
  unity_main.fpspread_pred.Row = 1
  unity_main.fpspread_pred.Col = 3
  unity_main.fpspread_pred.Row2 = MAX_NUM_PROPS
  unity_main.fpspread_pred.Col2 = unity_main.fpspread_pred.MaxCols
  unity_main.fpspread_pred.BlockMode = True
  unity_main.fpspread_pred.CellType = 9
  unity_main.fpspread_pred.BlockMode = False

  ' Check & show mdist if configured
  If (frmedmod.chk_md.Value = 1) Then
    For ii = 0 To frmedmod.numprops - 1
      tempstring = Trim(unity_main.lst_modtype.List(ii))
      
      ' Only PLS, CalStar & PRD models support M-Dist
      If (tempstring = "2") Or (tempstring = "3") Then
        GoTo NO_MDIST
      End If
      
      frmedmod.grid_models.Row = ii + 1
      frmedmod.grid_models.Col = 7
      cond1 = frmedmod.grid_models.Value
      frmedmod.grid_models.Col = 8
      cond2 = frmedmod.grid_models.Value
      cvalue = unity_main.lstmd.List(ii)
      
      If (cvalue < cond1) Then
        picType = "1"
      Else
        If (cvalue > cond2) Then
          picType = "3"
        Else
          If (cvalue > cond1) And (cvalue < cond2) Then
            picType = "2"
          End If
        End If
      End If

      unity_main.fpspread_pred.Col = 3
      unity_main.fpspread_pred.Row = ii + 1
      
      If (unity_main.m_olFormat = True) Then
        Select Case (picType)
          Case "1"              'pass
            unity_main.fpspread_pred.TypePictPicture = LoadPicture(GRAPHICS_DIR & "lt_go.BMP")
          Case "2"              'warning
            unity_main.fpspread_pred.TypePictPicture = LoadPicture(GRAPHICS_DIR & "lt_pause.BMP")
          Case "3"              'fail
            unity_main.fpspread_pred.TypePictPicture = LoadPicture(GRAPHICS_DIR & "lt_stop.BMP")
        End Select
        
        unity_main.fpspread_pred.TypePictStretch = True
        unity_main.fpspread_pred.TypePictMaintainScale = True
      Else
        rowHt = unity_main.fpspread_pred.RowHeight(ii + 1)
        unity_main.fpspread_pred.CellType = CellTypeNumber
        unity_main.fpspread_pred.TypeNumberDecPlaces = 2
        unity_main.fpspread_pred.FontSize = 10 '3/16
        unity_main.fpspread_pred.Text = cvalue
        
        If (unity_main.fpspread_pred.Text > 999.9) Then
          unity_main.fpspread_pred.CellType = CellTypeStaticText
          unity_main.fpspread_pred.Text = ">1000"
        End If
        
        unity_main.fpspread_pred.TypeHAlign = TypeHAlignLeft
        unity_main.fpspread_pred.RowHeight(ii + 1) = rowHt
        
        Select Case (picType)
          Case "1"              'pass
            unity_main.fpspread_pred.BackColor = &HFFFFFF
          Case "2"              'warning
            unity_main.fpspread_pred.BackColor = &H80FFFF
          Case "3"              'fail
            unity_main.fpspread_pred.BackColor = &H8080FF
        End Select
      End If
    
      ' Pause to allow background events
      DoEvents
NO_MDIST:
    Next ii
  End If
  
  ' Calculate residual
  For ii = 0 To frmedmod.numprops - 1
    tempstring = Trim(unity_main.lst_modtype.List(ii))
      
    ' Only PLS & CalStar models support residual
    If (tempstring = "2") Or (tempstring = "3") Or (tempstring = "4") Then
      unity_main.lstresrat.AddItem unity_main.m_noOLVal
      GoTo NO_RESID
    End If
      
    frmedmod.grid_models.Row = ii + 1
    frmedmod.grid_models.Col = 9
    cond1 = frmedmod.grid_models.Value
    frmedmod.grid_models.Col = 10
    cond2 = frmedmod.grid_models.Value
    res1 = unity_main.lstrr.List(ii)
    res2 = unity_main.lstrr2.List(ii)
    rratio = res1 / res2
    rratio = Format(rratio, "#0.0")
    unity_main.lstresrat.AddItem (rratio)
      
    ' show residual if configured
    If (frmedmod.chk_rr.Value = 1) Then
      cvalue = rratio
      
      If (cvalue < cond1) Then
        picType = "1"
      Else
        If (cvalue > cond2) Then
          picType = "3"
        Else
          If (cvalue > cond1) And (cvalue < cond2) Then
            picType = "2"
          End If
        End If
      End If

      unity_main.fpspread_pred.Col = 4
      unity_main.fpspread_pred.Row = ii + 1
      
      If (unity_main.m_olFormat = True) Then
        Select Case (picType)
          Case "1"              'pass
            unity_main.fpspread_pred.TypePictPicture = LoadPicture(GRAPHICS_DIR & "lt_go.BMP")
          Case "2"              'warning
            unity_main.fpspread_pred.TypePictPicture = LoadPicture(GRAPHICS_DIR & "lt_pause.BMP")
          Case "3"              'fail
            unity_main.fpspread_pred.TypePictPicture = LoadPicture(GRAPHICS_DIR & "lt_stop.BMP")
        End Select
        
        unity_main.fpspread_pred.TypePictStretch = True
        unity_main.fpspread_pred.TypePictMaintainScale = True
      Else
        rowHt = unity_main.fpspread_pred.RowHeight(ii + 1)
        unity_main.fpspread_pred.CellType = CellTypeNumber
        unity_main.fpspread_pred.TypeNumberDecPlaces = 2
        unity_main.fpspread_pred.FontSize = 10
        unity_main.fpspread_pred.Text = cvalue
       
        If (unity_main.fpspread_pred.Text > 999.9) Then
          unity_main.fpspread_pred.CellType = CellTypeStaticText
          unity_main.fpspread_pred.Text = ">1000"
        End If
        
        unity_main.fpspread_pred.TypeHAlign = TypeHAlignLeft
        unity_main.fpspread_pred.RowHeight(ii + 1) = rowHt
        
        Select Case (picType)
          Case "1"              'pass
            unity_main.fpspread_pred.BackColor = &HFFFFFF
          Case "2"              'warning
            unity_main.fpspread_pred.BackColor = &H80FFFF
          Case "3"              'fail
            unity_main.fpspread_pred.BackColor = &H8080FF
        End Select
      End If
    End If
    
    ' Pause to allow background events
    DoEvents
NO_RESID:
  Next ii

  ' Check value outlier if configured or required for LIMS, CSV and/or dymanic report
  If (frmedmod.chk_value.Value = 1) Or (unity_main.m_csvPropOutlier = 1) Or _
     (unity_main.pogpropoutlier = 1) Or (unity_main.m_saveDynRpt = True) Then
    For ii = 0 To frmedmod.numprops - 1
      'cond1=low warn cond2=low fail cond3=high warn cond4=high fail
      frmedmod.grid_models.Row = ii + 1
      frmedmod.grid_models.Col = 11
      cond1 = frmedmod.grid_models.Value
      frmedmod.grid_models.Col = 12
      cond2 = frmedmod.grid_models.Value
      frmedmod.grid_models.Col = 13
      cond3 = frmedmod.grid_models.Value
      frmedmod.grid_models.Col = 14
      cond4 = frmedmod.grid_models.Value
      unity_main.fpspread_pred.Col = 2
      unity_main.fpspread_pred.Row = ii + 1
      cvalue = unity_main.fpspread_pred.Value
  
      'test fail, if fail exit then test warn
      If (cvalue < cond2) Then
        picType = "3"
        unity_main.lst_qual.AddItem ("-2")
      Else
        If (cvalue > cond4) Then
          picType = "3"
          unity_main.lst_qual.AddItem ("2")
        Else
          If (cvalue < cond1) Then
            picType = "2"
            unity_main.lst_qual.AddItem ("-1")
          Else
            If (cvalue > cond3) Then
              picType = "2"
              unity_main.lst_qual.AddItem ("1")
            Else
              picType = "1"
              unity_main.lst_qual.AddItem ("0")
            End If
          End If
        End If
      End If
          
      ' Show value outlier if configured
      If (frmedmod.chk_value.Value = 1) Then
        unity_main.fpspread_pred.Col = 5
        unity_main.fpspread_pred.Row = ii + 1
      
        If (unity_main.m_olFormat = True) Then
          Select Case (picType)
            Case "1"              'pass
              unity_main.fpspread_pred.TypePictPicture = LoadPicture(GRAPHICS_DIR & "lt_go.BMP")
            Case "2"              'warning
              unity_main.fpspread_pred.TypePictPicture = LoadPicture(GRAPHICS_DIR & "lt_pause.BMP")
            Case "3"              'fail
              unity_main.fpspread_pred.TypePictPicture = LoadPicture(GRAPHICS_DIR & "lt_stop.BMP")
          End Select
        
          unity_main.fpspread_pred.TypePictStretch = True
          unity_main.fpspread_pred.TypePictMaintainScale = True
        Else
          rowHt = unity_main.fpspread_pred.RowHeight(ii + 1)
          unity_main.fpspread_pred.CellType = CellTypeNumber
          unity_main.fpspread_pred.TypeNumberDecPlaces = 2
          unity_main.fpspread_pred.FontSize = 10
          unity_main.fpspread_pred.CellType = CellTypeStaticText
          unity_main.fpspread_pred.Text = cond2 & "-" & cond4 'cvalue
          unity_main.fpspread_pred.TypeHAlign = TypeHAlignLeft
          unity_main.fpspread_pred.RowHeight(ii + 1) = rowHt
      
          Select Case (picType)
            Case "1"              'pass
              unity_main.fpspread_pred.BackColor = &HFFFFFF
            Case "2"              'warning
              unity_main.fpspread_pred.BackColor = &H80FFFF
            Case "3"              'fail
              unity_main.fpspread_pred.BackColor = &H8080FF
          End Select
        End If
      End If
    
      ' Pause to allow background events
      DoEvents
    Next ii
  End If
  
  ' Check if have working with PRD model(s)
  If (unity_main.m_prdModelType = True) Then
    ReDim labData(unity_main.m_numConstituents - 1)
    nn = 0
    expCodeFlg = False
  
    For ii = 0 To frmedmod.numprops - 1
      tempstring = Trim(unity_main.lst_modtype.List(ii))
      
      ' Only PRD models support ND
      If (tempstring <> 4) Then
        GoTo NO_ND
      End If
      
      pFExpCodes = unity_main.lst_pfexp.List(ii)
      ndPFExpCode = (pFExpCodes And &HF00) / &H100
      
      ' Check for expansion candidate
      If (ndPFExpCode = 2) Then
        labData(nn) = 999   ' mark constituent needs expansion
        expCodeFlg = True
      Else
        labData(nn) = 0
      End If
      
      nn = nn + 1
      
      ' Show neighborhood distance if configured
      If (frmedmod.chk_nd.Value = 1) Then
        Select Case (ndPFExpCode)
          Case 0            ' pass
            picType = "1"
          Case 1            ' merge candidate
            picType = "1"
          Case 2            ' expansion candidate
            picType = "2"
        End Select

        unity_main.fpspread_pred.Col = 6
        unity_main.fpspread_pred.Row = ii + 1
      
        If (unity_main.m_olFormat = True) Then
          Select Case (picType)
            Case "1"              'pass
              unity_main.fpspread_pred.TypePictPicture = LoadPicture(GRAPHICS_DIR & "lt_go.BMP")
            Case "2"              'warning
              unity_main.fpspread_pred.TypePictPicture = LoadPicture(GRAPHICS_DIR & "lt_pause.BMP")
            Case "3"              'fail
              unity_main.fpspread_pred.TypePictPicture = LoadPicture(GRAPHICS_DIR & "lt_stop.BMP")
          End Select
        
          unity_main.fpspread_pred.TypePictStretch = True
          unity_main.fpspread_pred.TypePictMaintainScale = True
        Else
          rowHt = unity_main.fpspread_pred.RowHeight(ii + 1)
          unity_main.fpspread_pred.CellType = CellTypeNumber
          unity_main.fpspread_pred.TypeNumberDecPlaces = 2
          unity_main.fpspread_pred.FontSize = 10 '3/16
          unity_main.fpspread_pred.Text = unity_main.lst_nd.List(ii)
        
          If (unity_main.fpspread_pred.Text > 999.9) Then
            unity_main.fpspread_pred.CellType = CellTypeStaticText
            unity_main.fpspread_pred.Text = ">1000"
          End If
        
          unity_main.fpspread_pred.TypeHAlign = TypeHAlignLeft
          unity_main.fpspread_pred.RowHeight(ii + 1) = rowHt
        
          Select Case (picType)
            Case "1"              'pass
              unity_main.fpspread_pred.BackColor = &HFFFFFF
            Case "2"              'warning
              unity_main.fpspread_pred.BackColor = &H80FFFF
            Case "3"              'fail
              unity_main.fpspread_pred.BackColor = &H8080FF
          End Select
        End If
      End If
      
      ' Pause to allow background events
      DoEvents
NO_ND:
    Next ii
    
    ' Check if need to save spectrum for expansion
    If (expCodeFlg = True) Then
      Dim svfFileExist As Boolean
      Dim fileName As String
      Dim fileNum As Long
      Dim svfFileDiff As Boolean
      Dim svfFileName As String
      Dim rc As Long
      Dim errMsg As String
    
      ' Check if SVF file exists to save expansion spectra data
      svfFileExist = unity_main.find_latest_svf_file(unity_main.m_saveDir, EXPANSION_SVF_FILE, fileName, fileNum)
      
      ' Check if SVF file different
      If (svfFileExist = True) And (unity_main.m_expSVFChanged = True) Then
        fileName = Replace(fileName, "_S" & fileNum, "_S" & (fileNum + 1))
        unity_main.m_expSVFChanged = False
        svfFileDiff = True
      End If

      svfFileName = (unity_main.m_saveDir & fileName)
      On Error GoTo OBJECT_ERROR
      
      ' Create file if does not exists or different
      If (svfFileExist = False) Or (svfFileDiff = True) Then
        rc = SVFObject.createFile_4("Unicode", svfFileName, unity_main.m_prdFileName, unity_main.m_stfFileName, unity_main.m_sysSerialNum, unity_main.m_stfMasterSerNum, unity_main.m_instModel, "", unity_main.lblProd1.Caption, "", unity_main.m_svfStartWvln, unity_main.m_svfEndWvln, unity_main.m_svfWaveInc, unity_main.m_svfIsStd, SVF_LAB_BASIS, SVF_WAVE_TYPE, unity_main.m_numConstituents, SampleConstituentNames)
        
        If (rc <> 0) Then GoTo SVF_ERROR
      End If
    
      rc = SVFObject.SaveSpectrum(svfFileName, -1, unity_main.m_scanDblTimestamp, unity_main.txtsamplename.Text, unity_main.lblProd1.Caption, SVFAbsYVals, 0, unity_main.m_numConstituents, labData)
      
      If (rc <> 0) Then
SVF_ERROR:
        errMsg = svfFileName & " UCal SVF spectra file error: " & rc
        uniMsg = MLSupport.GGS_Params("frmProduct.errMsg1", "%1 UCal SVF spectra file error: %2", svfFileName, CStr(rc))
        Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
        CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
      End If
    End If
  End If
  
  Exit Sub
  
OBJECT_ERROR:
  unity_main.errorstring = "Unity SVFComponent.dll component not installed or registered"
  unity_main.write_error (LOG_DBG_LEVEL1)
  uniMsg = MLSupport.GGS_Params("errMsg7", "Unity %1 component not installed or registered", "SVFComponent.dll")
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub

Sub calcavgpreds()
  Dim xx, ii, jj As Integer
  Dim sigfig
  Dim scanMsg As String
  Dim uniMsg As String

  scanMsg = "Performing repack properties value averaging"
  uniMsg = MLSupport.GSS("frm_collect", "statMsg4", "Performing repack properties value averaging")
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, scanMsg, uniMsg)
  
  frm_repacks.ss_repacks.MaxCols = frm_collect.numInc_numRepacks.Max + 2
  frm_repacks.ss_repacks.MaxRows = MAX_NUM_PROPS
  frm_repacks.ss_repacks.ClearRange 0, 0, frm_repacks.ss_repacks.MaxCols, frm_repacks.ss_repacks.MaxRows, False
  
  ' Add labels to columns
  frm_repacks.ss_repacks.Row = 0
  frm_repacks.ss_repacks.Col = 1
  frm_repacks.ss_repacks.Text = MLSupport.GSS("Headers", "property", "Property")
  frm_repacks.ss_repacks.Font.Bold = True
  frm_repacks.ss_repacks.Col = 2
  frm_repacks.ss_repacks.Text = MLSupport.GSS("Headers", "average", "Average")
  frm_repacks.ss_repacks.Font.Bold = True
  
  For xx = 3 To unity_main.m_repsAvg + 2
    frm_repacks.ss_repacks.Col = xx
    frm_repacks.ss_repacks.Text = MLSupport.GGS_Params("frm_collect.header1", "Repack %1", CStr((xx - 2)))
    frm_repacks.ss_repacks.Font.Bold = True
  Next xx
  
  For ii = 1 To Int(frmedmod.numprops.Text) 'numprops=rows
    ' Get property's # of significant digits
    frmedmod.grid_models.Col = 6
    frmedmod.grid_models.Row = ii
    sigfig = frmedmod.grid_models.Value
    
    frmpreds.gridpreds.Row = ii
    unity_main.sumavg = 0
   
    For jj = 1 To unity_main.m_repsAvg
      frmpreds.gridpreds.Col = jj
      unity_main.tempreal = frmpreds.gridpreds.Value
      unity_main.sumavg = unity_main.sumavg + unity_main.tempreal
      
      If (jj = 1) Then
        frm_repacks.ss_repacks.Col = 1
        unity_main.fpspread_pred.Col = 1
        
        For xx = 1 To unity_main.fpspread_pred.DataRowCnt 'get prop names in to repack ss
          unity_main.fpspread_pred.Row = xx
          frm_repacks.ss_repacks.Row = xx
          frm_repacks.ss_repacks.Text = unity_main.fpspread_pred.Text
        Next xx
      End If
      
      frm_repacks.ss_repacks.Col = (jj + 2)
      frm_repacks.ss_repacks.Row = ii
      frm_repacks.ss_repacks.Text = frmpreds.gridpreds.Text
      frm_repacks.ss_repacks.CellType = CellTypeNumber
      frm_repacks.ss_repacks.TypeNumberDecPlaces = sigfig
    Next jj
    
    frmpreds.gridpreds.Col = unity_main.m_repsAvg + 1
    frmpreds.gridpreds.Text = unity_main.sumavg / unity_main.m_repsAvg
    unity_main.tempval = frmpreds.gridpreds.Value
    
    frmedmod.grid_models.Row = ii '(was ii-1 on old form)
    frmedmod.grid_models.Col = 6 'sig fig col
    frmedmod.grid_models.Col = 5 ' skew
    frmedmod.grid_models.Col = 4 'bias/intercept
    frmpreds.gridpreds.Text = unity_main.tempval
    
    unity_main.fpspread_pred.Col = 2
    unity_main.fpspread_pred.Row = (ii) '+ 1)
    unity_main.fpspread_pred.Text = unity_main.tempval
    
    frm_repacks.ss_repacks.Col = 2
    frm_repacks.ss_repacks.Row = ii
    frm_repacks.ss_repacks.Text = unity_main.tempval
    frm_repacks.ss_repacks.CellType = CellTypeNumber
    frm_repacks.ss_repacks.TypeNumberDecPlaces = sigfig
  Next ii
  
  unity_main.m_scanDblTimestamp = Now
  unity_main.m_scanTimestamp = CDate(unity_main.m_scanDblTimestamp)
End Sub

Sub savescansettings(userReq As Boolean)
  Dim tempstring, temptext As String
  Dim strg(1 To 15) As String
  Dim nn, lenPath, rowCounter As Integer
  Dim tmpFile As String
  Dim prodFile As String
  Dim filePathName As String
  Dim uniFile As New clsUniFile
  Dim errMsg As String
  Dim uniMsg As String

  If (unity_main.passstring = "") Then
    Exit Sub
  End If
  
#If ABBFT Then
  ' Interferometer scanning parameters
  m_resolutionIndx = lst_resolution.ListIndex
  m_resolution = Trim(lst_resolution.List(lst_resolution.ListIndex))
  m_speedIndx = lst_speed.ListIndex
  m_gainIndx = lst_gain.ListIndex
  
  ' Interferometer CoAddition parameters
  m_numMeasures = txt_numMeasures.Text
  m_numSamples = numInc_numSamples.Text
  m_delayStart = txt_delayStart.Text
  m_delayMeasure = txt_delayMeasure.Text
  
  ' Interferometer sample/prediction parameters
  m_smplNumPts = txt_smplNumPts.Text
  m_waveNumIncr = txt_waveNumIncr.Text
  m_startWavenumIndx = txt_startWavenumIndx.Text
  m_startWavenum = txt_startWavenum.Text
  m_endWavenumIndx = txt_endWavenumIndx.Text
  m_endWavenum = txt_endWavenum.Text
#Else
  ' Determine if to use product default or custom wavelengths
  If (chk_dfltWavelens.Value = 0) Then
    frm_collect.m_smplEndWvln = frm_collect.txt_endWvln.Text
    frm_collect.m_smplStartWvln = frm_collect.txt_startWvln.Text
  Else
    frm_collect.m_smplEndWvln = frm_collect.txt_dfltEndWvln.Text
    frm_collect.m_smplStartWvln = frm_collect.txt_dfltStartWvln.Text
  End If
  
  ' Save external reference file name
  frm_collect.m_extRefFileName = frm_collect.combo_extRefFileName.Text
  
  ' Save multi-cup type selection
  frm_collect.m_multiCupType = CFG_NONE_MCT
  
  For nn = 1 To MAX_MULTI_CUP_TYPES
    If (opt_multiCupType(nn).Value = True) Then
      frm_collect.m_multiCupType = MS11MultiCupInfo(nn).cfgName
      GoTo LEAVE_LOOP
    End If
  Next nn
  
LEAVE_LOOP:
  ' Save number of sample scans
  m_smplNScans = numInc_smplNScans.Text
  
  ' Save offline reference file name
  frm_collect.m_olRefFileName = frm_collect.combo_olRefFileName.Text
  
  ' Save number of replicate samples
  frm_collect.m_repsAvg = numInc_numRepacks.Text
  
  ' Save rotation direction selection
  If (opt_rotateDirCW.Value = True) Then
    frm_collect.m_rotateDir = TRD_CW
  Else
    If (opt_rotateDirCCW.Value = True) Then
      frm_collect.m_rotateDir = TRD_CCW
    Else
      frm_collect.m_rotateDir = TRD_NONE
    End If
  End If
  
  ' Save rotation mode selection and values
  If (frm_collect.opt_rotateModeNone.Value = True) Then
    ' Setup for no rotation
    frm_collect.m_rotateMoveMode = TRM_NONE
    frm_collect.m_rotateIndexSteps = frm_collect.numInc_rotateIndexSteps.Text
    frm_collect.m_rotateSpeed = frm_collect.numInc_rotateSpeed.Text
    frm_collect.m_rotateStepSteps = frm_collect.numInc_rotateStepSteps.Text
  Else
    If (frm_collect.opt_rotateModeCont.Value = True) Then
      ' Setup for continuous rotation
      frm_collect.m_rotateMoveMode = TRM_CONT
      frm_collect.m_rotateIndexSteps = 0
      frm_collect.m_rotateSpeed = frm_collect.numInc_rotateSpeed.Text
      frm_collect.m_rotateStepSteps = 0
    Else
      If (frm_collect.opt_rotateModeStep.Value = True) Then
        ' Setup for steps for stepped rotation
        frm_collect.m_rotateMoveMode = TRM_STEP
        frm_collect.m_rotateIndexSteps = 0
        frm_collect.m_rotateSpeed = 0
        frm_collect.m_rotateStepSteps = frm_collect.numInc_rotateStepSteps.Text
      Else
        If (frm_collect.opt_rotateModeIndex.Value = True) Then
          ' Setup for steps for indexed rotation
          frm_collect.m_rotateMoveMode = TRM_INDEX
          frm_collect.m_rotateIndexSteps = frm_collect.numInc_rotateIndexSteps.Text
          frm_collect.m_rotateSpeed = 0
          frm_collect.m_rotateStepSteps = 0
        End If
      End If
    End If
  End If

  If (frm_collect.chk_savereps.Value = 1) Then
    frm_collect.m_saveReps = True
  Else
    frm_collect.m_saveReps = False
  End If
    
  ' Save use external reference tray configuration
  m_useExtRefTrayCfg = CBool(frm_collect.chk_useExtRefTrayCfg.Value)
#End If

  ' Save background frequency
  If (frm_collect.opt_backall.Value = True) Then
    frm_collect.m_backFreq = REF_FREQ_ALL_SMPLS
  Else
    If (frm_collect.opt_backdemand.Value = True) Then
      frm_collect.m_backFreq = REF_FREQ_ON_DEMAND
    End If
  End If

  ' Save bound value selection
  Call checkbound 'take care of checking value limits

  ' Save reference type selection
  If (optbginternal.Value = True) Then
    m_bType = "internal"
  Else
    If (optbgexternal.Value = True) Then
      m_bType = "external"
    Else
      If (optbgfile.Value = True) Then
        m_bType = "file"
      End If
    End If
  End If

  'Save clear manual name entry
  If (frm_collect.chk_clrManualName.Value = 1) Then
    m_clrManualName = True
  Else
    m_clrManualName = False
  End If
  
  'Save clear user inputs entry
  If (frm_collect.chk_clrUserInputs.Value = 1) Then
    m_clrUserInputs = True
  Else
    m_clrUserInputs = False
  End If
  
  ' Save day counter
  frm_collect.m_dayCounter = frm_collect.numInc_dateCounter.Text
  
  ' Save hide prediction value column selection
  If (frm_collect.chk_hideValCol.Value = 1) Then
    frm_collect.m_hideValCol = True
  Else
    frm_collect.m_hideValCol = False
  End If
  
  ' Save make predictions selection
  If (optpredyes.Value = True) Then
    m_makePred = "yes"
  Else
    m_makePred = "no"
  End If
  
  ' Save M-Dist alarm setting
  frm_collect.m_alarmMD = frmedmod.chk_md.Value
  
  ' Save use input button/list selection
  frm_collect.m_useMIV = frm_collect.chk_userInputs.Value
  
  ' Save sample base file name
  m_nameBase = Trim(txtsampnamebase.Text)
  
  ' Save sample base file name counter
  If (numInc_nameCounter.Text = "") Then
    numInc_nameCounter.Text = "0"
  End If
  
  m_nameCounter = numInc_nameCounter.Text
  
  ' Save sample name type
  If (optnamemanual.Value = True) Then
    m_nameScanType = "Manual"
    m_sNameMode = 1
  Else
    If (optnamecounter.Value = True) Then
      m_nameScanType = "Counter"
      m_sNameMode = 3
    Else
      If (opt_namedate.Value = True) Then
        m_nameScanType = "Date"
        m_sNameMode = 4
      End If
    End If
  End If
  
  ' Save neighborhood distance alarm setting
  frm_collect.m_alarmND = frmedmod.chk_nd.Value
  
  ' Save no outlier reported value
  m_noOLVal = Trim(txt_noOLVal.Text)
  
  ' Save no prediction reported value
  m_noPredVal = Trim(txt_noPredVal.Text)

  ' Save outlier format selection
  If (frm_collect.opt_ollts.Value = True) Then
    frm_collect.m_olFormat = True
  Else
    frm_collect.m_olFormat = False
  End If
  
  ' Save residual alarm selection
  frm_collect.m_alarmRR = frmedmod.chk_rr.Value
  
  ' Save scan file selection
  If (optsavescanyes.Value = True) Then
    m_saveIt = "save"
  Else
    m_saveIt = "nosave"
  End If

  ' Confirm path contains '\' instead of '/'
  filePathName = Trim(frm_collect.txt_caldir.Text)
  check_filepathname_delimiters filePathName
  frm_collect.txt_caldir.Text = filePathName
  
  ' Append "\" to spectrum file path if not present
  If (Right(frm_collect.txt_caldir.Text, 1) <> "\") Then
    frm_collect.txt_caldir.Text = frm_collect.txt_caldir.Text & "\"
  End If
  
  ' Save spectrum file directory
  frm_collect.m_saveDir = frm_collect.txt_caldir.Text
  
  ' Confirm path contains '\' instead of '/'
  filePathName = Trim(frm_collect.txt_csvfilename.Text)
  check_filepathname_delimiters filePathName
  frm_collect.txt_csvfilename.Text = filePathName
  
  If (InStr(frm_collect.txt_csvfilename.Text, "\") = 0) Then
    frm_collect.txt_csvfilename.Text = (REPORTS_DIR & filePathName)
  End If

  ' Save csv path\file name
  frm_collect.m_saveCSVFile = frm_collect.txt_csvfilename.Text

  ' Confirm path contains '\' instead of '/'
  filePathName = Trim(frm_collect.txtpredfile.Text)
  check_filepathname_delimiters filePathName
  frm_collect.txtpredfile.Text = filePathName
  
  If (InStr(frm_collect.txtpredfile.Text, "\") = 0) Then
    frm_collect.txtpredfile.Text = (REPORTS_DIR & filePathName)
  End If

  ' Save report path\file name
  frm_collect.m_savePredFile = frm_collect.txtpredfile.Text

  ' Save CSV report option
  If (frm_collect.optcsvno.Value = True) Then
    frm_collect.m_saveCSV = False
  Else
    frm_collect.m_saveCSV = True
  End If
  
  ' Save dynamic report option
  If (frm_collect.opt_DynRptNo.Value = True) Then
    frm_collect.m_saveDynRpt = False
  Else
    frm_collect.m_saveDynRpt = True
  End If
  
  ' Save LIMS output
  frm_collect.m_doLIMS = chk_lims.Value
  
  ' Save sample PPT
  frm_collect.m_smplPPT = frm_collect.numInc_smplPPT.Text
    
  ' Save sample type
  If (optbk.Value = True) Then
    m_sType = "back"
  Else
    If (optabs.Value = True) Then
      m_sType = "abs"
    End If
  End If
  
  ' Save value alarm selection
  frm_collect.m_alarmProp = frmedmod.chk_value.Value
  
  ' Save ticket report selection
  If (frm_collect.opt_tktno.Value = True) Then
    frm_collect.m_writeTkt = 0
  Else
    If (frm_collect.opt_tktall.Value = True) Then
      frm_collect.m_writeTkt = 1
    Else
      If frm_collect.opt_tktdemand.Value = True Then
        frm_collect.m_writeTkt = 2
      End If
    End If
  End If
  
  On Error GoTo FILE_ERROR
  tmpFile = (PRODUCTS_CFG_DIR & TMP_SAVE_PROD_CFG_FILE)
  prodFile = (PRODUCTS_CFG_DIR & Trim(unity_main.passstring))
  
  If (uniFile.OpenFileWrite(tmpFile) = True) Then
    On Error GoTo FILE_ERROR
    uniFile.WriteBOM fe_UTF16LE
    uniFile.WriteUnicodeLine "[product type, sampling]"
    uniFile.WriteUnicodeLine (Trim(unity_main.m_productName) & "," & Trim(unity_main.current_sampling))
    uniFile.WriteUnicodeLine "[signature settings]"
#If ABBFT Then
    uniFile.WriteUnicodeLine ("DevID=" & DTID_ABBFT)
    uniFile.WriteUnicodeLine ("Version=" & INFOSTAR_VER)
    uniFile.WriteUnicodeLine "[analyzer settings]"
    uniFile.WriteUnicodeLine ("Background_Frequency=" & frm_collect.m_backFreq)
    uniFile.WriteUnicodeLine ("Bound_Values=" & frm_collect.m_valueBound)
    uniFile.WriteUnicodeLine ("BType=" & frm_collect.m_bType)
    uniFile.WriteUnicodeLine ("ClrManualName=" & frm_collect.m_clrManualName)
    uniFile.WriteUnicodeLine ("ClrUserInputs=" & frm_collect.m_clrUserInputs)
    uniFile.WriteUnicodeLine ("Day_Counter=" & frm_collect.m_dayCounter)
    uniFile.WriteUnicodeLine ("DelayMeasure=" & frm_collect.m_delayMeasure)
    uniFile.WriteUnicodeLine ("DelayStart=" & frm_collect.m_delayStart)
    uniFile.WriteUnicodeLine ("EndWavenumIndx=" & frm_collect.m_endWavenumIndx)
    uniFile.WriteUnicodeLine ("GainIndx=" & frm_collect.m_gainIndx)
    uniFile.WriteUnicodeLine ("HideValCol=" & frm_collect.m_hideValCol)
    uniFile.WriteUnicodeLine ("MakePred=" & frm_collect.m_makePred)
    uniFile.WriteUnicodeLine ("MD_Alarm=" & frm_collect.m_alarmMD)
    uniFile.WriteUnicodeLine ("Menu_Input_Buttons=" & frm_collect.m_useMIV)
    uniFile.WriteUnicodeLine ("NameBase=" & frm_collect.m_nameBase)
    uniFile.WriteUnicodeLine ("NameCounter=" & frm_collect.m_nameCounter)
    uniFile.WriteUnicodeLine ("NameScanType=" & frm_collect.m_nameScanType)
    uniFile.WriteUnicodeLine ("ND_Alarm=" & frm_collect.m_alarmND)
    uniFile.WriteUnicodeLine ("NoOLVal=" & frm_collect.m_noOLVal)
    uniFile.WriteUnicodeLine ("NoRefVal=" & frm_collect.m_noPredVal)
    uniFile.WriteUnicodeLine ("NumMeasures=" & frm_collect.m_numMeasures)
    uniFile.WriteUnicodeLine ("NumSamples=" & frm_collect.m_numSamples)
    uniFile.WriteUnicodeLine ("Outlier_Lights=" & frm_collect.m_olFormat)
    uniFile.WriteUnicodeLine ("ResolutionIndx=" & frm_collect.m_resolutionIndx)
    uniFile.WriteUnicodeLine ("RR_Alarm=" & frm_collect.m_alarmRR)
    uniFile.WriteUnicodeLine ("SaveIt=" & frm_collect.m_saveIt)
    uniFile.WriteUnicodeLine ("SaveScansDir=" & frm_collect.m_saveDir)
    uniFile.WriteUnicodeLine ("Save_CsvFile=" & frm_collect.m_saveCSVFile)
    uniFile.WriteUnicodeLine ("Save_PredFile=" & frm_collect.m_savePredFile)
    uniFile.WriteUnicodeLine ("Save_Predictions=" & frm_collect.m_savePredictions)
    uniFile.WriteUnicodeLine ("Save_Predictions_Csv=" & frm_collect.m_saveCSV)
    uniFile.WriteUnicodeLine ("Save_Predictions_DynRpt=" & frm_collect.m_saveDynRpt)
    uniFile.WriteUnicodeLine ("Send_LIMS_Output=" & frm_collect.m_doLIMS)
    uniFile.WriteUnicodeLine ("SpeedIndx=" & frm_collect.m_speedIndx)
    uniFile.WriteUnicodeLine ("StartWavenumIndx=" & frm_collect.m_startWavenumIndx)
    uniFile.WriteUnicodeLine ("SType=" & frm_collect.m_sType)
    uniFile.WriteUnicodeLine ("Value_Alarm=" & frm_collect.m_alarmProp)
    uniFile.WriteUnicodeLine ("Write_Ticket_Printer=" & frm_collect.m_writeTkt)
#Else
    uniFile.WriteUnicodeLine ("DevID=" & MS11CfgData.devID)
    uniFile.WriteUnicodeLine ("SmplTable=" & unity_main.m_smplTable)
    uniFile.WriteUnicodeLine ("ScanMode=" & unity_main.m_sysScanMode)
    uniFile.WriteUnicodeLine ("Version=" & INFOSTAR_VER)
    uniFile.WriteUnicodeLine "[analyzer settings]"
    uniFile.WriteUnicodeLine ("AdapterType=" & frm_collect.m_adapterType)
    uniFile.WriteUnicodeLine ("Background_Frequency=" & frm_collect.m_backFreq)
    uniFile.WriteUnicodeLine ("Bound_Values=" & frm_collect.m_valueBound)
    uniFile.WriteUnicodeLine ("BType=" & frm_collect.m_bType)
    uniFile.WriteUnicodeLine ("ClrManualName=" & frm_collect.m_clrManualName)
    uniFile.WriteUnicodeLine ("ClrUserInputs=" & frm_collect.m_clrUserInputs)
    uniFile.WriteUnicodeLine ("Day_Counter=" & frm_collect.m_dayCounter)
    uniFile.WriteUnicodeLine ("EndWvln=" & frm_collect.m_smplEndWvln)
    uniFile.WriteUnicodeLine ("ExtRefFile=" & frm_collect.m_extRefFileName)
    uniFile.WriteUnicodeLine ("HideValCol=" & frm_collect.m_hideValCol)
    uniFile.WriteUnicodeLine ("MakePred=" & frm_collect.m_makePred)
    uniFile.WriteUnicodeLine ("MD_Alarm=" & frm_collect.m_alarmMD)
    uniFile.WriteUnicodeLine ("Menu_Input_Buttons=" & frm_collect.m_useMIV)
    uniFile.WriteUnicodeLine ("MultiCupType=" & frm_collect.m_multiCupType)
    uniFile.WriteUnicodeLine ("NameBase=" & frm_collect.m_nameBase)
    uniFile.WriteUnicodeLine ("NameCounter=" & frm_collect.m_nameCounter)
    uniFile.WriteUnicodeLine ("NameScanType=" & frm_collect.m_nameScanType)
    uniFile.WriteUnicodeLine ("ND_Alarm=" & frm_collect.m_alarmND)
    uniFile.WriteUnicodeLine ("NoOLVal=" & frm_collect.m_noOLVal)
    uniFile.WriteUnicodeLine ("NoRefVal=" & frm_collect.m_noPredVal)
    uniFile.WriteUnicodeLine ("NScansS=" & frm_collect.m_smplNScans)
    uniFile.WriteUnicodeLine ("OLRefFile=" & frm_collect.m_olRefFileName)
    uniFile.WriteUnicodeLine ("Outlier_Lights=" & frm_collect.m_olFormat)
    uniFile.WriteUnicodeLine ("RepsAvg=" & frm_collect.m_repsAvg)
    uniFile.WriteUnicodeLine ("RotateDir=" & frm_collect.m_rotateDir)
    uniFile.WriteUnicodeLine ("RotateIndexSteps=" & frm_collect.m_rotateIndexSteps)
    uniFile.WriteUnicodeLine ("RotateMoveMode=" & frm_collect.m_rotateMoveMode)
    uniFile.WriteUnicodeLine ("RotateSpeed=" & frm_collect.m_rotateSpeed)
    uniFile.WriteUnicodeLine ("RotateStepSteps=" & frm_collect.m_rotateStepSteps)
    uniFile.WriteUnicodeLine ("RR_Alarm=" & frm_collect.m_alarmRR)
    uniFile.WriteUnicodeLine ("SaveIt=" & frm_collect.m_saveIt)
    uniFile.WriteUnicodeLine ("SaveScansDir=" & frm_collect.m_saveDir)
    uniFile.WriteUnicodeLine ("Save_CsvFile=" & frm_collect.m_saveCSVFile)
    uniFile.WriteUnicodeLine ("Save_PredFile=" & frm_collect.m_savePredFile)
    uniFile.WriteUnicodeLine ("Save_Predictions=" & frm_collect.m_savePredictions)
    uniFile.WriteUnicodeLine ("Save_Predictions_Csv=" & frm_collect.m_saveCSV)
    uniFile.WriteUnicodeLine ("Save_Predictions_DynRpt=" & frm_collect.m_saveDynRpt)
    uniFile.WriteUnicodeLine ("Save_Replicates=" & frm_collect.m_saveReps)
    uniFile.WriteUnicodeLine ("Send_LIMS_Output=" & frm_collect.m_doLIMS)
    uniFile.WriteUnicodeLine ("SmplPPT=" & frm_collect.m_smplPPT)
    uniFile.WriteUnicodeLine ("StartWvln=" & frm_collect.m_smplStartWvln)
    uniFile.WriteUnicodeLine ("SType=" & frm_collect.m_sType)
    uniFile.WriteUnicodeLine ("UseExtRefTrayCfg=" & frm_collect.m_useExtRefTrayCfg)
    uniFile.WriteUnicodeLine ("Value_Alarm=" & frm_collect.m_alarmProp)
    uniFile.WriteUnicodeLine ("Write_Ticket_Printer=" & frm_collect.m_writeTkt)
#End If

    'now do the model table
    uniFile.WriteUnicodeLine "[analysis models]"
  
    For rowCounter = 1 To frmedmod.grid_models.MaxRows
      frmedmod.grid_models.Row = rowCounter
      frmedmod.grid_models.Col = 1
      strg(1) = Trim(frmedmod.grid_models.Text)
    
      If Trim(strg(1)) = "" Then
        Exit For
      End If
      
      ' Enclose each model variable within ""
      strg(1) = Chr(34) & strg(1) & Chr(34)
      frmedmod.grid_models.Col = 2
      strg(2) = Chr(34) & Trim(frmedmod.grid_models.Text) & Chr(34)
      frmedmod.grid_models.Col = 3
      strg(3) = Chr(34) & Trim(frmedmod.grid_models.Text) & Chr(34)
      frmedmod.grid_models.Col = 4
      strg(4) = Chr(34) & Trim(frmedmod.grid_models.Text) & Chr(34)
      frmedmod.grid_models.Col = 5
      strg(5) = Chr(34) & Trim(frmedmod.grid_models.Text) & Chr(34)
      frmedmod.grid_models.Col = 6
      strg(6) = Chr(34) & Trim(frmedmod.grid_models.Text) & Chr(34)
      frmedmod.grid_models.Col = 7
      strg(7) = Chr(34) & Trim(frmedmod.grid_models.Text) & Chr(34)
      frmedmod.grid_models.Col = 8
      strg(8) = Chr(34) & Trim(frmedmod.grid_models.Text) & Chr(34)
      frmedmod.grid_models.Col = 9
      strg(9) = Chr(34) & Trim(frmedmod.grid_models.Text) & Chr(34)
      frmedmod.grid_models.Col = 10
      strg(10) = Chr(34) & Trim(frmedmod.grid_models.Text) & Chr(34)
      frmedmod.grid_models.Col = 11
      strg(11) = Chr(34) & Trim(frmedmod.grid_models.Text) & Chr(34)
      frmedmod.grid_models.Col = 12
      strg(12) = Chr(34) & Trim(frmedmod.grid_models.Text) & Chr(34)
      frmedmod.grid_models.Col = 13
      strg(13) = Chr(34) & Trim(frmedmod.grid_models.Text) & Chr(34)
      frmedmod.grid_models.Col = 14
      strg(14) = Chr(34) & Trim(frmedmod.grid_models.Text) & Chr(34)
      frmedmod.grid_models.Col = 15
      strg(15) = Chr(34) & Trim(frmedmod.grid_models.Text) & Chr(34)
      uniFile.WriteUnicodeLine (strg(1) & "," & strg(2) & "," & strg(3) & "," & strg(4) & "," & strg(5) & "," & strg(6) & "," & strg(7) & "," & strg(8) & "," & strg(9) & "," & strg(10) & "," & strg(11) & "," & strg(12) & "," & strg(13) & "," & strg(14) & "," & strg(15))
    Next rowCounter

    uniFile.Flush
    uniFile.CloseFile
        
    If (uniFile.st_FileExist(prodFile) = True) Then
      uniFile.st_SetFileAttr prodFile, vbNormal
      uniFile.st_RmFile prodFile
    End If
  
    uniFile.st_CopyFile tmpFile, prodFile
    uniFile.st_RmFile tmpFile
  
    If (userReq = True) Then
      unity_main.errorstring = ("User saved new settings for configuration file: " & prodFile)
      unity_main.write_error (LOG_DBG_LEVEL1)
    End If
  
    Call frmedmod.fixthesize
  Else
FILE_ERROR:
    uniFile.CloseFile
  
    If (tmpFile <> "") And (uniFile.st_FileExist(tmpFile) = True) Then
      uniFile.st_RmFile tmpFile
    End If
 
    errMsg = (prodFile & " file write error. " & Error$)
    unity_main.errorstring = errMsg
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", prodFile, Error$)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  End If
End Sub

Sub checkbound()
  
  If (frm_collect.opt_boundno.Value = True) Then
    m_valueBound = 0
  Else
    If (frm_collect.opt_boundl.Value = True) Then
      m_valueBound = 1
    Else
      If (frm_collect.opt_boundh.Value = True) Then
        m_valueBound = 2
      Else
        If (frm_collect.opt_boundhl.Value = True) Then
          m_valueBound = 3
        End If
      End If
    End If
  End If
End Sub

Sub makedatename()
  Dim datex As String
  Dim searchName As String
  Dim tempFound As Boolean
  Dim dirname As String
  Dim buildName As String
  Dim fileName As String
  Dim ctr As Long
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim rc As Boolean

  datex = Date
  
hadit:
  fileName = datex
  Call rebuild_date(Date, buildName)
  buildName = buildName & "_"

  ' Check if spectrum file to be saved
  If (unity_main.m_saveIt = "save") Then
    ' If no spectrum file with today's date then set date counter to 1
    If (any_files_exist(unity_main.m_saveDir, buildName & "*" & SPC_FILE_EXT) = False) Then
      frm_collect.numInc_dateCounter.Text = 1
      buildName = buildName & frm_collect.numInc_dateCounter.Text
    Else
      ' Build file name to see if it already exits
      buildName = buildName & frm_collect.numInc_dateCounter.Text
      fileName = buildName & SPC_FILE_EXT
      dirname = unity_main.m_saveDir
      tempFound = CFile.st_FileExist(dirname & fileName)

      ' Increment counter if file exists
      If (tempFound = True) Then
        ctr = frm_collect.numInc_dateCounter.Text + 1
      
        If (ctr > frm_collect.numInc_dateCounter.Max) Then
          ctr = frm_collect.numInc_dateCounter.Min
        End If
      
        frm_collect.numInc_dateCounter.Text = ctr
        GoTo hadit
      End If
    End If
  Else
    fileName = unity_main.m_saveDir & TMP_DATE_CTR_FILE
    tempFound = CFile.st_FileExist(fileName)
  
    If (tempFound = False) Then
      frm_collect.numInc_dateCounter.Text = 1
    Else
      If (uniFile.OpenFileRead(fileName) = False) Then GoTo BAD_FILE
  
      On Error GoTo BAD_FILE
      fEncoding = uniFile.ReadBOM
      
      If (fEncoding = fe_ANSI) Then
        rc = uniFile.ReadAnsiLine(searchName)
      Else
        rc = uniFile.ReadUnicodeLine(searchName)
      End If
      
      If (rc = False) Then GoTo BAD_FILE
      
      If (searchName = datex) Then
        ctr = frm_collect.numInc_dateCounter.Text + 1
      
        If (ctr > frm_collect.numInc_dateCounter.Max) Then
          ctr = frm_collect.numInc_dateCounter.Min
        End If
      
        frm_collect.numInc_dateCounter.Text = ctr
      Else
BAD_FILE:
        frm_collect.numInc_dateCounter.Text = 1
      End If
      
      uniFile.CloseFile
    End If
  
    buildName = buildName & frm_collect.numInc_dateCounter.Text
    
    If (frm_collect.numInc_dateCounter.Text = 1) Then
      If (uniFile.OpenFileWrite(fileName) = True) Then
        uniFile.WriteBOM fe_UTF16LE
        uniFile.WriteUnicodeLine datex
        uniFile.Flush
      End If
      
      uniFile.CloseFile
    End If
  End If
  
  unity_main.txtsamplename.Text = Trim(buildName)
End Sub

Sub rebuild_date(orgDateStrg As Variant, newDateStrg As String)
  Dim i As Integer
  Dim S As String

  i = 1
  newDateStrg = ""
  
  Do While (i <= Len(orgDateStrg))
    S = Mid(orgDateStrg, i, 1)
    
    If (S = "\") Or (S = "/") Or (S = ":") Or (S = "*") Or (S = "?") Or (S = "<") Or _
       (S = ">") Or (S = "|") Or (S = Chr(34)) Then
      S = "-"
    End If
        
    newDateStrg = newDateStrg & S
    i = i + 1
  Loop
End Sub

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

#If SSTAR Then
Public Sub build_ref_name_list(refType As String)
  Dim nn As Integer
  
  If (refType = "external") Then
    file_refNames.Refresh
    file_refNames.Path = EXT_REFS_CFG_DIR
    file_refNames.Pattern = ("*" & EXT_REF_SCAN_FILE & "*" & CFG_FILE_EXT)
    combo_extRefFileName.Clear
    frm_extRefMgmt.lst_refFileNames.Clear
    frm_extRefPPT.lst_refFileNames.Clear
    
    For nn = 0 To file_refNames.ListCount - 1
      If (file_refNames.List(nn) <> "..") Then
        combo_extRefFileName.AddItem CFile.st_FileNameNoExt(file_refNames.List(nn))
        frm_extRefMgmt.lst_refFileNames.AddItem CFile.st_FileNameNoExt(file_refNames.List(nn))
        frm_extRefPPT.lst_refFileNames.AddItem CFile.st_FileNameNoExt(file_refNames.List(nn))
      End If
    Next nn
    
    combo_extRefFileName.AddItem ""
  Else
    file_refNames.Path = REFERENCES_DIR
    file_refNames.Pattern = ("*" & OFFLINE_REF_SCAN_FILE & "*" & SPC_FILE_EXT)
    combo_olRefFileName.Clear
    
    For nn = 0 To file_refNames.ListCount - 1
      If (file_refNames.List(nn) <> "..") Then
        combo_olRefFileName.AddItem CFile.st_FileNameNoExt(file_refNames.List(nn))
      End If
    Next nn
    
    combo_olRefFileName.AddItem ""
  End If
End Sub
#End If

Private Sub chk_lims_Click()
  
  frm_collect.m_doLIMS = chk_lims.Value
End Sub

#If SSTAR Then
Private Sub chk_useExtRefTrayCfg_Click()

  If (chk_useExtRefTrayCfg.Value = 0) Then
    frame_adpaterType.enabled = True
    frame_multiCupType.enabled = True
    frame_rotateDir.enabled = True
    frame_rotateMode.enabled = True
    frame_rotatePlatter.enabled = True
    frame_trayCfg.enabled = True
    
    If (m_ignoreEvent = False) Then
      frm_collect.m_adapterType = ProdDfltData.adapterType
      frm_collect.m_multiCupType = ProdDfltData.multiCupType
      frm_collect.m_rotateDir = ProdDfltData.rotateDir
      frm_collect.m_rotateIndexSteps = ProdDfltData.rotateIndexSteps
      frm_collect.m_rotateMoveMode = ProdDfltData.rotateMoveMode
      frm_collect.m_rotateSpeed = ProdDfltData.rotateSpeed
      frm_collect.m_rotateStepSteps = ProdDfltData.rotateStepSteps
      setup_prod_tray_cfg
    End If
  Else
    frame_adpaterType.enabled = False
    frame_multiCupType.enabled = False
    frame_rotateDir.enabled = False
    frame_rotateMode.enabled = False
    frame_rotatePlatter.enabled = False
    frame_trayCfg.enabled = False
    
    If (m_ignoreEvent = False) Then
      frm_collect.m_adapterType = frm_extRef.m_extRefAdapterType
      frm_collect.m_multiCupType = frm_extRef.m_extRefMultiCupType
      frm_collect.m_rotateDir = frm_extRef.m_extRefRotateDir
      frm_collect.m_rotateIndexSteps = frm_extRef.m_extRefRotateIndexSteps
      frm_collect.m_rotateMoveMode = frm_extRef.m_extRefRotateMoveMode
      frm_collect.m_rotateSpeed = frm_extRef.m_extRefRotateSpeed
      frm_collect.m_rotateStepSteps = frm_extRef.m_extRefRotateStepSteps
      setup_prod_tray_cfg
    End If
  End If
End Sub
#End If

Private Sub cmd_cancel_Click()

  unity_main.errorstring = "Product Configuration screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Call unity_main.load_prod_file("", False)
  frm_collect.Visible = False
End Sub

Private Sub cmd_save_Click()
  
  ' Check if valid configuration
  If (check_scan_settings = True) Then
    unity_main.errorstring = "Product Configuration screen Save Changes button selected"
    unity_main.write_error (LOG_DBG_LEVEL3)
    
    Call frm_collect.savescansettings(True)
    frm_collect.Visible = False
    
    Call unity_main.load_prod_file("", True)
    unity_main.txtsamplename.Text = ""
    unity_main.txtsampcomment.Text = ""
    unity_main.repcounter = 0
  End If
End Sub

#If SSTAR Then
Private Sub combo_extRefFileName_Click()
  
  If (m_ignoreEvent = False) Then
    If (frm_collect.combo_extRefFileName.ListIndex >= 0) Then
      If (frm_collect.combo_extRefFileName.List(frm_collect.combo_extRefFileName.ListIndex) <> "") Then
        Call update_prod_ext_ref_wvlns(frm_collect.combo_extRefFileName.List(frm_collect.combo_extRefFileName.ListIndex))
      End If
    End If
  End If
End Sub
#End If

#If SSTAR Then
Private Sub combo_olRefFileName_Click()
  Dim fileName As String

  If (m_ignoreEvent = False) Then
    If (frm_collect.combo_olRefFileName.ListIndex >= 0) Then
      If (frm_collect.combo_olRefFileName.List(frm_collect.combo_olRefFileName.ListIndex) <> "") Then
        Call update_prod_ol_ref_wvlns(frm_collect.combo_olRefFileName.List(frm_collect.combo_olRefFileName.ListIndex))
      End If
    End If
  End If
End Sub
#End If

Private Sub Form_Activate()
  Dim dateStrg As String
  Dim ii As Integer

  Select Case (m_sNameMode)
    Case 1
      txt_scanname.Text = MLSupport.GSS("frm_collect", "txt_scanname", "Manual Entry")
    Case 3
      txt_scanname.Text = Trim(txtsampnamebase.Text) & numInc_nameCounter.Text
    Case 4
      Call rebuild_date(Date, dateStrg)
      txt_scanname.Text = dateStrg & "_" & numInc_dateCounter.Text
  End Select
End Sub

Private Sub Form_Load()
  Dim ii As Integer

  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
  ' Setup tab headers
#If ABBFT Then
  tab_frame.AddTab MLSupport.GSS("frm_collect", "TabCaption0", "Scan Settings")
  tab_frame.AddTab MLSupport.GSS("frm_collect", "TabCaption2", "Sample Naming")
  tab_frame.AddTab MLSupport.GSS("frm_collect", "TabCaption3", "Outputs")
  tab_frame.AddTab MLSupport.GSS("frm_collect", "TabCaption4", "Miscellaneous")
#Else
  tab_frame.AddTab MLSupport.GSS("frm_collect", "TabCaption0", "Scan Settings")
  tab_frame.AddTab MLSupport.GSS("frm_collect", "TabCaption1", "Tray Settings")
  tab_frame.AddTab MLSupport.GSS("frm_collect", "TabCaption2", "Sample Naming")
  tab_frame.AddTab MLSupport.GSS("frm_collect", "TabCaption3", "Outputs")
  tab_frame.AddTab MLSupport.GSS("frm_collect", "TabCaption4", "Miscellaneous")
#End If
  ' Setup to display first tab frame
  show_frame 0
  
#If ABBFT Then
  Frame5.Visible = False
  Frame9.Visible = False
  Frame15.Visible = False
  chk_useExtRefTrayCfg.Visible = False
  optbginternal.enabled = False
  optbgfile.enabled = False
  opt_backall.enabled = False
#Else
  frame_scanning.Visible = False
  frame_coAdditon.Visible = False
  frame_prediction.Visible = False
#End If
End Sub

#If ABBFT Then
Private Sub lst_resolution_Click()
  Dim numPts As Long
  Dim waveNumIncr As Double
  
  Call calc_wavenum_incr(lst_resolution.List(lst_resolution.ListIndex), numPts, waveNumIncr)
  txt_smplNumPts.Text = numPts
  txt_waveNumIncr.Text = waveNumIncr
  
  txt_startWavenumIndx_Change
  txt_endWavenumIndx_Change
End Sub
#End If

Private Sub numInc_dateCounter_Change()
  
  If (m_sNameMode = 4) Then
    Call opt_namedate_Click
  End If
End Sub

Private Sub numInc_dateCounter_DblClick()
  
  unity_main.formfrom = 2
  unity_main.varfrom = 9
  frm_numpad.lbl_num.Caption = MLSupport.GSS("frm_collect", "lbl_num2", "Date Counter Value")
  frm_numpad.txt_num.Text = frm_collect.numInc_dateCounter.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_nameCounter_Change()
  
  If (m_sNameMode = 3) Then
    Call optnamecounter_Click
  End If
End Sub

Private Sub numInc_nameCounter_DblClick()

  unity_main.formfrom = 2
  unity_main.varfrom = 10
  frm_numpad.lbl_num.Caption = MLSupport.GSS("frm_collect", "lbl_num3", "Base Name Counter Value")
  frm_numpad.txt_num.Text = frm_collect.numInc_nameCounter.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_numRepacks_DblClick()
  
  unity_main.formfrom = 2
  unity_main.varfrom = 3
  frm_numpad.lbl_num.Caption = frm_collect.lbl_numRepacks.Caption
  frm_numpad.txt_num.Text = frm_collect.numInc_numRepacks.Text
  frm_numpad.Show 1
End Sub

#If ABBFT Then
Private Sub numInc_numSamples_DblClick()
  
  unity_main.formfrom = 2
  unity_main.varfrom = 12
  frm_numpad.lbl_num.Caption = frm_collect.Label12.Caption
  frm_numpad.txt_num.Text = frm_collect.numInc_numSamples.Text
  frm_numpad.Show 1
End Sub
#End If

Private Sub numInc_smplNScans_DblClick()
  
  unity_main.formfrom = 2
  unity_main.varfrom = 1
  frm_numpad.lbl_num.Caption = lbl_smplNScans.Caption
  frm_numpad.txt_num.Text = numInc_smplNScans.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_smplPPT_DblClick()
  
  unity_main.formfrom = 2
  unity_main.varfrom = 2
  frm_numpad.lbl_num.Caption = lbl_numSmplPPT.Caption
  frm_numpad.txt_num.Text = numInc_smplPPT.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_rotateIndexSteps_DblClick()
  
  unity_main.formfrom = 2
  unity_main.varfrom = 6
  frm_numpad.lbl_num.Caption = lbl_rotateIndexSteps.Caption
  frm_numpad.txt_num.Text = numInc_rotateIndexSteps.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_rotateSpeed_DblClick()
  
  unity_main.formfrom = 2
  unity_main.varfrom = 7
  frm_numpad.lbl_num.Caption = lbl_rotateSpeed.Caption
  frm_numpad.txt_num.Text = numInc_rotateSpeed.Text
  frm_numpad.Show 1
End Sub

Private Sub numInc_rotateStepSteps_DblClick()
  
  unity_main.formfrom = 2
  unity_main.varfrom = 8
  frm_numpad.lbl_num.Caption = lbl_rotateStepSteps.Caption
  frm_numpad.txt_num.Text = numInc_rotateStepSteps.Text
  frm_numpad.Show 1
End Sub

#If SSTAR Then
Private Sub opt_adapterType_Click(Index As Integer)
  
  If (m_ignoreEvent = False) Then
    m_ignoreEvent = True
    frm_collect.m_adapterType = MS11AdapterInfo(Index).cfgName
    frm_collect.m_rotateMoveMode = TRM_NONE
    setup_adapter_tray_num
    m_ignoreEvent = False
  End If
End Sub
#End If

#If SSTAR Then
Private Sub optbgexternal_Click()

  If (m_ignoreEvent = False) Then
    build_ref_name_list "external"
    frm_collect.chk_useExtRefTrayCfg.Value = 1
    
    If (frm_collect.combo_extRefFileName.Text <> "") Then
      Call update_prod_ext_ref_wvlns(frm_collect.combo_extRefFileName.Text)
    End If
  End If

  frm_collect.combo_extRefFileName.Visible = True
  frm_collect.combo_olRefFileName.Visible = False
  frm_collect.chk_useExtRefTrayCfg.Visible = True
  frm_collect.Frame5.enabled = False
  frm_collect.opt_backall.enabled = False
  frm_collect.opt_backdemand.Value = True
End Sub
#End If

#If SSTAR Then
Private Sub optbgfile_Click()

  If (m_ignoreEvent = False) Then
    build_ref_name_list "file"
    frm_collect.chk_useExtRefTrayCfg.Value = 0
    
    If (frm_collect.combo_olRefFileName.Text <> "") Then
      Call update_prod_ol_ref_wvlns(frm_collect.combo_olRefFileName.Text)
    End If
  End If
  
  frm_collect.combo_extRefFileName.Visible = False
  frm_collect.combo_olRefFileName.Visible = True
  frm_collect.chk_useExtRefTrayCfg.Visible = False
  frm_collect.Frame5.enabled = False
  frm_collect.opt_backall.enabled = False
  frm_collect.opt_backdemand.Value = True
End Sub
#End If

#If SSTAR Then
Private Sub optbginternal_Click()

  If (m_ignoreEvent = False) Then
    frm_collect.chk_useExtRefTrayCfg.Value = 0
  End If

  frm_collect.combo_extRefFileName.Visible = False
  frm_collect.combo_olRefFileName.Visible = False
  frm_collect.chk_useExtRefTrayCfg.Visible = False
  frm_collect.Frame5.enabled = True
  frm_collect.opt_backall.enabled = True
  
  ' Setup product starting/ending wavelengths to match default
  frm_collect.txt_endWvln.Text = ProdDfltData.endWvln
  frm_collect.txt_startWvln.Text = ProdDfltData.startWvln
  frm_collect.chk_dfltWavelens.Value = 1
End Sub
#End If

#If SSTAR Then
Private Sub optbk_Click()

  optbginternal.Value = True
  optpredno.Value = True
End Sub
#End If

#If SSTAR Then
Private Sub opt_multiCupType_Click(Index As Integer)
  
  If (m_ignoreEvent = False) Then
    m_ignoreEvent = True
    frm_collect.m_multiCupType = MS11MultiCupInfo(Index).cfgName
    frm_collect.m_rotateMoveMode = TRM_NONE
    setup_multi_cup_type
    m_ignoreEvent = False
  End If
End Sub
#End If

Private Sub optnamecounter_Click()
  
  chk_clrManualName.enabled = False
  txt_scanname.Text = Trim(txtsampnamebase.Text) & numInc_nameCounter.Text
  m_sNameMode = 3
End Sub

Private Sub opt_namedate_Click()
  Dim dateStrg As String
  
  chk_clrManualName.enabled = False
  Call rebuild_date(Date, dateStrg)
  txt_scanname.Text = dateStrg & "_" & numInc_dateCounter.Text
  m_sNameMode = 4
End Sub

Private Sub optnamemanual_Click()
  
  chk_clrManualName.enabled = True
  txt_scanname.Text = MLSupport.GSS("frm_collect", "txt_scanname", "Manual Entry")
  m_sNameMode = 1
End Sub

#If SSTAR Then
Private Sub opt_rotateDirCCW_Click()
  
  If (m_ignoreEvent = False) Then
    frm_collect.m_rotateDir = TRD_CCW
  End If
End Sub
#End If

#If SSTAR Then
Private Sub opt_rotateDirCW_Click()
  
  If (m_ignoreEvent = False) Then
    frm_collect.m_rotateDir = TRD_CW
  End If
End Sub
#End If

#If SSTAR Then
Private Sub opt_rotateModeCont_Click()
  
  If (m_ignoreEvent = False) Then
    m_ignoreEvent = True
    frm_collect.m_rotateMoveMode = TRM_CONT
    setup_rotate_move_mode
    m_ignoreEvent = False
  End If
End Sub
#End If

#If SSTAR Then
Private Sub opt_rotateModeIndex_Click()
  
  If (m_ignoreEvent = False) Then
    m_ignoreEvent = True
    frm_collect.m_rotateMoveMode = TRM_INDEX
    setup_rotate_move_mode
    m_ignoreEvent = False
  End If
End Sub
#End If

#If SSTAR Then
Private Sub opt_rotateModeNone_Click()
  
  If (m_ignoreEvent = False) Then
    m_ignoreEvent = True
    frm_collect.m_rotateMoveMode = TRM_NONE
    frm_collect.m_rotateDir = TRD_NONE
    setup_rotate_move_mode
    m_ignoreEvent = False
  End If
End Sub
#End If

#If SSTAR Then
Private Sub opt_rotateModeStep_Click()
  
  If (m_ignoreEvent = False) Then
    m_ignoreEvent = True
    frm_collect.m_rotateMoveMode = TRM_STEP
    setup_rotate_move_mode
    m_ignoreEvent = False
  End If
End Sub
#End If

Private Sub optsavepredsno_Click()
  
  frm_collect.m_savePredictions = False
End Sub

Private Sub optsavepredsyes_Click()
  
  frm_collect.m_savePredictions = True
End Sub

Private Sub tab_frame_TabChanged(Index As Integer)

  show_frame (Index - 1)
End Sub

Private Sub txt_caldir_DblClick(Button As Integer)
  
  unity_main.formfrom = 2
  unity_main.varfrom = 3
  frm_kybd.lbl_kybd.Caption = frm_collect.lbl_caldir.Caption
  frm_kybd.txt_kybd.Text = txt_caldir.Text
  frm_kybd.Show 1
End Sub

Private Sub txt_csvfilename_DblClick(Button As Integer)
  
  unity_main.formfrom = 2
  unity_main.varfrom = 6
  frm_kybd.lbl_kybd.Caption = frm_collect.Label1.Caption
  frm_kybd.txt_kybd.Text = txt_csvfilename.Text
  frm_kybd.Show 1
End Sub

#If ABBFT Then
Private Sub txt_delayMeasure_DblCLick(Button As Integer)

  unity_main.formfrom = 2
  unity_main.varfrom = 14
  frm_numpad.lbl_num.Caption = lbl_delayMeasure.Caption
  frm_numpad.txt_num.Text = txt_delayMeasure.Text
  frm_numpad.Show 1
End Sub
#End If

#If ABBFT Then
Private Sub txt_delayStart_DblCLick(Button As Integer)

  unity_main.formfrom = 2
  unity_main.varfrom = 13
  frm_numpad.lbl_num.Caption = lbl_delayStart.Caption
  frm_numpad.txt_num.Text = txt_delayStart.Text
  frm_numpad.Show 1
End Sub
#End If

Private Sub txt_endWvln_DblCLick(Button As Integer)
  
  unity_main.formfrom = 2
  unity_main.varfrom = 5
  frm_numpad.lbl_num.Caption = lbl_endWvln.Caption & " " & MLSupport.GSS("frm_collect", "lbl_num1", "Wavelength")
  frm_numpad.txt_num.Text = txt_endWvln.Text
  frm_numpad.Show 1
End Sub

Private Sub txt_noOLVal_DblClick(Button As Integer)

  unity_main.formfrom = 2
  unity_main.varfrom = 11
  frm_numpad.lbl_num.Caption = lbl_noOLVal.Caption
  frm_numpad.txt_num.Text = txt_noOLVal.Text
  frm_numpad.Show 1
End Sub

Private Sub txt_noPredVal_DblClick(Button As Integer)

  unity_main.formfrom = 2
  unity_main.varfrom = 7
  frm_kybd.lbl_kybd.Caption = lbl_noPredVal.Caption
  frm_kybd.txt_kybd.Text = txt_noPredVal.Text
  frm_kybd.Show 1
End Sub

Private Sub txtpredfile_DblClick(Button As Integer)
  
  unity_main.formfrom = 2
  unity_main.varfrom = 5
  frm_kybd.lbl_kybd.Caption = frm_collect.Label5.Caption
  frm_kybd.txt_kybd.Text = txtpredfile.Text
  frm_kybd.Show 1
End Sub

Private Sub txtsampnamebase_Change()
  
  If (m_sNameMode = 3) Then
    txt_scanname.Text = Trim(txtsampnamebase.Text) & numInc_nameCounter.Text
  End If
End Sub

Private Sub txtsampnamebase_DblClick(Button As Integer)
  
  unity_main.formfrom = 2
  unity_main.varfrom = 1
  frm_kybd.lbl_kybd.Caption = frm_collect.Label4.Caption
  frm_kybd.txt_kybd.Text = txtsampnamebase.Text
  frm_kybd.Show 1
End Sub

Private Sub txt_startWvln_DblCLick(Button As Integer)
  
  unity_main.formfrom = 2
  unity_main.varfrom = 4
  frm_numpad.lbl_num.Caption = lbl_startWvln.Caption & " " & MLSupport.GSS("frm_collect", "lbl_num1", "Wavelength")
  frm_numpad.txt_num.Text = txt_startWvln.Text
  frm_numpad.Show 1
End Sub

#If ABBFT Then
Private Sub txt_endWavenumIndx_Change()
  Dim wavenumIndx As Long
  Dim wavenum As Double
  
  If (txt_endWavenumIndx.Text = "") Or (IsNumeric(txt_endWavenumIndx.Text) = False) Then
    txt_endWavenumIndx.Text = 0
    txt_endWavenum.Text = 0
  Else
    wavenumIndx = txt_endWavenumIndx.Text
  
    If (calc_wavenum(wavenumIndx, wavenum) = False) Then
      txt_endWavenumIndx.Text = wavenumIndx
    End If
  
    txt_endWavenum.Text = wavenum
  End If
End Sub
#End If

#If ABBFT Then
Private Sub txt_endWavenumIndx_DblCLick(Button As Integer)
  
  unity_main.formfrom = 2
  unity_main.varfrom = 16
  frm_numpad.lbl_num.Caption = lbl_endWavenum.Caption & " " & MLSupport.GSS("frm_collect", "lbl_num4", "Wavenumber")
  frm_numpad.txt_num.Text = txt_endWavenumIndx.Text
  frm_numpad.Show 1
End Sub
#End If

#If ABBFT Then
Private Sub txt_startWavenumIndx_Change()
  Dim wavenumIndx As Long
  Dim wavenum As Double
  
  If (txt_startWavenumIndx.Text = "") Or (IsNumeric(txt_startWavenumIndx) = False) Then
    txt_startWavenumIndx.Text = 0
    txt_startWavenum.Text = 0
  Else
    wavenumIndx = txt_startWavenumIndx.Text
  
    If (calc_wavenum(wavenumIndx, wavenum) = False) Then
      txt_startWavenumIndx.Text = wavenumIndx
    End If
  
    txt_startWavenum.Text = wavenum
  End If
End Sub
#End If

#If ABBFT Then
Private Sub txt_startWavenumIndx_DblCLick(Button As Integer)
  
  unity_main.formfrom = 2
  unity_main.varfrom = 15
  frm_numpad.lbl_num.Caption = lbl_startWavenum.Caption & " " & MLSupport.GSS("frm_collect", "lbl_num4", "Wavenumber")
  frm_numpad.txt_num.Text = txt_startWavenumIndx.Text
  frm_numpad.Show 1
End Sub
#End If








