VERSION 5.00
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_sendlims 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LIMS Out"
   ClientHeight    =   3870
   ClientLeft      =   675
   ClientTop       =   2235
   ClientWidth     =   2910
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
   Icon            =   "frm_sendlims.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   2400
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_sendlims.frx":0442
      BackColor       =   255
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_sendlims.frx":046E
      Style           =   1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_sendlims.frx":048E
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_ok 
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2400
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_sendlims.frx":04AA
      BackColor       =   49152
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_sendlims.frx":04CE
      Style           =   1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_sendlims.frx":04EE
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   735
      Left            =   360
      Top             =   120
      Width           =   2175
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_sendlims.frx":050A
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_sendlims.frx":054E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_sendlims.frx":056E
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   2640
      Top             =   360
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
      Left            =   2520
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_sendlims.frx":058A
   End
End
Attribute VB_Name = "frm_sendlims"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_cancel_Click()
  
  unity_main.errorstring = "LIMS Out screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  unity_main.m_acceptLims = False
  Unload frm_sendlims
End Sub

Private Sub cmd_ok_Click()
  
  unity_main.errorstring = "LIMS Out screen OK button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  unity_main.m_acceptLims = True
  Unload frm_sendlims
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub








