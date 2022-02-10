VERSION 5.00
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frmParse2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Math Expression Parser"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "frm_parse2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniTextBoxXP Text2 
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   3000
      Width           =   1575
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_parse2.frx":08CA
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
      Tip             =   "frm_parse2.frx":08F4
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_parse2.frx":0914
   End
   Begin HexUniControls.ctlUniButtonImageXP Command1 
      Height          =   855
      Left            =   5280
      TabIndex        =   1
      Top             =   6000
      Width           =   1455
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "frm_parse2.frx":0930
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_parse2.frx":0960
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_parse2.frx":0980
   End
   Begin HexUniControls.ctlUniTextBoxXP Text1 
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   5280
      Width           =   4095
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_parse2.frx":099C
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
      Tip             =   "frm_parse2.frx":09C6
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_parse2.frx":09E6
   End
   Begin HexUniControls.ctlUniFrameXP Frame3 
      Height          =   1455
      Left            =   120
      Top             =   120
      Width           =   4575
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_parse2.frx":0A02
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_parse2.frx":0A32
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_parse2.frx":0A52
      Begin HexUniControls.ctlUniTextBoxXP txtExpression 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2895
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frm_parse2.frx":0A6E
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
         Tip             =   "frm_parse2.frx":0A94
         NoHideSel       =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_parse2.frx":0AB4
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdParse 
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         Top             =   480
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "frm_parse2.frx":0AD0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_parse2.frx":0AFA
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_parse2.frx":0B1A
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdMultiParse 
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   960
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "frm_parse2.frx":0B36
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_parse2.frx":0B6E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_parse2.frx":0B8E
      End
      Begin HexUniControls.ctlUniTextBoxXP txtNumParses 
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   960
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frm_parse2.frx":0BAA
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
         Tip             =   "frm_parse2.frx":0BD2
         NoHideSel       =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_parse2.frx":0BF2
      End
      Begin HexUniControls.ctlUniLabel Label1 
         Height          =   255
         Left            =   120
         Top             =   240
         Width           =   1575
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_parse2.frx":0C0E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_parse2.frx":0C44
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_parse2.frx":0C64
      End
      Begin HexUniControls.ctlUniLabel Label2 
         Height          =   255
         Left            =   120
         Top             =   1000
         Width           =   1695
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_parse2.frx":0C80
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_parse2.frx":0CBE
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_parse2.frx":0CDE
      End
   End
   Begin HexUniControls.ctlUniFrameXP Frame2 
      Height          =   2295
      Left            =   120
      Top             =   1680
      Width           =   4575
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_parse2.frx":0CFA
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_parse2.frx":0D30
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_parse2.frx":0D50
      Begin HexUniControls.ctlUniButtonImageXP cmdAddConst 
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   480
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "frm_parse2.frx":0D6C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_parse2.frx":0DA4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_parse2.frx":0DC4
      End
      Begin HexUniControls.ctlUniTextBoxXP txtConstName 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frm_parse2.frx":0DE0
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
         Tip             =   "frm_parse2.frx":0E00
         NoHideSel       =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_parse2.frx":0E20
      End
      Begin HexUniControls.ctlUniTextBoxXP txtConstValue 
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   480
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frm_parse2.frx":0E3C
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
         Tip             =   "frm_parse2.frx":0E5C
         NoHideSel       =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_parse2.frx":0E7C
      End
      Begin HexUniControls.ctlUniListBoxXP lstConstants 
         Height          =   1020
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2895
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
         Tip             =   "frm_parse2.frx":0E98
         MultiSelect     =   0
         Sorted          =   0   'False
         HScroll         =   0   'False
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
         RoundedBorders  =   0   'False
         SelectorStyle   =   -1
         MousePointer    =   0
         MouseIcon       =   "frm_parse2.frx":0EB8
         ManualStart     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRemoveConst 
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   960
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "frm_parse2.frx":0ED4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frm_parse2.frx":0F0C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frm_parse2.frx":0F2C
      End
      Begin HexUniControls.ctlUniLabel Label3 
         Height          =   255
         Left            =   120
         Top             =   240
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_parse2.frx":0F48
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_parse2.frx":0F84
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_parse2.frx":0FA4
      End
      Begin HexUniControls.ctlUniLabel Label4 
         Height          =   255
         Left            =   2160
         Top             =   240
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_parse2.frx":0FC0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_parse2.frx":0FEC
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_parse2.frx":100C
      End
   End
   Begin HexUniControls.ctlUniFrameXP Frame1 
      Height          =   855
      Left            =   120
      Top             =   4080
      Width           =   4575
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_parse2.frx":1028
      Enabled         =   -1  'True
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Tip             =   "frm_parse2.frx":105E
      VistaStyle      =   -1  'True
      UseShadow       =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_parse2.frx":107E
      Begin HexUniControls.ctlUniLabel Label5 
         Height          =   495
         Left            =   120
         Top             =   310
         Width           =   3975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frm_parse2.frx":109A
         BackColor       =   -2147483633
         ForeColor       =   0
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frm_parse2.frx":1158
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frm_parse2.frx":1178
      End
   End
End
Attribute VB_Name = "frmParse2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyParser As New clsExpressionParser

' Keep track of user-defined constants
Dim ConstNames As New Collection

Private Sub cmdAddConst_Click()
On Error GoTo cmdAddConst_ErrHandler
    
    ' Validity checks
    If Trim(txtConstName) = "" Then
        CWrap.ShowMessageBoxW MLSupport.GSS("frmParse2", "errMsg1", "A valid constant name must begin with a letter. Additional letters may include the alphanumeric letters or underscores. For Example: MyConst_2"), vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(txtConstValue) Then
        CWrap.ShowMessageBoxW MLSupport.GSS("frmParse2", "errMsg2", "Please enter a valid number"), vbExclamation
        Exit Sub
    End If
    
    MyParser.AddConstant txtConstName, CDbl(txtConstValue)
    
    lstConstants.AddItem txtConstName & " - " & txtConstValue
    ConstNames.Add txtConstName.Text
    
    Exit Sub

cmdAddConst_ErrHandler:
   CWrap.ShowMessageBoxW MLSupport.GGS_Params("frmParse2.errMsg1", "Parse error: %1", err.Description), vbCritical
End Sub

Private Sub cmdParse_Click()
On Error GoTo cmdParse_ErrHandler

Dim Result As Double

    Result = MyParser.ParseExpression(txtExpression.Text)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("frmParse2.statMsg1", "Result: %1", Format(Result, "#0.0#####"))
    Exit Sub
    
cmdParse_ErrHandler:
     If err.Number >= PERR_FIRST And _
       err.Number <= PERR_LAST Then
        ShowParseError
    Else
        CWrap.ShowMessageBoxW MLSupport.GGS_Params("frmParse2.errMsg1", "Parse error: %1", err.Description), vbCritical
    End If
End Sub

Private Sub cmdMultiParse_Click()
  Dim Value As Integer
  Dim i As Long
  Dim Expression As String
  Dim StartTime As Single, EndTime As Single
  Dim NumParses As Long

    On Error GoTo cmdMultiParse_ErrHandler
    
    If Not IsNumeric(txtNumParses) Or Val(txtNumParses) < 1 Then
        CWrap.ShowMessageBoxW MLSupport.GSS("frmParse2", "errMsg3", "Please enter a positive number"), vbExclamation
        Exit Sub
    End If
    
    NumParses = CLng(txtNumParses)
    Expression = txtExpression.Text
    
    StartTime = Timer
    For i = 1 To NumParses
        Value = MyParser.ParseExpression(Expression)
    Next
    EndTime = Timer

    CWrap.ShowMessageBoxW MLSupport.GGS_Params("frmParse2.statMsg2", "Parse time took %1 seconds", Format(CStr(EndTime - StartTime), "#0.0##")), vbInformation
    Exit Sub
    
cmdMultiParse_ErrHandler:
    If (err.Number >= PERR_FIRST) And (err.Number <= PERR_LAST) Then
      ShowParseError
    Else
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("frmParse2.errMsg1", "Parse error: %1", err.Description), vbCritical
    End If
End Sub

Private Sub ShowParseError()
  
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("frmParse2.errMsg2", "Parse error no. %1, %2, Raised from: %3", CStr(err.Number - PERR_FIRST + 1), err.Description, err.Source), vbCritical
  txtExpression.SelStart = MyParser.LastErrorPosition - 1
  txtExpression.SelLength = 1
  txtExpression.SetFocus
End Sub

Private Sub cmdRemoveConst_Click()
    
    On Error GoTo cmdRemoveConst_ErrHandler

    If lstConstants.ListIndex = -1 Then
        CWrap.ShowMessageBoxW MLSupport.GSS("frmParse2", "errMsg4", "Select a constant to remove"), vbExclamation
        Exit Sub
    End If

    MyParser.RemoveConstant ConstNames(lstConstants.ListIndex + 1)
    
    ConstNames.Remove lstConstants.ListIndex + 1
    lstConstants.RemoveItem lstConstants.ListIndex
    Exit Sub
    
cmdRemoveConst_ErrHandler:
   CWrap.ShowMessageBoxW MLSupport.GGS_Params("frmParse2.errMsg1", "Parse error: %1", err.Description), vbCritical
End Sub

Private Sub Command1_Click()
txtExpression.Text = Trim(Text1.Text)
Call runparse
End Sub

Sub runparse()
  Dim Result As Double
  
  On Error GoTo cmdParse_ErrHandler
  Result = MyParser.ParseExpression(txtExpression.Text)
  Text2.Text = Format(Result, "#0.0#####")
  unity_main.sec_value = Format(Result, "#0.0#####")
  Exit Sub
    
cmdParse_ErrHandler:
  If (err.Number >= PERR_FIRST) And (err.Number <= PERR_LAST) Then
    ShowParseError
  Else
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("frmParse2.errMsg1", "Parse error: %1", err.Description), vbCritical
  End If
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub








