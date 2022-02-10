VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form Frm_prior_preds 
   Caption         =   "Product Analyses (Report File)"
   ClientHeight    =   7530
   ClientLeft      =   360
   ClientTop       =   645
   ClientWidth     =   8790
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
   Icon            =   "Frm_prior_preds.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   840
      Top             =   6480
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7530
      FormDesignWidth =   8790
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_delete 
      Height          =   650
      Left            =   1595
      TabIndex        =   1
      Top             =   6360
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
      Caption         =   "Frm_prior_preds.frx":0442
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "Frm_prior_preds.frx":0478
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "Frm_prior_preds.frx":0498
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_exit 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   5195
      TabIndex        =   0
      Top             =   6360
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
      Caption         =   "Frm_prior_preds.frx":04B4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "Frm_prior_preds.frx":04DC
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "Frm_prior_preds.frx":04FC
   End
   Begin FPUSpreadADO.fpSpread ss_prior_preds 
      Height          =   5775
      Left            =   360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   8055
      _Version        =   458752
      _ExtentX        =   14208
      _ExtentY        =   10186
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   -2147483633
      MaxCols         =   1
      MaxRows         =   365000
      OperationMode   =   1
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   0
      SpreadDesigner  =   "Frm_prior_preds.frx":0518
      UserResize      =   1
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   240
      Top             =   6840
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
      Left            =   120
      Top             =   6480
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "Frm_prior_preds.frx":078C
   End
End
Attribute VB_Name = "Frm_prior_preds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_exit_Click()
  
  unity_main.errorstring = "Product Analyses (Report File) screen Exit button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Unload Frm_prior_preds
End Sub

Private Sub cmd_delete_Click()
  Dim ftokill As String
  Dim optVal As Integer
  Dim errMsg As String
  Dim uniMsg As String

  unity_main.errorstring = "Product Analyses (Report File) screen Delete File button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  On Error GoTo NO_DELETE
  ftokill = unity_main.m_savePredFile
  uniMsg = MLSupport.GGS_Params("fileMsg1", "Are you sure you want to delete %1?", ftokill)
  optVal = CWrap.ShowMessageBoxW(uniMsg, vbYesNo)
  
  If (optVal = vbYes) Then
    CFile.st_RmFile ftokill
    unity_main.errorstring = ("User cleared report file: " & ftokill)
    unity_main.write_error (LOG_DBG_LEVEL1)
    Unload Frm_prior_preds
  End If
  
  Exit Sub
  
NO_DELETE:
  errMsg = (ftokill & " file delete error. " & Error$)
  unity_main.errorstring = errMsg
  unity_main.write_error (LOG_DBG_LEVEL1)
  uniMsg = MLSupport.GGS_Params("fileErrMsg8", "%1 file delete error. %2", ftokill, Error$)
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub








