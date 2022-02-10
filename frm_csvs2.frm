VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_csvs2 
   Caption         =   "Product Analyses (CSV File)"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
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
   ScaleHeight     =   8070
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   Begin FPUSpreadADO.fpSpread ss1 
      Height          =   6255
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   10455
      _Version        =   458752
      _ExtentX        =   18441
      _ExtentY        =   11033
      _StockProps     =   64
      ColHeaderDisplay=   0
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
      MaxCols         =   105
      MaxRows         =   5000
      OperationMode   =   1
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   0
      SpreadDesigner  =   "frm_csvs2.frx":0000
      UserResize      =   1
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   240
      Top             =   7080
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   8070
      FormDesignWidth =   11010
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_exit 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   6480
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
      Caption         =   "frm_csvs2.frx":0274
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_csvs2.frx":029C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_csvs2.frx":02BC
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_delete 
      Height          =   650
      Left            =   2400
      TabIndex        =   1
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
      Caption         =   "frm_csvs2.frx":02D8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_csvs2.frx":030E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_csvs2.frx":032E
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   240
      Top             =   7440
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
      Left            =   840
      Top             =   7080
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_csvs2.frx":034A
   End
End
Attribute VB_Name = "frm_csvs2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_exit_Click()
  
  ' Restart auto operations if no delete button
  If (frm_csvs2.cmd_delete.Visible = False) Then
    Call unity_main.restart_loop(LOG_DBG_LEVEL3, "Product Analyses (CSV File) screen Exit button selected")
  Else
    unity_main.errorstring = "Product Analyses (CSV File) screen Exit button selected"
    unity_main.write_error (LOG_DBG_LEVEL3)
  End If

  Unload frm_csvs2
End Sub

Private Sub cmd_delete_Click()
  Dim ftokill As String
  Dim optVal As Integer
  Dim errMsg As String
  Dim uniMsg As String
  
  unity_main.errorstring = "Product Analyses (CSV File) screen Delete File button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  On Error GoTo NO_DELETE
  ftokill = unity_main.m_saveCSVFile
  uniMsg = MLSupport.GGS_Params("fileMsg1", "Are you sure you want to delete %1?", unity_main.m_saveCSVFile)
  optVal = CWrap.ShowMessageBoxW(uniMsg, vbYesNo)
  
  If (optVal = vbYes) Then
    CFile.st_RmFile ftokill
    unity_main.errorstring = ("User cleared csv file: " & ftokill)
    unity_main.write_error (LOG_DBG_LEVEL1)
    Unload frm_csvs2
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








