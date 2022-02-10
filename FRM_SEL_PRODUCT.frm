VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form FRM_SEL_PRODUCT 
   Caption         =   "Product Selection"
   ClientHeight    =   9840
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10635
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
   Icon            =   "FRM_SEL_PRODUCT.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9840
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   6480
      TabIndex        =   0
      Top             =   8520
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
      Caption         =   "FRM_SEL_PRODUCT.frx":0442
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "FRM_SEL_PRODUCT.frx":046E
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "FRM_SEL_PRODUCT.frx":048E
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9240
      Top             =   360
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9840
      FormDesignWidth =   10635
   End
   Begin HexUniControls.ctlUniListBoxXP LSTPRODUCTS 
      Height          =   7260
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   9255
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TrapTab         =   0   'False
      Tip             =   "FRM_SEL_PRODUCT.frx":04AA
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   -1  'True
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "FRM_SEL_PRODUCT.frx":04CA
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_select 
      Default         =   -1  'True
      Height          =   650
      Left            =   2280
      TabIndex        =   2
      Top             =   8520
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
      Caption         =   "FRM_SEL_PRODUCT.frx":04E6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "FRM_SEL_PRODUCT.frx":0512
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "FRM_SEL_PRODUCT.frx":0532
   End
   Begin HexUniControls.ctlUniListBoxXP LST_INIFILE 
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   45
      Visible         =   0   'False
      Width           =   180
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
      Tip             =   "FRM_SEL_PRODUCT.frx":054E
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "FRM_SEL_PRODUCT.frx":056E
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniListBoxXP LST_SAMPMODE 
      Height          =   255
      Left            =   6720
      TabIndex        =   6
      Top             =   45
      Visible         =   0   'False
      Width           =   180
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
      Tip             =   "FRM_SEL_PRODUCT.frx":058A
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "FRM_SEL_PRODUCT.frx":05AA
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_sampmode 
      Height          =   360
      Left            =   7800
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   635
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   0   'False
      Locked          =   0   'False
      Text            =   "FRM_SEL_PRODUCT.frx":05C6
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
      Tip             =   "FRM_SEL_PRODUCT.frx":05E6
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "FRM_SEL_PRODUCT.frx":0606
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_inifile 
      Height          =   360
      Left            =   8160
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   635
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   0   'False
      Locked          =   0   'False
      Text            =   "FRM_SEL_PRODUCT.frx":0622
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
      Tip             =   "FRM_SEL_PRODUCT.frx":0642
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "FRM_SEL_PRODUCT.frx":0662
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_product 
      Height          =   360
      Left            =   7440
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   635
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   0   'False
      Locked          =   0   'False
      Text            =   "FRM_SEL_PRODUCT.frx":067E
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
      Tip             =   "FRM_SEL_PRODUCT.frx":069E
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "FRM_SEL_PRODUCT.frx":06BE
   End
   Begin HexUniControls.ctlUniImage ctlUniImage1 
      Height          =   1035
      Left            =   4470
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1826
      Picture         =   "FRM_SEL_PRODUCT.frx":06DA
      Tip             =   "FRM_SEL_PRODUCT.frx":55E1
      Enabled         =   -1  'True
      Border          =   0   'False
      BackColor       =   -2147483633
      BorderColor     =   -1
      RoundedBorders  =   -1  'True
      Stretch         =   -1  'True
      XTransp         =   0
      YTransp         =   0
      MousePointer    =   0
      MouseIcon       =   "FRM_SEL_PRODUCT.frx":5601
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   7440
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
      Left            =   8400
      Top             =   360
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "FRM_SEL_PRODUCT.frx":561D
   End
End
Attribute VB_Name = "FRM_SEL_PRODUCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_select_Click()

  selprod
End Sub

Sub product_clicked()
  Dim iii As Integer
  
  iii = Int(FRM_SEL_PRODUCT.LSTPRODUCTS.ListIndex)
  txt_product.Text = FRM_SEL_PRODUCT.LSTPRODUCTS.List(iii)
  txt_sampmode.Text = FRM_SEL_PRODUCT.LST_SAMPMODE.List(iii)
  txt_inifile.Text = FRM_SEL_PRODUCT.LST_INIFILE.List(iii)
End Sub

Sub selprod()
  
  If Trim(FRM_SEL_PRODUCT.txt_inifile.Text) = "" Then
    CWrap.ShowMessageBoxW MLSupport.GSS("FRM_SEL_PRODUCT", "errMsg1", "You must select a product from the list"), vbExclamation
    Exit Sub
  End If
  
  Call frm_collect.savescansettings(False)
  Call frm_Inst.savemyinsts(False, False)
  Call frm_dynRptCfg.save_cfg(False, False)
  
  If (unity_main.load_prod_file(Trim(FRM_SEL_PRODUCT.txt_inifile.Text), True) = True) Then
    Call unity_main.save_last_product(Trim(FRM_SEL_PRODUCT.txt_inifile.Text))
    FRM_SEL_PRODUCT.txt_inifile.Text = ""
    frm_labData.m_lastSVFDir = ""
    
    frm_scanname.lbl_prod.Caption = unity_main.lblProd1.Caption
    frmedmod.fixthesize
  
    unity_main.errorstring = ("User product selection: " & unity_main.lblProd1.Caption)
    unity_main.write_error
    
    If (unity_main.cmd_start.Visible = True) Then
      unity_main.cmd_start.Visible = False
      unity_main.cmd_stop.Visible = True
    End If
    
    Call unity_main.restart_loop(LOG_DBG_LEVEL3, "Product Selection screen Select button selected")
    FRM_SEL_PRODUCT.Visible = False
  End If
End Sub

Private Sub List1_Click()
  Dim iii As Integer
  
  iii = Int(FRM_SEL_PRODUCT.LSTPRODUCTS.ListIndex)
  txt_product.Text = FRM_SEL_PRODUCT.LSTPRODUCTS.List(iii)
  txt_sampmode.Text = FRM_SEL_PRODUCT.LST_SAMPMODE.List(iii)
  txt_inifile.Text = FRM_SEL_PRODUCT.LST_INIFILE.List(iii)
End Sub

Private Sub cmd_cancel_Click()
  Dim msg As String
  
  msg = "Product Selection screen Cancel button selected"
  Call unity_main.restart_loop(LOG_DBG_LEVEL3, msg)
  FRM_SEL_PRODUCT.Visible = False
End Sub

Private Sub Form_Load()
  
  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
  Call loadproducts
End Sub

Sub loadproducts()
  Dim fname As String
  Dim lineCnt As Integer
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim rc As Boolean
  Dim inString As String
  Dim varStr As Variant
  Dim errMsg As String
  Dim uniMsg As String

  FRM_SEL_PRODUCT.LSTPRODUCTS.Clear
  FRM_SEL_PRODUCT.LST_SAMPMODE.Clear
  FRM_SEL_PRODUCT.LST_INIFILE.Clear
  
  fname = (CFG_DIR & PRODUCTS_CFG_FILE)
  
  If (uniFile.OpenFileRead(fname) = True) Then
    On Error GoTo BAD_FILE
    fEncoding = uniFile.ReadBOM
    
    While Not (uniFile.EOF())
      lineCnt = lineCnt + 1
      
      If (fEncoding = fe_ANSI) Then
        rc = uniFile.ReadAnsiLine(inString)
      Else
        rc = uniFile.ReadUnicodeLine(inString)
      End If
      
      If (rc = False) Then GoTo BAD_FILE

      varStr = Split(inString, ",")
      FRM_SEL_PRODUCT.LSTPRODUCTS.AddItem Trim(varStr(0))
      FRM_SEL_PRODUCT.LST_SAMPMODE.AddItem Trim(varStr(1))
      FRM_SEL_PRODUCT.LST_INIFILE.AddItem Trim(varStr(2))
    Wend
  Else
BAD_FILE:
   If (lineCnt = 0) Then
      errMsg = (fname & " file read error")
      uniMsg = MLSupport.GGS_Params("fileErrMsg4", "%1 file read error. %2", fname, Error$)
    Else
      errMsg = (fname & " file has error on line " & CStr(lineCnt))
      uniMsg = MLSupport.GGS_Params("fileErrMsg3", "%1 file has error on line %2. %3", fname, CStr(lineCnt), Error$)
    End If
  
    errMsg = (errMsg & ". " & Error$)
    unity_main.errorstring = errMsg
    unity_main.write_error
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  End If
  
  uniFile.CloseFile
End Sub

Private Sub LSTPRODUCTS_Click()
  
  Call FRM_SEL_PRODUCT.product_clicked
End Sub

Private Sub LSTPRODUCTS_DblClick()
  
  Call product_clicked
  Call selprod
End Sub








