VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_buildprods 
   Caption         =   "Add Existing Product"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10815
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
   Icon            =   "frm_buildprods.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniButtonImageXP cmd_add 
      Height          =   650
      Left            =   4440
      TabIndex        =   3
      Top             =   1440
      Width           =   2000
      _ExtentX        =   3519
      _ExtentY        =   1138
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
      Caption         =   "frm_buildprods.frx":0442
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_buildprods.frx":0468
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_buildprods.frx":0488
   End
   Begin HexUniControls.ctlUniListBoxExXP lst_temp 
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      IconDim         =   16
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
      Tip             =   "frm_buildprods.frx":04A4
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   0   'False
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_buildprods.frx":04C4
   End
   Begin HexUniControls.ctlUniListBoxExXP lst_includeini 
      Height          =   3015
      Left            =   6600
      TabIndex        =   5
      Top             =   480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5318
      IconDim         =   16
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
      Tip             =   "frm_buildprods.frx":04E0
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   -1  'True
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_buildprods.frx":0500
   End
   Begin HexUniControls.ctlUniListBoxExXP lst_inis 
      Height          =   3015
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5318
      IconDim         =   16
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
      Tip             =   "frm_buildprods.frx":051C
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   -1  'True
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_buildprods.frx":053C
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   8040
      TabIndex        =   0
      Top             =   5040
      Width           =   2000
      _ExtentX        =   3519
      _ExtentY        =   1138
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
      Caption         =   "frm_buildprods.frx":0558
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_buildprods.frx":0584
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_buildprods.frx":05A4
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   650
      Left            =   5040
      TabIndex        =   1
      Top             =   5040
      Width           =   2000
      _ExtentX        =   3519
      _ExtentY        =   1138
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
      Caption         =   "frm_buildprods.frx":05C0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_buildprods.frx":05F8
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_buildprods.frx":0618
   End
   Begin HexUniControls.ctlUniLabel lbl_samp 
      Height          =   375
      Left            =   4920
      Top             =   4320
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_buildprods.frx":0634
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_buildprods.frx":0654
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_buildprods.frx":0674
   End
   Begin HexUniControls.ctlUniLabel lbl_prod 
      Height          =   375
      Left            =   4920
      Top             =   3720
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_buildprods.frx":0690
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_buildprods.frx":06B0
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_buildprods.frx":06D0
   End
   Begin HexUniControls.ctlUniLabel label2 
      Height          =   375
      Left            =   3240
      Top             =   4320
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
      Caption         =   "frm_buildprods.frx":06EC
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_buildprods.frx":071A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_buildprods.frx":073A
   End
   Begin HexUniControls.ctlUniLabel label1 
      Height          =   375
      Left            =   3240
      Top             =   3720
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
      Caption         =   "frm_buildprods.frx":0756
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_buildprods.frx":0784
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_buildprods.frx":07A4
   End
   Begin HexUniControls.ctlUniLabel label5 
      Height          =   1935
      Left            =   600
      Top             =   3720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3413
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_buildprods.frx":07C0
      BackColor       =   -2147483633
      ForeColor       =   0
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_buildprods.frx":0920
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_buildprods.frx":0940
   End
   Begin HexUniControls.ctlUniLabel label4 
      Height          =   255
      Left            =   6600
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_buildprods.frx":095C
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_buildprods.frx":09AE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_buildprods.frx":09CE
   End
   Begin HexUniControls.ctlUniLabel label3 
      Height          =   255
      Left            =   360
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_buildprods.frx":09EA
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_buildprods.frx":0A40
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_buildprods.frx":0A60
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4680
      Top             =   120
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6060
      FormDesignWidth =   10815
   End
   Begin HexUniControls.ctlUniFileBoxXP File1 
      Height          =   1260
      Left            =   0
      TabIndex        =   2
      Top             =   5760
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   0
      _ExtentY        =   0
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
      Tip             =   "frm_buildprods.frx":0A7C
      Path            =   ""
      Pattern         =   "*.*"
      ReadOnly        =   -1  'True
      System          =   0   'False
      Hidden          =   0   'False
      PermitNavigation=   -1  'True
      MultiSelect     =   0
      HScroll         =   -1  'True
      ShowFullPath    =   0   'False
      DisplayMode     =   1
      MousePointer    =   0
      MouseIcon       =   "frm_buildprods.frx":0A9C
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   5880
      Top             =   120
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
      Left            =   5280
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_buildprods.frx":0AB8
   End
End
Attribute VB_Name = "frm_buildprods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub getheader()
  Dim prodFile As String
  Dim tmpFile As String
  Dim tempstring As String
  Dim varStr As Variant
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim rc As Boolean
  Dim errMsg As String
  Dim uniMsg As String
  
  If (frm_buildprods.lst_inis.ListIndex < 0) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_buildprods", "errMsg1", "You must select a product file to load"), vbOKOnly
  End If

  On Error GoTo FILE_ERROR
  prodFile = (PRODUCTS_CFG_DIR & frm_buildprods.lst_inis.Text)
  tmpFile = (PRODUCTS_CFG_DIR & TMP_ADD_PROD_CFG_FILE)
  uniFile.st_CopyFile prodFile, tmpFile
                                                                                                                                                                         
  If (uniFile.OpenFileRead(tmpFile) = True) Then
    fEncoding = uniFile.ReadBOM

    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(tempstring)
      rc = rc And uniFile.ReadAnsiLine(tempstring)
    Else
      rc = uniFile.ReadUnicodeLine(tempstring)
      rc = rc And uniFile.ReadUnicodeLine(tempstring)
    End If
      
    If (rc = False) Then GoTo FILE_ERROR
                                                                                                                                                                         
    varStr = Split(tempstring, ",")
  
    frm_buildprods.lbl_prod.Caption = varStr(0)
    frm_buildprods.lbl_samp.Caption = varStr(1)
  Else
FILE_ERROR:
    errMsg = (prodFile & " file open/read error. " & Error$)
    unity_main.errorstring = errMsg
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg1", "%1 file open/read error. %2", prodFile, Error$)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  End If
  
  uniFile.CloseFile

  If (tmpFile <> "") And (uniFile.st_FileExist(tmpFile) = True) Then
    uniFile.st_RmFile tmpFile
  End If
End Sub

Sub addit()
  Dim prodFile As String
  Dim tmpFile As String
  Dim tempstring As String
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim rc As Boolean
  Dim errMsg As String
  Dim uniMsg As String

  If (frm_buildprods.lst_inis.ListIndex < 0) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_buildprods", "errMsg1", "You must select a product file to load"), vbOKOnly
  End If

  On Error GoTo FILE_ERROR
  prodFile = (PRODUCTS_CFG_DIR & frm_buildprods.lst_inis.Text)
  tmpFile = (PRODUCTS_CFG_DIR & TMP_ADD_PROD_CFG_FILE)
  uniFile.st_CopyFile prodFile, tmpFile
  
  If (uniFile.OpenFileRead(tmpFile) = True) Then
    fEncoding = uniFile.ReadBOM

    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(tempstring)
      rc = rc And uniFile.ReadAnsiLine(tempstring)
    Else
      rc = uniFile.ReadUnicodeLine(tempstring)
      rc = rc And uniFile.ReadUnicodeLine(tempstring)
    End If
      
    If (rc = False) Then GoTo FILE_ERROR
  
    frm_buildprods.lst_includeini.AddItem (tempstring & "," & frm_buildprods.lst_inis.Text)
  Else
FILE_ERROR:
    errMsg = (prodFile & " file open/read error. " & Error$)
    unity_main.errorstring = errMsg
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg1", "%1 file open/read error. %2", prodFile, Error$)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  End If
  
  uniFile.CloseFile

  If (tmpFile <> "") And (uniFile.st_FileExist(tmpFile) = True) Then
    uniFile.st_RmFile tmpFile
  End If
End Sub

Private Sub cmd_add_Click()
  
  unity_main.errorstring = "Add Existing Product screen Add button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  frm_buildprods.addit
  frm_buildprods.getheader
End Sub

Private Sub cmd_cancel_Click()
  
  unity_main.errorstring = "Add Existing Product screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  frmProduct.txt_inifile.Text = ""
  frmProduct.txt_product.Text = ""
  frmProduct.txt_sampmode.Text = ""
  frmProduct.LSTPRODUCTS.ListIndex = -1
  Unload frm_buildprods
End Sub

Private Sub cmd_save_Click()
  
  unity_main.errorstring = "Add Existing Product screen Save Changes button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  frm_buildprods.save_added_prod
  frmProduct.loadprodini
  
  frmProduct.txt_inifile.Text = ""
  frmProduct.txt_product.Text = ""
  frmProduct.txt_sampmode.Text = ""
  frmProduct.LSTPRODUCTS.ListIndex = -1
End Sub

Sub save_added_prod()
  Dim fileName As String
  Dim zz, nn As Integer
  Dim pos As Integer
  Dim tmpFile As String
  Dim inString As String
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim rc As Boolean
  Dim errMsg As String
  Dim uniMsg As String

  If (frm_buildprods.lst_includeini.ListCount = 0) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_buildprods", "errMsg2", "You must select and add product file first"), vbOKOnly
    Exit Sub
  End If
  
  frm_buildprods.lst_temp.Clear

  On Error GoTo FILE_ERROR1
  fileName = (CFG_DIR & PRODUCTS_CFG_FILE)
  tmpFile = (CFG_DIR & TMP_PRODUCTS_CFG_FILE)
  uniFile.st_CopyFile fileName, tmpFile
  
  If (uniFile.OpenFileRead(tmpFile) = False) Then GoTo FILE_ERROR1
  
  fEncoding = uniFile.ReadBOM

  ' Find first model info line in .ini file
  While Not (uniFile.EOF())
    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(inString)
    Else
      rc = uniFile.ReadUnicodeLine(inString)
    End If
      
    If (rc = False) Then GoTo FILE_ERROR1
    
    If (Trim(inString) <> "") Then
      frm_buildprods.lst_temp.AddItem (inString)
    End If
  Wend

  uniFile.CloseFile

  For zz = 0 To (frm_buildprods.lst_includeini.ListCount - 1)
    For nn = 0 To (frm_buildprods.lst_temp.ListCount - 1)
      ' Check if product is already in list
      If ((frm_buildprods.lst_includeini.List(zz)) = frm_buildprods.lst_temp.List(nn)) Then
        Exit For
      End If
    Next nn
    
    ' Add product if not in list
    If (nn = frm_buildprods.lst_temp.ListCount) Then
      frm_buildprods.lst_temp.AddItem (frm_buildprods.lst_includeini.List(zz))
    Else
      pos = InStr(frm_buildprods.lst_includeini.List(zz), ",")
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_buildprods.errMsg1", "%1 product already in list. The product will not be added", Left(frm_buildprods.lst_includeini.List(zz), pos - 1)), vbExclamation
    End If
  Next zz

  If (uniFile.OpenFileWrite(CFG_DIR & TMP_PRODUCTS_CFG_FILE) = False) Then GoTo FILE_ERROR2
  
  On Error GoTo FILE_ERROR2
  uniFile.WriteBOM fe_UTF16LE

  For zz = 0 To (frm_buildprods.lst_temp.ListCount - 1)
    uniFile.WriteUnicodeLine Trim(frm_buildprods.lst_temp.List(zz))
  Next zz

  uniFile.Flush
  uniFile.CloseFile
  uniFile.st_CopyFile tmpFile, fileName
  uniFile.st_RmFile tmpFile
  frmProduct.m_prodListChanged = True
  Unload frm_buildprods
  Exit Sub

FILE_ERROR1:
  errMsg = (tmpFile & " file read error. " & Error$)
  uniMsg = MLSupport.GGS_Params("fileErrMsg4", "%1 file read error. %2", tmpFile, Error$)
  GoTo FILE_ERROR
  
FILE_ERROR2:
  errMsg = (tmpFile & " file write error. " & Error$)
  uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", tmpFile, Error$)
  GoTo FILE_ERROR
  
FILE_ERROR:
  uniFile.CloseFile
  
  If (uniFile.st_FileExist(CFG_DIR & TMP_PRODUCTS_CFG_FILE) = True) Then
    uniFile.st_RmFile (CFG_DIR & TMP_PRODUCTS_CFG_FILE)
  End If
  
  unity_main.errorstring = errMsg
  unity_main.write_error
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  Unload frm_buildprods
End Sub

Private Sub Form_Load()
  
  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
  File1.Path = PRODUCTS_CFG_DIR
  File1.Pattern = ("*" & CFG_FILE_EXT)
  makeinilist
End Sub

Sub makeinilist()
  Dim zz As Integer
  Dim flname As String
  
  For zz = 0 To frm_buildprods.File1.ListCount - 1
    flname = frm_buildprods.File1.List(zz)
    lst_inis.AddItem (flname)
  Next zz
End Sub

Private Sub lst_inis_DblClick()
  
  frm_buildprods.getheader
End Sub








