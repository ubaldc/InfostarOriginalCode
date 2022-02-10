VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Object = "{488D84E9-A9A9-4C74-949B-F7752803ADEE}#99.2#0"; "EMDVirtualKeyBoard.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frm_labData 
   Caption         =   "Lab Data Management"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
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
   ScaleHeight     =   9390
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniLabelNoFlick lbl_svfFileName 
      Height          =   660
      Left            =   2280
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1164
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_labData.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_labData.frx":0020
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_labData.frx":0040
   End
   Begin FPUSpreadADO.fpSpread ss_labData 
      Height          =   4225
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   10935
      _Version        =   458752
      _ExtentX        =   19288
      _ExtentY        =   7461
      _StockProps     =   64
      ColHeaderDisplay=   0
      EditModePermanent=   -1  'True
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   -2147483633
      MaxCols         =   51
      MaxRows         =   200
      ProcessTab      =   -1  'True
      SelectBlockOptions=   0
      SpreadDesigner  =   "frm_labData.frx":005C
      UserResize      =   1
   End
   Begin HexUniControls.ctlUniLabel lbl_prodName 
      Height          =   420
      Left            =   2280
      Top             =   720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_labData.frx":02D9
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_labData.frx":02F9
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_labData.frx":0319
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9360
      Top             =   8880
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9390
      FormDesignWidth =   11250
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   645
      Left            =   9000
      TabIndex        =   0
      Top             =   7800
      Width           =   1995
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
      Caption         =   "frm_labData.frx":0335
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_labData.frx":0361
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_labData.frx":0381
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_save 
      Height          =   645
      Left            =   9000
      TabIndex        =   1
      Top             =   6840
      Width           =   1995
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
      Caption         =   "frm_labData.frx":039D
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_labData.frx":03D5
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_labData.frx":03F5
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_browse 
      Height          =   645
      Left            =   9000
      TabIndex        =   2
      Top             =   120
      Width           =   1995
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
      Caption         =   "frm_labData.frx":0411
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_labData.frx":043D
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_labData.frx":045D
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   420
      Left            =   240
      Top             =   120
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_labData.frx":0479
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_labData.frx":04B1
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_labData.frx":04D1
   End
   Begin HexUniControls.ctlUniLabel Label2 
      Height          =   420
      Left            =   240
      Top             =   720
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_labData.frx":04ED
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_labData.frx":0525
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_labData.frx":0545
   End
   Begin VBoard_EMD.KeySet_STD KeySet_STD1 
      Height          =   2430
      Left            =   240
      Top             =   5760
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   4286
      LargeFontName   =   "Arial Unicode MS"
      LargeFontSize   =   6.75
      SmallFontName   =   "Arial Unicode MS"
      SendKeysEnabled =   -1  'True
      RegAtDT         =   -1  'True
   End
   Begin HexUniControls.ctlUniFormCaption uniTitle 
      Left            =   1080
      Top             =   8760
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_labData.frx":0561
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   10320
      Top             =   8880
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
Attribute VB_Name = "frm_labData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_lastSVFDir As String
Private m_fileName As String
Private m_numConstituents As Long
Private m_numSamples As Long
Private m_labDataLoaded As Boolean

Public Sub setup_spread()

  ss_labData.StartingRowNumber = 1
  ss_labData.ColHeadersShow = True
  ss_labData.MaxRows = 200
  ss_labData.MaxCols = MAX_NUM_PROPS + 1
  ss_labData.ClearRange 1, 1, ss_labData.MaxCols, ss_labData.MaxRows, True

  ' Setup header font info
  ss_labData.Row = 0
  ss_labData.Col = 0
  ss_labData.Row2 = 0
  ss_labData.Col2 = ss_labData.MaxCols
  ss_labData.BlockMode = True
  ss_labData.Font.Name = "Arial Unicode MS"
  ss_labData.Font.Size = 10
  ss_labData.BlockMode = False
 
  ' Setup 1st column label and width
  ss_labData.Row = 0
  ss_labData.Col = 1
  ss_labData.Font.Name = "Arial Unicode MS"
  ss_labData.Font.Size = 10
  ss_labData.FontBold = True
  ss_labData.Text = "Sample ID"
  ss_labData.ColWidth(1) = 20
  
  ss_labData.EditEnterAction = EditEnterActionNext
  m_labDataLoaded = False
End Sub

Private Sub get_file_info()
  Dim rc As Long
  Dim ssSerialNum As String
  Dim masterSerialNum As String
  Dim clientName As String
  Dim productName As String
  Dim hdrComment As String
  Dim startWaveln As Double
  Dim endWaveln As Double
  Dim wavelnInc As Double
  Dim uniMsg As String
  
  ' Get SVF file header info
  On Error GoTo OBJECT_ERROR
  rc = SVFObject.getFileHeader(m_fileName, ssSerialNum, masterSerialNum, clientName, productName, hdrComment, startWaveln, endWaveln, wavelnInc, m_numConstituents, m_numSamples)
  
  If (rc = 0) Then
    lbl_prodName.Caption = productName
    ss_labData.MaxRows = m_numSamples
    ss_labData.MaxCols = m_numConstituents + 1    ' include 1st column for sample name
    
    ' Setup cells font info and clear data
    ss_labData.Row = 1
    ss_labData.Col = 1
    ss_labData.Row2 = ss_labData.MaxRows
    ss_labData.Col2 = ss_labData.MaxCols
    ss_labData.BlockMode = True
    ss_labData.Font.Name = "Arial Unicode MS"
    ss_labData.Font.Size = 10
    ss_labData.BlockMode = False
    ss_labData.ClearRange 1, 1, ss_labData.MaxCols, ss_labData.MaxRows, True

    If (m_numConstituents > 0) Then
      get_lab_data
    End If
    
    Exit Sub
  End If
  
  unity_main.errorstring = m_fileName & " UCal SVF spectra file error: " & rc
  unity_main.write_error (LOG_DBG_LEVEL1)
  uniMsg = MLSupport.GGS_Params("frm_labData.errMsg1", "%1 UCal SVF spectra file error: %2", m_fileName, CStr(rc))
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  Exit Sub
  
OBJECT_ERROR:
  unity_main.errorstring = "Unity SVFComponent.dll component not installed or registered"
  unity_main.write_error (LOG_DBG_LEVEL1)
  uniMsg = MLSupport.GGS_Params("errMsg7", "Unity %1 component not installed or registered", "SVFComponent.dll")
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub

Private Sub get_lab_data()
  Dim rc As Long
  Dim constituentNames() As String
  Dim sampleIDs() As String
  Dim labData() As Single
  Dim ii As Long
  Dim jj As Long
  Dim uniMsg As String
    
  ' Get SVF file constituent names
  On Error GoTo OBJECT_ERROR
  ReDim constituentNames(m_numConstituents)
  rc = SVFObject.getConstituentNames(m_fileName, 0, m_numConstituents, constituentNames)
  
  If (rc = 0) Then
    For ii = 0 To m_numConstituents - 1
      ' Setup column label with constituent name and width
      ss_labData.Row = 0
      ss_labData.Col = 2 + ii
      ss_labData.Font.Name = "Arial Unicode MS"
      ss_labData.Font.Size = 10
      ss_labData.FontBold = True
      ss_labData.Text = constituentNames(ii)
      ss_labData.ColWidth(2 + ii) = 15
    Next ii
    
    If (m_numSamples > 0) Then
      ' Get samples lab data
      ReDim sampleIDs(0 To m_numSamples - 1)
      ReDim labData(0 To m_numSamples * m_numConstituents - 1)
      rc = SVFObject.getMultipleLabData(m_fileName, 0, m_numSamples, 0, m_numConstituents, sampleIDs, labData)
      
      If (rc <> 0) Then GoTo BAD_FILE
      
      ' Show sample ID and lab data for each sample record
      For ii = 0 To m_numSamples - 1
        ' Show ID for sample record
        ss_labData.Row = 1 + ii
        ss_labData.Col = 1
        ss_labData.Font.Name = "Arial Unicode MS"
        ss_labData.Font.Size = 10
        ss_labData.FontBold = False
        ss_labData.Text = sampleIDs(ii)
      
        ' Show lab data for sample record
        For jj = 0 To m_numConstituents - 1
          ss_labData.Col = 2 + jj
          ss_labData.Font.Name = "Arial Unicode MS"
          ss_labData.Font.Size = 10
          ss_labData.FontBold = False
          ss_labData.Value = labData(ii * m_numConstituents + jj)
        Next jj
      Next ii
    
      ' Setup column 1 for static cells w/ word wrap
      ss_labData.Row = 1
      ss_labData.Col = 1
      ss_labData.Row2 = ss_labData.MaxRows
      ss_labData.Col2 = 1
      ss_labData.BlockMode = True
      ss_labData.CellType = CellTypeStaticText
      ss_labData.TypeTextWordWrap = True
      ss_labData.BlockMode = False
      
      m_labDataLoaded = True
    End If
    
    ss_labData.SetActiveCell 1, 1
    Exit Sub
  End If
  
BAD_FILE:
  unity_main.errorstring = m_fileName & " UCal SVF spectra file error: " & rc
  unity_main.write_error (LOG_DBG_LEVEL1)
  uniMsg = MLSupport.GGS_Params("frm_labData.errMsg1", "%1 UCal SVF spectra file error: %2", m_fileName, CStr(rc))
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  Exit Sub
  
OBJECT_ERROR:
  unity_main.errorstring = "Unity SVFComponent.dll component not installed or registered"
  unity_main.write_error (LOG_DBG_LEVEL1)
  uniMsg = MLSupport.GGS_Params("errMsg7", "Unity %1 component not installed or registered", "SVFComponent.dll")
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub

Private Sub cmd_browse_Click()
  Dim dialog As New clsCommDialogs
  Dim fileDir As String
  Dim sFilter As String
  Dim dlgTitle As String
  Dim strlen As Integer
  Dim indx As Integer

  unity_main.errorstring = "Enter Lab Data screen Browse button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  On Error GoTo BAD_FILE
  dialog.InitDialogs
  
  If (m_lastSVFDir = "") Then
    m_lastSVFDir = unity_main.m_saveDir
  End If
  
  fileDir = m_lastSVFDir
  sFilter = ("UCal SVF (*" & SVF_FILE_EXT & ")" & Chr(0) + "*" & SVF_FILE_EXT & Chr(0))
  dlgTitle = MLSupport.GSS("frm_labData", "dlgTitle", "Select UCal SVF Spectra File")
  m_fileName = dialog.ShowOpen(Me.hWnd, sFilter, dlgTitle, OFN_FILEMUSTEXIST + OFN_HIDEREADONLY + OFN_NOCHANGEDIR, fileDir)

  If (m_fileName <> "") Then
    strlen = Len(m_fileName)
    indx = InStrRev(m_fileName, "\")
    m_lastSVFDir = Left(m_fileName, indx)
    frm_labData.lbl_svfFileName.Caption = m_fileName
    m_labDataLoaded = False
    get_file_info
  End If

  Exit Sub

BAD_FILE:
  CWrap.ShowMessageBoxW MLSupport.GSS("frm_labData", "errMsg1", "Error selecting file, please confirm you selected a valid file"), vbCritical
End Sub

Private Sub cmd_cancel_Click()

  unity_main.errorstring = "Enter Lab Data screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Unload frm_labData
End Sub

Private Sub cmd_save_Click()
  Dim ii As Long
  Dim jj As Long
  Dim labData() As Single
  Dim rc As Long
  Dim uniMsg As String

  unity_main.errorstring = "Enter Lab Data screen Save Changes button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  If (m_labDataLoaded = False) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_labData", "errMsg2", "Have not selected and loaded a valid file"), vbCritical
    Exit Sub
  End If
  
  ReDim labData(0 To m_numSamples * m_numConstituents - 1)
  
  ' Check lab data for each sample record
  For ii = 0 To m_numSamples - 1
    ss_labData.Row = 1 + ii
      
    For jj = 0 To m_numConstituents - 1
      ss_labData.Col = 2 + jj
          
      If (IsNumeric(ss_labData.Text) = False) Then
        ss_labData.SetActiveCell ss_labData.Col, ss_labData.Row
        CWrap.ShowMessageBoxW MLSupport.GSS("frm_labData", "errMsg3", "Invalid lab data value. Correct value to save to file"), vbCritical
        Exit Sub
      End If
      
      labData(ii * m_numConstituents + jj) = ss_labData.Value
    Next jj
  Next ii
    
  On Error GoTo OBJECT_ERROR
  rc = SVFObject.saveMultipleLabData(m_fileName, 0, m_numSamples, 0, m_numConstituents, labData)
      
  If (rc <> 0) Then
    unity_main.errorstring = m_fileName & " UCal SVF spectra file error: " & rc
    unity_main.write_error (LOG_DBG_LEVEL1)
    uniMsg = MLSupport.GGS_Params("frm_labData.errMsg1", "%1 UCal SVF spectra file error: %2", m_fileName, CStr(rc))
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  Else
    Unload frm_labData
  End If
  
  Exit Sub
  
OBJECT_ERROR:
  unity_main.errorstring = "Unity SVFComponent.dll component not installed or registered"
  unity_main.write_error (LOG_DBG_LEVEL1)
  uniMsg = MLSupport.GGS_Params("errMsg7", "Unity %1 component not installed or registered", "SVFComponent.dll")
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub








