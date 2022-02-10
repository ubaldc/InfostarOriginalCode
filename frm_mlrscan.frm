VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frm_mlrscan 
   Caption         =   "MLR Scan form "
   ClientHeight    =   6075
   ClientLeft      =   465
   ClientTop       =   840
   ClientWidth     =   6045
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
   ScaleHeight     =   6075
   ScaleWidth      =   6045
Begin HexUniControls.ctlUniTextBoxXP txtpred
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin FPUSpreadADO.fpSpread gridscan 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2655
      _Version        =   458752
      _ExtentX        =   4683
      _ExtentY        =   8705
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      MaxRows         =   3000
      SpreadDesigner  =   "frm_mlrscan.frx":0000
   End
End
Attribute VB_Name = "frm_mlrscan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub domlrpred()
  Dim tempAbs As Double
  Dim zz As Integer
  Dim initWl As Integer
  Dim inString As String
  Dim numWls As Integer
  Dim bias As Single
  Dim wl As Single
  Dim fVal As Single
  Dim absVal As Single
  Dim predValue As Single
  Dim tempval As Single
  Dim top_counter As Integer
  Dim jj As Integer
  Dim smPts As Integer ' # of pts to smooth (2n+1)
  Dim modVers As Integer ' if 1 = no smoothing (just wl,f value), if 2 = wl,fval,smoothpts
  Dim errMsg As String
  Dim lineCnt As Integer
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim rc As Boolean
  Dim varStr As Variant
  Dim uniMsg As String
  
#If ABBFT Then
  top_counter = (unity_main.m_mb3000.m_endWavenum - unity_main.m_mb3000.m_startWavenum) / unity_main.m_mb3000.m_waveNumIncr + 1
#Else
  top_counter = (unity_main.m_smplEndWvln - unity_main.m_smplStartWvln) / MS11CfgData.wvlnIncr + 1
#End If

  ' Check if spectrum data treated
  If (unity_main.m_enableTreatment = True) Then
    For zz = 1 To top_counter
      frm_mlrscan.gridscan.Row = zz
      frm_mlrscan.gridscan.Col = 0
      frm_mlrscan.gridscan.Text = ProdSmplXVals(zz - 1)
      frm_mlrscan.gridscan.Col = 1
      frm_mlrscan.gridscan.Text = ProdTreatAbsYVals(zz - 1)
    Next zz
  Else
    For zz = 1 To top_counter
      frm_mlrscan.gridscan.Row = zz
      frm_mlrscan.gridscan.Col = 0
      frm_mlrscan.gridscan.Text = ProdSmplXVals(zz - 1)
      frm_mlrscan.gridscan.Col = 1
      frm_mlrscan.gridscan.Text = ProdAbsYVals(zz - 1)
    Next zz
  End If

  'make mlr predictions
  zz = 1
  frm_mlrscan.gridscan.Row = 1
  frm_mlrscan.gridscan.Col = 0
  initWl = frm_mlrscan.gridscan.value
 
  uniMsg = MLSupport.GGS_Params("frm_mlrscan.statMsg1", "Loading MLR model: %1", unity_main.modlname)
  Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Loading MLR model: " & unity_main.modlname), uniMsg)
  
  If (uniFile.st_FileExist(unity_main.fullmodelname) = False) Then
    errMsg = (unity_main.fullmodelname & " MLR model not found")
    uniMsg = MLSupport.GGS_Params("frm_mlrscan.errMsg1", "%1 MLR model not found", unity_main.fullmodelname)
    Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    GoTo PROB_MLR2
  End If

  If (uniFile.OpenFileRead(unity_main.fullmodelname) = True) Then
    On Error GoTo PROB_MLR
    fEncoding = uniFile.ReadBOM
    lineCnt = lineCnt + 1
   
    ' Number of wavelengths
    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(inString)
    Else
      rc = uniFile.ReadUnicodeLine(inString)
    End If
      
    If (rc = False) Then GoTo PROB_MLR
  
    numWls = Trim(inString)
    lineCnt = lineCnt + 1
    
    ' Model info
    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(inString)
    Else
      rc = uniFile.ReadUnicodeLine(inString)
    End If
      
    If (rc = False) Then GoTo PROB_MLR
    
    lineCnt = lineCnt + 1
    
    ' Property
    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(inString)
    Else
      rc = uniFile.ReadUnicodeLine(inString)
    End If
      
    If (rc = False) Then GoTo PROB_MLR
  
    lineCnt = lineCnt + 1
    
    ' Bias
    If (fEncoding = fe_ANSI) Then
      rc = uniFile.ReadAnsiLine(inString)
    Else
      rc = uniFile.ReadUnicodeLine(inString)
    End If
      
    If (rc = False) Then GoTo PROB_MLR
  
    bias = Trim(inString)
    uniMsg = MLSupport.GGS_Params("frm_mlrscan.statMsg2", "Applying Coefficient and P/Side for %1 wavelengths", CStr(numWls))
    Call unity_main.log_scan_status(LOG_DBG_LEVEL2, ("Applying Coefficient and P/Side for " & numWls & " wavelengths"), uniMsg)

    For zz = 1 To numWls
      lineCnt = lineCnt + 1
      
      If (fEncoding = fe_ANSI) Then
        rc = uniFile.ReadAnsiLine(inString)
      Else
        rc = uniFile.ReadUnicodeLine(inString)
      End If
      
      If (rc = False) Then GoTo PROB_MLR
  
      varStr = Split(inString, ",")
      modVers = UBound(varStr)

      If (modVers = 1) Then
        smPts = 0
      Else
        smPts = Trim(varStr(2))
      End If
    
      wl = Trim(varStr(0))
      fVal = Trim(varStr(1))

      'validate wavelengths
      If (wl < unity_main.m_smplStartWvln) Or (wl > unity_main.m_smplEndWvln) Then
        errMsg = (unity_main.fullmodelname & " MLR model contains wavelength beyond the range of this product (" & wl & "; " & unity_main.m_smplStartWvln & " - " & unity_main.m_smplEndWvln & ")")
        uniMsg = MLSupport.GGS_Params("frm_mlrscan.errMsg2", "%1 MLR model contains wavelength beyond the range of this product (%2; %3 - %4)", unity_main.fullmodelname, CStr(wl), CStr(unity_main.m_smplStartWvln), CStr(unity_main.m_smplEndWvln))
        Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
        CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
        GoTo PROB_MLR2
      End If

      'add smoothing 10/04
      frm_mlrscan.gridscan.Row = wl - initWl + 1
      frm_mlrscan.gridscan.Col = 1
    
      If (smPts = 0) Then
        absVal = frm_mlrscan.gridscan.value
      Else
        tempAbs = 0
      
        For jj = -smPts To smPts Step 1
          frm_mlrscan.gridscan.Row = wl - initWl + 1 + jj
          absVal = frm_mlrscan.gridscan.value
          tempAbs = tempAbs + absVal
        Next jj
      
        absVal = tempAbs / ((smPts * 2) + 1)
      End If

      tempval = absVal * fVal
      predValue = predValue + tempval
    Next zz
  
    predValue = predValue + bias
  Else
PROB_MLR:
    If (lineCnt = 0) Then
      errMsg = (unity_main.fullmodelname & " file open error. " & Error$)
      uniMsg = MLSupport.GGS_Params("fileErrMsg2", "%1 file open error. %2", unity_main.fullmodelname, Error$)
    Else
      errMsg = (unity_main.fullmodelname & " file has error on line " & CStr(lineCnt) & ". " & Error$)
      uniMsg = MLSupport.GGS_Params("fileErrMsg3", "%1 file has error on line %2. %3", unity_main.fullmodelname, CStr(lineCnt), Error$)
    End If
  
    Call unity_main.log_scan_status(LOG_DBG_LEVEL1, errMsg, uniMsg)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
    
PROB_MLR2:
    predValue = 0
    unity_main.pukedonpred = True
  End If
  
  uniFile.CloseFile
  txtpred.Text = predValue
End Sub  'domlrpred

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub




