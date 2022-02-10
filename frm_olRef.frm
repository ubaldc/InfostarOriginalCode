VERSION 5.00
Begin VB.Form frm_olRef 
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frm_olRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_olRefEndWvln As Double
Public m_olRefFileName As String
Public m_olRefStartWvln As Double

Function get_ol_ref_wvlns(olRefFileName As String) As Boolean
  Dim rc As Boolean
  Dim spcIO As GSpcIOLib.GSPCio
  Dim numSubfiles As Long
  Dim errStrg As String
  Dim subFileIndx As Long
  Dim varXVals As Variant
  Dim varYVals As Variant
   
  ' Load reference qualification file
  m_olRefFileName = (REFERENCES_DIR & olRefFileName & SPC_FILE_EXT)
  rc = LoadSpcFile(m_olRefFileName, spcIO, numSubfiles, errStrg)
  
  If (rc = True) Then
    ' Get reference wavelengths
    m_olRefStartWvln = spcIO.FirstPoint
    m_olRefEndWvln = spcIO.LastPoint
    Set spcIO = Nothing
    CloseSpcFile
  Else
    unity_main.m_ansiErrMsg = "Error opening spectrum file: " & m_olRefFileName
    unity_main.m_uniErrMsg = MLSupport.GGS_Params("errMsg4", "Error opening spectrum file: %1", m_olRefFileName)
    unity_main.errorstring = unity_main.m_ansiErrMsg
    unity_main.write_error
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", unity_main.m_uniErrMsg), vbCritical
  End If
  
  get_ol_ref_wvlns = rc
End Function

Function setup_ol_ref(olRefFileName As String) As Boolean
  Dim rc As Boolean
  Dim spcIO As GSpcIOLib.GSPCio
  Dim numSubfiles As Long
  Dim errStrg As String
  Dim subFileIndx As Long
  Dim varXVals As Variant
  Dim varYVals As Variant
  Dim uniMsg As String
   
  ' Load reference qualification file
  m_olRefFileName = (REFERENCES_DIR & olRefFileName & SPC_FILE_EXT)
  rc = LoadSpcFile(m_olRefFileName, spcIO, numSubfiles, errStrg)
  
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
        unity_main.m_ansiErrMsg = "Error setting offline reference scan data"
        unity_main.m_uniErrMsg = MLSupport.GSS("OperStatus", "status80", "Error setting offline reference scan data")
      End If
    Else
      unity_main.m_ansiErrMsg = "Error reading spectrum file: " & m_olRefFileName
      unity_main.m_uniErrMsg = MLSupport.GGS_Params("errMsg5", "Error reading from spectrum file: %1", m_olRefFileName)
    End If
  
    Set spcIO = Nothing
    CloseSpcFile
  Else
    unity_main.m_ansiErrMsg = "Error opening spectrum file: " & m_olRefFileName
    unity_main.m_uniErrMsg = MLSupport.GGS_Params("errMsg4", "Error opening spectrum file: %1", m_olRefFileName)
  End If
  
  If (rc = False) Then
    unity_main.errorstring = unity_main.m_ansiErrMsg
    unity_main.write_error
  End If

  setup_ol_ref = rc
End Function

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub



