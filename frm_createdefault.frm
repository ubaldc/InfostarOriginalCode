VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_createdefault 
   Caption         =   "Create Initial Default Configuration File"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
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
   Icon            =   "frm_createdefault.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   240
      Top             =   3480
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4545
      FormDesignWidth =   9015
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_shutdown 
      Cancel          =   -1  'True
      Height          =   840
      Left            =   5160
      TabIndex        =   0
      Top             =   3240
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   1482
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
      Caption         =   "frm_createdefault.frx":0442
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_createdefault.frx":0472
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_createdefault.frx":0492
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_createDflt 
      Height          =   840
      Left            =   1080
      TabIndex        =   1
      Top             =   3240
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   1482
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
      Caption         =   "frm_createdefault.frx":04AE
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_createdefault.frx":0506
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_createdefault.frx":0526
   End
   Begin HexUniControls.ctlUniLabel Label1 
      Height          =   2535
      Left            =   480
      Top             =   240
      Width           =   8055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_createdefault.frx":0542
      BackColor       =   -2147483633
      ForeColor       =   255
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_createdefault.frx":0716
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_createdefault.frx":0736
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   240
      Top             =   3000
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
      Left            =   240
      Top             =   4080
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_createdefault.frx":0752
   End
End
Attribute VB_Name = "frm_createdefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_createDflt_Click()
  
  unity_main.errorstring = "Create Initial Default Configuration File screen Create Default Configuration button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Call frm_createdefault.make_prod_default_file
  Unload frm_createdefault
  Call unity_main.load_prod_file(PROD_DFLTS_CFG_FILE, True)
End Sub

Sub make_prod_default_file()
  Dim fileName As String
  Dim uniFile As New clsUniFile
  Dim errMsg As String
  Dim uniMsg As String

  fileName = (PRODUCTS_CFG_DIR & PROD_DFLTS_CFG_FILE)
  
  If (uniFile.OpenFileWrite(fileName) = True) Then
    On Error GoTo FILE_ERROR
    uniFile.WriteBOM fe_UTF16LE
    uniFile.WriteUnicodeLine "[product type, sampling]"
    
#If ABBFT Then
    uniFile.WriteUnicodeLine "Default,External Ref"
    uniFile.WriteUnicodeLine "[signature settings]"
    uniFile.WriteUnicodeLine ("DevID=" & DTID_ABBFT)
    uniFile.WriteUnicodeLine ("Version=" & INFOSTAR_VER)
    uniFile.WriteUnicodeLine "[analyzer settings]"
    uniFile.WriteUnicodeLine ("Background_Frequency=" & ProdDfltData.backFreq)
    uniFile.WriteUnicodeLine ("Bound_Values=" & ProdDfltData.boundValue)
    uniFile.WriteUnicodeLine ("BType=" & ProdDfltData.btype)
    uniFile.WriteUnicodeLine ("ClrManualName=" & ProdDfltData.clrManualName)
    uniFile.WriteUnicodeLine ("ClrUserInputs=" & ProdDfltData.clrUserInputs)
    uniFile.WriteUnicodeLine ("Day_Counter=" & ProdDfltData.dayCounter)
    uniFile.WriteUnicodeLine ("DelayMeasure=" & ProdDfltData.delayMeasure)
    uniFile.WriteUnicodeLine ("DelayStart=" & ProdDfltData.delayStart)
    uniFile.WriteUnicodeLine ("EndWavenumIndx=" & ProdDfltData.endWavenumIndx)
    uniFile.WriteUnicodeLine ("GainIndx=" & ProdDfltData.gainIndx)
    uniFile.WriteUnicodeLine ("HideValCol=" & ProdDfltData.hideValCol)
    uniFile.WriteUnicodeLine ("MakePred=" & ProdDfltData.makePred)
    uniFile.WriteUnicodeLine ("MD_Alarm=" & ProdDfltData.mdAlarm)
    uniFile.WriteUnicodeLine ("Menu_Input_Buttons=" & ProdDfltData.useMenuInput)
    uniFile.WriteUnicodeLine ("NameBase=" & ProdDfltData.namebase)
    uniFile.WriteUnicodeLine ("NameCounter=" & ProdDfltData.nameCounter)
    uniFile.WriteUnicodeLine ("NameScanType=" & ProdDfltData.nameScanType)
    uniFile.WriteUnicodeLine ("ND_Alarm=" & ProdDfltData.ndAlarm)
    uniFile.WriteUnicodeLine ("NoOLVal=" & ProdDfltData.noOLVal)
    uniFile.WriteUnicodeLine ("NoRefVal=" & ProdDfltData.noPredVal)
    uniFile.WriteUnicodeLine ("NumMeasures=" & ProdDfltData.numMeasures)
    uniFile.WriteUnicodeLine ("NumSamples=" & ProdDfltData.numSamples)
    uniFile.WriteUnicodeLine ("Outlier_Lights=" & ProdDfltData.outlierLights)
    uniFile.WriteUnicodeLine ("ResolutionIndx=" & ProdDfltData.resolutionIndx)
    uniFile.WriteUnicodeLine ("RR_Alarm=" & ProdDfltData.rrAlarm)
    uniFile.WriteUnicodeLine ("SaveIt=" & ProdDfltData.saveIt)
    uniFile.WriteUnicodeLine ("SaveScansDir=" & ProdDfltData.saveScanDir)
    uniFile.WriteUnicodeLine ("Save_CsvFile=" & ProdDfltData.saveCSVFile)
    uniFile.WriteUnicodeLine ("Save_PredFile=" & ProdDfltData.savePredFile)
    uniFile.WriteUnicodeLine ("Save_Predictions=" & ProdDfltData.savePredictions)
    uniFile.WriteUnicodeLine ("Save_Predictions_Csv=" & ProdDfltData.savePredictionsCSV)
    uniFile.WriteUnicodeLine ("Save_Predictions_DynRpt=" & ProdDfltData.savePredictionsDynRpt)
    uniFile.WriteUnicodeLine ("Send_LIMS_Output=" & ProdDfltData.sendLimsOutput)
    uniFile.WriteUnicodeLine ("SpeedIndx=" & ProdDfltData.speedIndx)
    uniFile.WriteUnicodeLine ("StartWavenumIndx=" & ProdDfltData.startWavenumIndx)
    uniFile.WriteUnicodeLine ("SType=" & ProdDfltData.sType)
    uniFile.WriteUnicodeLine ("Value_Alarm=" & ProdDfltData.valueAlarm)
    uniFile.WriteUnicodeLine ("Write_Ticket_Printer=" & ProdDfltData.writeTkt)
#Else
    Select Case (MS11CfgData.devID)
      Case DTID_DRAWER0            ' SS2200/SS2400 standard drawer system
        uniFile.WriteUnicodeLine "Default,Internal Ref"
      Case DTID_TOPWIND0           ' Top window w/out internal reflectance
        uniFile.WriteUnicodeLine "Default,External Ref"
      Case DTID_DRAWER1            ' SS2200/SS2400 drawer w/out stepper system
        uniFile.WriteUnicodeLine "Default,Internal Ref"
      Case DTID_TOPWIND1           ' Top window with internal reflectance
        uniFile.WriteUnicodeLine "Default,Internal Ref"
    End Select
  
    uniFile.WriteUnicodeLine "[signature settings]"
    uniFile.WriteUnicodeLine ("DevID=" & MS11CfgData.devID)
    uniFile.WriteUnicodeLine ("SmplTable=" & MS11CfgData.smplTblIX)
    uniFile.WriteUnicodeLine ("ScanMode=" & MS11CfgData.sysScanMode)
    uniFile.WriteUnicodeLine ("Version=" & INFOSTAR_VER)
    uniFile.WriteUnicodeLine "[analyzer settings]"
    uniFile.WriteUnicodeLine ("AdapterType=" & ProdDfltData.adapterType)
    uniFile.WriteUnicodeLine ("Background_Frequency=" & ProdDfltData.backFreq)
    uniFile.WriteUnicodeLine ("Bound_Values=" & ProdDfltData.boundValue)
    uniFile.WriteUnicodeLine ("BType=" & ProdDfltData.btype)
    uniFile.WriteUnicodeLine ("ClrManualName=" & ProdDfltData.clrManualName)
    uniFile.WriteUnicodeLine ("ClrUserInputs=" & ProdDfltData.clrUserInputs)
    uniFile.WriteUnicodeLine ("Day_Counter=" & ProdDfltData.dayCounter)
    uniFile.WriteUnicodeLine ("EndWvln=" & ProdDfltData.endWvln)
    uniFile.WriteUnicodeLine ("ExtRefFile=" & ProdDfltData.extRefFileName)
    uniFile.WriteUnicodeLine ("HideValCol=" & ProdDfltData.hideValCol)
    uniFile.WriteUnicodeLine ("MakePred=" & ProdDfltData.makePred)
    uniFile.WriteUnicodeLine ("MD_Alarm=" & ProdDfltData.mdAlarm)
    uniFile.WriteUnicodeLine ("Menu_Input_Buttons=" & ProdDfltData.useMenuInput)
    uniFile.WriteUnicodeLine ("MultiCupType=" & ProdDfltData.multiCupType)
    uniFile.WriteUnicodeLine ("NameBase=" & ProdDfltData.namebase)
    uniFile.WriteUnicodeLine ("NameCounter=" & ProdDfltData.nameCounter)
    uniFile.WriteUnicodeLine ("NameScanType=" & ProdDfltData.nameScanType)
    uniFile.WriteUnicodeLine ("ND_Alarm=" & ProdDfltData.ndAlarm)
    uniFile.WriteUnicodeLine ("NoOLVal=" & ProdDfltData.noOLVal)
    uniFile.WriteUnicodeLine ("NoRefVal=" & ProdDfltData.noPredVal)
    uniFile.WriteUnicodeLine ("NScansS=" & ProdDfltData.numSmplScans)
    uniFile.WriteUnicodeLine ("OLRefFile=" & ProdDfltData.olRefFileName)
    uniFile.WriteUnicodeLine ("Outlier_Lights=" & ProdDfltData.outlierLights)
    uniFile.WriteUnicodeLine ("RepsAvg=" & ProdDfltData.repsAvg)
    uniFile.WriteUnicodeLine ("RotateDir=" & ProdDfltData.rotateDir)
    uniFile.WriteUnicodeLine ("RotateIndexSteps=" & ProdDfltData.rotateIndexSteps)
    uniFile.WriteUnicodeLine ("RotateMoveMode=" & ProdDfltData.rotateMoveMode)
    uniFile.WriteUnicodeLine ("RotateSpeed=" & ProdDfltData.rotateSpeed)
    uniFile.WriteUnicodeLine ("RotateStepSteps=" & ProdDfltData.rotateStepSteps)
    uniFile.WriteUnicodeLine ("RR_Alarm=" & ProdDfltData.rrAlarm)
    uniFile.WriteUnicodeLine ("SaveIt=" & ProdDfltData.saveIt)
    uniFile.WriteUnicodeLine ("SaveScansDir=" & ProdDfltData.saveScanDir)
    uniFile.WriteUnicodeLine ("Save_CsvFile=" & ProdDfltData.saveCSVFile)
    uniFile.WriteUnicodeLine ("Save_PredFile=" & ProdDfltData.savePredFile)
    uniFile.WriteUnicodeLine ("Save_Predictions=" & ProdDfltData.savePredictions)
    uniFile.WriteUnicodeLine ("Save_Predictions_Csv=" & ProdDfltData.savePredictionsCSV)
    uniFile.WriteUnicodeLine ("Save_Predictions_DynRpt=" & ProdDfltData.savePredictionsDynRpt)
    uniFile.WriteUnicodeLine ("Save_Replicates=" & ProdDfltData.saveReplicates)
    uniFile.WriteUnicodeLine ("Send_LIMS_Output=" & ProdDfltData.sendLimsOutput)
    uniFile.WriteUnicodeLine ("SmplPPT=" & ProdDfltData.smplPPT)
    uniFile.WriteUnicodeLine ("StartWvln=" & ProdDfltData.startWvln)
    uniFile.WriteUnicodeLine ("SType=" & ProdDfltData.sType)
    uniFile.WriteUnicodeLine ("UseExtRefTrayCfg=" & ProdDfltData.useExtRefTrayCfg)
    uniFile.WriteUnicodeLine ("Value_Alarm=" & ProdDfltData.valueAlarm)
    uniFile.WriteUnicodeLine ("Write_Ticket_Printer=" & ProdDfltData.writeTkt)
#End If

    uniFile.WriteUnicodeLine "[analysis models]"
    uniFile.Flush
  Else
FILE_ERROR:
    errMsg = (PRODUCTS_CFG_DIR & PROD_DFLTS_CFG_FILE & " file write error. " & Error$)
    unity_main.errorstring = errMsg
    unity_main.write_error
    uniMsg = MLSupport.GGS_Params("fileErrMsg5", "%1 file write error. %2", PRODUCTS_CFG_DIR & PROD_DFLTS_CFG_FILE, Error$)
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("errMsg1", "%1. Please contact your supervisor!", uniMsg), vbCritical
  End If
  
  uniFile.CloseFile
End Sub

Private Sub cmd_shutdown_Click()
  
  unity_main.errorstring = "Create Initial Default Configuration File screen Shutdown button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  End
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub








