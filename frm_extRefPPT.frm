VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_extRefPPT 
   Caption         =   "External Reference PPT Qualification"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
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
   ScaleHeight     =   5100
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   4080
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5100
      FormDesignWidth =   6600
   End
   Begin HexUniControls.ctlUniListBoxXP lst_refFileNames 
      Height          =   3570
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6297
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TrapTab         =   0   'False
      Tip             =   "frm_extRefPPT.frx":0000
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   -1  'True
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_extRefPPT.frx":0020
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_start 
      Height          =   645
      Left            =   960
      TabIndex        =   2
      Top             =   4080
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
      Caption         =   "frm_extRefPPT.frx":003C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_extRefPPT.frx":0070
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_extRefPPT.frx":0090
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   645
      Left            =   3600
      TabIndex        =   1
      Top             =   4080
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
      Caption         =   "frm_extRefPPT.frx":00AC
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_extRefPPT.frx":00D8
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_extRefPPT.frx":00F8
   End
   Begin HexUniControls.ctlUniFormCaption uniTitle 
      Left            =   240
      Top             =   4680
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_extRefPPT.frx":0114
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   960
      Top             =   4680
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
Attribute VB_Name = "frm_extRefPPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_selection As Integer

Private Sub cmd_cancel_Click()

  unity_main.errorstring = "External Reference Qualification Setup screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  m_selection = vbCancel
  Unload frm_extRefPPT
End Sub

Private Sub cmd_start_Click()
  Dim spcFilename As String
  Dim extRefSPCFilename As String
  Dim userReq As Integer

  unity_main.errorstring = "External Reference Qualification Setup screen Start Scan button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  ' Load external reference selection
  If (lst_refFileNames.ListIndex < 0) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_extrefPPT", "errMsg1", "Please select an external reference file"), vbExclamation
    Exit Sub
  Else
    spcFilename = lst_refFileNames.List(lst_refFileNames.ListIndex)
    
    If (frm_extRef.load_ext_ref_cfg_file(spcFilename, False) = False) Then
      Exit Sub
    Else
      ' Check if external reference qualification spectrum file already exists
      If (unity_main.check_ext_ref_ppt_file(spcFilename, extRefSPCFilename) = True) Then
        userReq = CWrap.ShowMessageBoxW(MLSupport.GGS_Params("frm_extRefPPT.statMsg1", "%1 external reference qualification file already exits for this wavelength range. Do you want to replace it?", extRefSPCFilename), vbYesNo)
          
        If (userReq = vbNo) Then
          Exit Sub
        End If
      End If
   
    End If
  End If
  
  unity_main.errorstring = ("User requested new External Reference Qualification file for " & spcFilename)
  unity_main.write_error (LOG_DBG_LEVEL1)
  
  unity_main.m_extRefPPTAdapterType = frm_extRef.m_extRefAdapterType
  unity_main.m_extRefPPTAdaptIndx = frm_extRef.m_extRefAdaptIndx
  unity_main.m_extRefPPTEndWvln = frm_extRef.m_extRefEndWvln
  unity_main.m_extRefPPTFileName = spcFilename
  unity_main.m_extRefPPTFileSetup = False
  unity_main.m_extRefPPTMultiCupIndx = frm_extRef.m_extRefMultiCupIndx
  unity_main.m_extRefPPTMultiCupType = frm_extRef.m_extRefMultiCupType
  unity_main.m_extRefPPTNScans = frm_extRef.m_extRefNScans
  unity_main.m_extRefPPTRotateDir = frm_extRef.m_extRefRotateDir
  unity_main.m_extRefPPTRotateIndexSteps = frm_extRef.m_extRefRotateIndexSteps
  unity_main.m_extRefPPTRotateMoveMode = frm_extRef.m_extRefRotateMoveMode
  unity_main.m_extRefPPTRotateSpeed = frm_extRef.m_extRefRotateSpeed
  unity_main.m_extRefPPTRotateStepSteps = frm_extRef.m_extRefRotateStepSteps
  unity_main.m_extRefPPTStartWvln = frm_extRef.m_extRefStartWvln
  unity_main.m_extRefPPTTrayNum = frm_extRef.m_extRefTrayNum
  
  m_selection = vbOK
  Unload frm_extRefPPT
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub






