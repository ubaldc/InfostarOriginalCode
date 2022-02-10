VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Object = "{488D84E9-A9A9-4C74-949B-F7752803ADEE}#99.2#0"; "EMDVirtualKeyBoard.ocx"
Begin VB.Form frm_kybd 
   Caption         =   "Keyboard Input"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8460
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
   Icon            =   "frm_kybd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   8460
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniButtonImageXP cmd_clear 
      Height          =   650
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
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
      Caption         =   "frm_kybd.frx":0442
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_kybd.frx":046C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_kybd.frx":048C
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_accept 
      Default         =   -1  'True
      Height          =   650
      Left            =   6240
      TabIndex        =   2
      Top             =   1080
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
      Caption         =   "frm_kybd.frx":04A8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_kybd.frx":04D4
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_kybd.frx":04F4
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   6240
      TabIndex        =   3
      Top             =   1800
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
      Caption         =   "frm_kybd.frx":0510
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_kybd.frx":053C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_kybd.frx":055C
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5400
      Top             =   360
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5910
      FormDesignWidth =   8460
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_kybd 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4935
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_kybd.frx":0578
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   0
      MultiLine       =   -1  'True
      Alignment       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frm_kybd.frx":0598
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_kybd.frx":05B8
   End
   Begin VBoard_EMD.KeySet_STD KeySet_STD1 
      Height          =   2805
      Left            =   120
      Top             =   2880
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   4948
      LargeFontName   =   "Arial Unicode MS"
      LargeFontSize   =   6.75
      SmallFontName   =   "Arial Unicode MS"
      ButtonWidth     =   35
      ButtonHeight    =   35
      SendKeysEnabled =   -1  'True
      RegAtDT         =   -1  'True
   End
   Begin HexUniControls.ctlUniLabel lbl_kybd 
      Height          =   615
      Left            =   120
      Top             =   240
      Width           =   4935
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_kybd.frx":05D4
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_kybd.frx":05F4
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_kybd.frx":0614
   End
   Begin HexUniControls.ctlUniImage ctlUniImage1 
      Height          =   960
      Left            =   6480
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1693
      Picture         =   "frm_kybd.frx":0630
      Tip             =   "frm_kybd.frx":5537
      Enabled         =   -1  'True
      Border          =   0   'False
      BackColor       =   -2147483633
      BorderColor     =   -1
      RoundedBorders  =   -1  'True
      Stretch         =   -1  'True
      XTransp         =   0
      YTransp         =   0
      MousePointer    =   0
      MouseIcon       =   "frm_kybd.frx":5557
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   5640
      Top             =   960
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
      Left            =   5880
      Top             =   360
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_kybd.frx":5573
   End
End
Attribute VB_Name = "frm_kybd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_cancel_Click()

  unity_main.errorstring = "Keyboard Input screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Unload frm_kybd
End Sub

Private Sub cmd_accept_Click()
  Dim jpgName As String
  Dim fileExt As String

  unity_main.errorstring = "Keyboard Input screen Accept button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  On Error GoTo donehere
  If (unity_main.formfrom = 1) Then  'unity main form
    Select Case unity_main.varfrom
      Case 1
        unity_main.txtsampcomment.Text = txt_kybd.Text
      Case 2
        unity_main.txtsamplename.Text = txt_kybd.Text
    End Select
  End If
  
  If (unity_main.formfrom = 2) Then 'frm_collect
    Select Case unity_main.varfrom
      Case 1
        frm_collect.txtsampnamebase.Text = txt_kybd.Text
      Case 3
        frm_collect.txt_caldir.Text = txt_kybd.Text
      Case 5
        frm_collect.txtpredfile.Text = txt_kybd.Text
      Case 6
        frm_collect.txt_csvfilename.Text = txt_kybd.Text
      Case 7
        frm_collect.txt_noPredVal.Text = txt_kybd.Text
    End Select
  End If
  
  If (unity_main.formfrom = 3) Then 'frm_1model
    Select Case unity_main.varfrom
      Case 0
        frm_1model.txt_modvar(0).Text = txt_kybd.Text
      Case 1
        frm_1model.txt_modvar(1).Text = txt_kybd.Text
    End Select
  End If
  
  If (unity_main.formfrom = 4) Then 'frm_Inst
    Select Case unity_main.varfrom
      Case 1
        frm_Inst.txt_globalSampNameBase.Text = txt_kybd.Text
      Case 2
        frm_Inst.txt_batchRptPath.Text = txt_kybd.Text
    End Select
  End If
  
  If (unity_main.formfrom = 5) Then
    Select Case unity_main.varfrom
      Case 1
        frm_secmodel.txt_modinfo.Text = txt_kybd.Text
      Case 2
        frm_secmodel.txt_cprop.Text = txt_kybd.Text
      Case 3
        frm_secmodel.txt_mname2.Text = txt_kybd.Text
    End Select
  End If
  
  If (unity_main.formfrom = 9) Then
    Select Case unity_main.varfrom
      Case 1
        frmProduct.txt_product.Text = Trim(txt_kybd.Text)
      Case 2
        frmProduct.txt_sampmode.Text = Trim(txt_kybd.Text)
      Case 3
        frmProduct.txt_inifile.Text = Trim(txt_kybd.Text)
    End Select
  End If

  If (unity_main.formfrom = 10) Then 'graph export to jpg
    jpgName = Trim(txt_kybd.Text)
    
    If (jpgName = "") Then
      Exit Sub
    End If
    
    fileExt = ("." & LCase(CFile.st_FileExt(jpgName)))
    
    If (fileExt <> SPECTRA_IMAGE_FILE_EXT) Then
      jpgName = CFile.st_FileNameNoExt(jpgName) & SPECTRA_IMAGE_FILE_EXT
    End If
    
    frm_scan.jpg_toexport = jpgName
    frm_kybd.lbl_kybd.Caption = MLSupport.GSS("frm_kybd", "statMsg1", "Please wait, exporting graph")
  End If
  
  If (unity_main.formfrom = 11) Then 'report form
    frmReport.Text1.Text = txt_kybd.Text
  End If
  
  If (unity_main.formfrom = 12) Then   'pog configuration form
    Select Case unity_main.varfrom
      Case 1
        frm_POG.txt_outpath.Text = Trim(txt_kybd.Text)
      Case 3
        frm_POG.txt_pogfilename.Text = Trim(txt_kybd.Text)
      Case 4
        frm_POG.txt_limsin.Text = Trim(txt_kybd.Text)
      Case 5
        frm_POG.txt_limsinpath.Text = Trim(txt_kybd.Text)
    End Select
  End If
  
  If (unity_main.formfrom = 13) Then   'calstar model setup screen
    Select Case unity_main.varfrom
      Case 1
        frmMain.txt_propname.Text = Trim(txt_kybd.Text)
    End Select
  End If
  
  If (unity_main.formfrom = 14) Then   'printer ticket setup screen
    Select Case unity_main.varfrom
      Case 1
        frm_ticket.txt_header1.Text = Trim(txt_kybd.Text)
      Case 2
        frm_ticket.txt_header2.Text = Trim(txt_kybd.Text)
    End Select
  End If
  
  If (unity_main.formfrom = 15) Then
    frm_guilevel.txt_pw.Text = Trim(txt_kybd.Text)
  End If
  
#If SSTAR Then
  If (unity_main.formfrom = 16) Then  ' external reference configuration
    Select Case unity_main.varfrom
      Case 1
        frm_extRef.txt_extRefName.Text = Trim(txt_kybd.Text)
      Case 2
        frm_extRef.txt_extRefDesc.Text = Trim(txt_kybd.Text)
    End Select
  End If
#End If
  
  If (unity_main.formfrom = 17) Then   'PRD model setup screen
    Select Case unity_main.varfrom
      Case 1
        frm_prd.txt_propname.Text = Trim(txt_kybd.Text)
      Case 2
        frm_prd.txt_stfName.Text = Trim(txt_kybd.Text)
    End Select
  End If
  
  If (unity_main.formfrom = 18) Then   'Dynamic report setup screen
    Select Case unity_main.varfrom
      Case 0
        frm_dynRptCfg.txt_filePath.Text = Trim(txt_kybd.Text)
      Case 1
        frm_dynRptCfg.txt_fileExt.Text = Trim(txt_kybd.Text)
      Case 2
        frm_dynRptCfg.txt_baseName.Text = Trim(txt_kybd.Text)
      Case 3
        frm_dynRptCfg.txt_manualPrefix.Text = Trim(txt_kybd.Text)
      Case 4
        frm_dynRptCfg.txt_manualSuffix.Text = Trim(txt_kybd.Text)
      Case 51
        frm_dynRptCfg.txt_hdrFieldTxt(1).Text = Trim(txt_kybd.Text)
      Case 52
        frm_dynRptCfg.txt_hdrFieldTxt(2).Text = Trim(txt_kybd.Text)
      Case 53
        frm_dynRptCfg.txt_hdrFieldTxt(3).Text = Trim(txt_kybd.Text)
      Case 54
        frm_dynRptCfg.txt_hdrFieldTxt(4).Text = Trim(txt_kybd.Text)
      Case 55
        frm_dynRptCfg.txt_hdrFieldTxt(5).Text = Trim(txt_kybd.Text)
      Case 56
        frm_dynRptCfg.txt_hdrFieldTxt(6).Text = Trim(txt_kybd.Text)
      Case 57
        frm_dynRptCfg.txt_hdrFieldTxt(7).Text = Trim(txt_kybd.Text)
      Case 58
        frm_dynRptCfg.txt_hdrFieldTxt(8).Text = Trim(txt_kybd.Text)
      Case 59
        frm_dynRptCfg.txt_hdrFieldTxt(9).Text = Trim(txt_kybd.Text)
      Case 61
        frm_dynRptCfg.txt_usrFieldTxt(1).Text = Trim(txt_kybd.Text)
      Case 62
        frm_dynRptCfg.txt_usrFieldTxt(2).Text = Trim(txt_kybd.Text)
      Case 63
        frm_dynRptCfg.txt_usrFieldTxt(3).Text = Trim(txt_kybd.Text)
      Case 64
        frm_dynRptCfg.txt_usrFieldTxt(4).Text = Trim(txt_kybd.Text)
      Case 65
        frm_dynRptCfg.txt_usrFieldTxt(5).Text = Trim(txt_kybd.Text)
      Case 66
        frm_dynRptCfg.txt_usrFieldTxt(6).Text = Trim(txt_kybd.Text)
      Case 67
        frm_dynRptCfg.txt_usrFieldTxt(7).Text = Trim(txt_kybd.Text)
      Case 68
        frm_dynRptCfg.txt_usrFieldTxt(8).Text = Trim(txt_kybd.Text)
      Case 69
        frm_dynRptCfg.txt_usrFieldTxt(9).Text = Trim(txt_kybd.Text)
      Case 71
        frm_dynRptCfg.txt_recFieldTxt(1).Text = Trim(txt_kybd.Text)
      Case 72
        frm_dynRptCfg.txt_recFieldTxt(2).Text = Trim(txt_kybd.Text)
      Case 73
        frm_dynRptCfg.txt_recFieldTxt(3).Text = Trim(txt_kybd.Text)
      Case 74
        frm_dynRptCfg.txt_recFieldTxt(4).Text = Trim(txt_kybd.Text)
      Case 75
        frm_dynRptCfg.txt_recFieldTxt(5).Text = Trim(txt_kybd.Text)
      Case 76
        frm_dynRptCfg.txt_recFieldTxt(6).Text = Trim(txt_kybd.Text)
      Case 77
        frm_dynRptCfg.txt_recFieldTxt(7).Text = Trim(txt_kybd.Text)
      Case 78
        frm_dynRptCfg.txt_recFieldTxt(8).Text = Trim(txt_kybd.Text)
      Case 79
        frm_dynRptCfg.txt_recFieldTxt(9).Text = Trim(txt_kybd.Text)
      Case 81
        frm_dynRptCfg.txt_trlFieldTxt(1).Text = Trim(txt_kybd.Text)
      Case 82
        frm_dynRptCfg.txt_trlFieldTxt(2).Text = Trim(txt_kybd.Text)
      Case 83
        frm_dynRptCfg.txt_trlFieldTxt(3).Text = Trim(txt_kybd.Text)
      Case 84
        frm_dynRptCfg.txt_trlFieldTxt(4).Text = Trim(txt_kybd.Text)
      Case 85
        frm_dynRptCfg.txt_trlFieldTxt(5).Text = Trim(txt_kybd.Text)
      Case 86
        frm_dynRptCfg.txt_trlFieldTxt(6).Text = Trim(txt_kybd.Text)
      Case 87
        frm_dynRptCfg.txt_trlFieldTxt(7).Text = Trim(txt_kybd.Text)
      Case 88
        frm_dynRptCfg.txt_trlFieldTxt(8).Text = Trim(txt_kybd.Text)
      Case 89
        frm_dynRptCfg.txt_trlFieldTxt(9).Text = Trim(txt_kybd.Text)
    End Select
  End If
  
#If SSRCS Then
  If (unity_main.formfrom = 20) Then   ' SpectraStar RCS connect screen
    Select Case unity_main.varfrom
      Case 1
        frm_ssrcsConnect.txt_ipAddr.Text = Trim(txt_kybd.Text)
    End Select
  End If
#End If
  
donehere:
  Unload frm_kybd
End Sub

Private Sub cmd_clear_Click()
  
  unity_main.errorstring = "Keyboard Input screen Clear button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  txt_kybd.Text = ""
  txt_kybd.SetFocus
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub








