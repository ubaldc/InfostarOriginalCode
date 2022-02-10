VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{488D84E9-A9A9-4C74-949B-F7752803ADEE}#99.2#0"; "EMDVirtualKeyBoard.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_numpad 
   Caption         =   "Number Entry"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
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
   ScaleHeight     =   6945
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   5160
      TabIndex        =   3
      Top             =   5400
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
      Caption         =   "frm_numpad.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_numpad.frx":002C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_numpad.frx":004C
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_accept 
      Default         =   -1  'True
      Height          =   650
      Left            =   5160
      TabIndex        =   2
      Top             =   4560
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
      Caption         =   "frm_numpad.frx":0068
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_numpad.frx":0094
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_numpad.frx":00B4
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_clear 
      Height          =   650
      Left            =   5160
      TabIndex        =   1
      Top             =   3720
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
      Caption         =   "frm_numpad.frx":00D0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_numpad.frx":00FA
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_numpad.frx":011A
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_num 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   0
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_numpad.frx":0136
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   0
      MultiLine       =   0   'False
      Alignment       =   1
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frm_numpad.frx":0156
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_numpad.frx":0176
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6120
      Top             =   1800
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6945
      FormDesignWidth =   7440
   End
   Begin VBoard_EMD.KeySet_Num KeySet_Num1 
      Height          =   4680
      Left            =   240
      Top             =   1800
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   8255
      LargeFontName   =   "Arial Unicode MS"
      LargeFontSize   =   6.75
      SmallFontName   =   "Arial Unicode MS"
      ButtonWidth     =   60
      ButtonHeight    =   60
      SendKeysEnabled =   -1  'True
      RegAtDT         =   -1  'True
   End
   Begin HexUniControls.ctlUniLabel lbl_num 
      Height          =   615
      Left            =   240
      Top             =   240
      Width           =   4335
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
      Caption         =   "frm_numpad.frx":0192
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_numpad.frx":01B2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_numpad.frx":01D2
   End
   Begin HexUniControls.ctlUniImage ctlUniImage1 
      Height          =   1215
      Left            =   5160
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2143
      Picture         =   "frm_numpad.frx":01EE
      Tip             =   "frm_numpad.frx":50F5
      Enabled         =   -1  'True
      Border          =   0   'False
      BackColor       =   -2147483633
      BorderColor     =   -1
      RoundedBorders  =   -1  'True
      Stretch         =   -1  'True
      XTransp         =   0
      YTransp         =   0
      MousePointer    =   0
      MouseIcon       =   "frm_numpad.frx":5115
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   6120
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
      Left            =   6120
      Top             =   2400
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_numpad.frx":5131
   End
End
Attribute VB_Name = "frm_numpad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_cancel_Click()

  unity_main.errorstring = "Number Entry screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Unload frm_numpad
End Sub

Private Sub cmd_accept_Click()
  Dim lennum As Integer
  Dim onechar As String
  Dim rebuildit As String
  Dim wasf As String
  Dim fieldcounter As Integer
  Dim buildcounter As Integer

  unity_main.errorstring = "Number Entry screen Accept button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  On Error GoTo problemwithentry
  
  If (Trim(txt_num.Text) = "") Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_numpad", "errMsg1", "You must enter a number!"), vbOKOnly
    Exit Sub
  End If

  If (IsNumeric(Trim(txt_num.Text)) = False) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_numpad", "errMsg1", "You must enter a number!"), vbOKOnly
    Exit Sub
  End If

  wasf = Trim(frm_numpad.txt_num.Text)
  lennum = Len(wasf)
  rebuildit = ""
  
  For buildcounter = 1 To lennum
    onechar = Mid(wasf, buildcounter, 1)
    
    If (onechar = ",") Then
      onechar = "."
    End If
    
    rebuildit = rebuildit & onechar
  Next buildcounter
  
  frm_numpad.txt_num.Text = rebuildit
  
  ' Determine which form the data input is for
  Select Case (unity_main.formfrom)
    Case 2              ' frm_collect
      ' Determine which variable the data input is for
      Select Case unity_main.varfrom
        Case 1
          frm_collect.numInc_smplNScans.Text = txt_num.Text
        Case 2
          frm_collect.numInc_smplPPT.Text = txt_num.Text
        Case 3
          frm_collect.numInc_numRepacks.Text = txt_num.Text
        Case 4
          frm_collect.txt_startWvln.Text = txt_num.Text
        Case 5
          frm_collect.txt_endWvln.Text = txt_num.Text
        Case 6
          frm_collect.numInc_rotateIndexSteps.Text = txt_num.Text
        Case 7
          frm_collect.numInc_rotateSpeed.Text = txt_num.Text
        Case 8
          frm_collect.numInc_rotateStepSteps.Text = txt_num.Text
        Case 9
          frm_collect.numInc_dateCounter.Text = txt_num.Text
        Case 10
          frm_collect.numInc_nameCounter.Text = txt_num.Text
        Case 11
          frm_collect.txt_noOLVal.Text = txt_num.Text
#If ABBFT Then
        Case 12
          frm_collect.numInc_numSamples.Text = txt_num.Text
        Case 13
          frm_collect.txt_delayStart.Text = txt_num.Text
        Case 14
          frm_collect.txt_delayMeasure.Text = txt_num.Text
        Case 15
          frm_collect.txt_startWavenumIndx.Text = txt_num.Text
        Case 16
          frm_collect.txt_endWavenumIndx.Text = txt_num.Text
#End If
      End Select

    Case 3                ' frm_1model
      ' Determine which variable the data input is for
      Select Case unity_main.varfrom
        Case 2
          frm_1model.txt_modvar(2).Text = txt_num.Text
        Case 3
          frm_1model.txt_modvar(3).Text = txt_num.Text
        Case 4
          frm_1model.txt_modvar(4).Text = txt_num.Text
        Case 5
          frm_1model.txt_modvar(5).Text = txt_num.Text
        Case 6
          frm_1model.txt_modvar(6).Text = txt_num.Text
        Case 7
          frm_1model.txt_modvar(7).Text = txt_num.Text
        Case 8
          frm_1model.txt_modvar(8).Text = txt_num.Text
        Case 9
          frm_1model.txt_modvar(9).Text = txt_num.Text
        Case 10
          frm_1model.txt_modvar(10).Text = txt_num.Text
        Case 11
          frm_1model.txt_modvar(11).Text = txt_num.Text
        Case 12
          frm_1model.txt_modvar(12).Text = txt_num.Text
        Case 13
          frm_1model.txt_modvar(13).Text = txt_num.Text
      End Select

    Case 4                ' frm_inst
      ' Determine which variable the data input is for
      Select Case unity_main.varfrom
        Case 1
          frm_Inst.numInc_refTimeout.Text = txt_num.Text
        Case 2
          frm_Inst.numInc_refPPT.Text = txt_num.Text
        Case 3
          frm_Inst.numInc_refNScans.Text = txt_num.Text
        Case 4
          frm_Inst.numInc_dbgLevel.Text = txt_num.Text
        Case 5
          frm_Inst.numInc_globalDateCtr.Text = txt_num.Text
        Case 6
          frm_Inst.numInc_globalBaseCtr.Text = txt_num.Text
#If ABBFT Then
        Case 7
          frm_Inst.numInc_refTempDiff.Text = txt_num.Text
        Case 8
          frm_Inst.txt_portNum.Text = txt_num.Text
        Case 9
          frm_Inst.numInc_rspTimeout.Text = txt_num.Text
        Case 11
          frm_Inst.txt_ipAddr(1).Text = txt_num.Text
        Case 12
          frm_Inst.txt_ipAddr(2).Text = txt_num.Text
        Case 13
          frm_Inst.txt_ipAddr(3).Text = txt_num.Text
        Case 14
          frm_Inst.txt_ipAddr(4).Text = txt_num.Text
#Else
        Case 15
          frm_Inst.numInc_commPort.Text = txt_num.Text
        Case 16
          frm_Inst.numInc_refVerifyTimeout.Text = txt_num.Text
#End If
      End Select
      
    Case 5                ' frm_secmodel
      ' Determine which variable the data input is for
      Select Case unity_main.varfrom
        Case 1
          frm_secmodel.txt_value.Text = txt_num.Text
      End Select

    Case 6   ' frmMain
      ' Determine which variable the data input is for
      Select Case unity_main.varfrom
        Case 3
          frmMain.txt_modvar(3).Text = txt_num.Text
        Case 4
          frmMain.txt_modvar(4).Text = txt_num.Text
        Case 5
          frmMain.txt_modvar(5).Text = txt_num.Text
        Case 6
          frmMain.txt_modvar(6).Text = txt_num.Text
        Case 7
          frmMain.txt_modvar(7).Text = txt_num.Text
        Case 8
          frmMain.txt_modvar(8).Text = txt_num.Text
        Case 9
          frmMain.txt_modvar(9).Text = txt_num.Text
        Case 10
          frmMain.txt_modvar(10).Text = txt_num.Text
        Case 11
          frmMain.txt_modvar(11).Text = txt_num.Text
        Case 12
          frmMain.txt_modvar(12).Text = txt_num.Text
        Case 13
          frmMain.txt_modvar(13).Text = txt_num.Text
      End Select
      
    Case 8              ' frm_refPPT
      ' Determine which variable the data input is for
      Select Case unity_main.varfrom
        Case 1
          frm_refPPT.txt_startWvln.Text = txt_num.Text
        Case 2
          frm_refPPT.txt_endWvln.Text = txt_num.Text
      End Select
      
    Case 14   ' frm_ticket
      ' Determine which variable the data input is for
      Select Case unity_main.varfrom
        Case 3
          frm_ticket.numInc_fontSize.Text = txt_num.Text
        Case 4
          frm_ticket.numInc_preLFs.Text = txt_num.Text
        Case 5
          frm_ticket.numInc_postLFs.Text = txt_num.Text
      End Select
      
#If SSTAR Then
    Case 16              ' frm_extRef
      ' Determine which variable the data input is for
      Select Case unity_main.varfrom
        Case 1
          frm_extRef.numInc_refTimeout.Text = txt_num.Text
        Case 2
          frm_extRef.numInc_refPPT.Text = txt_num.Text
        Case 3
          frm_extRef.numInc_refNScans.Text = txt_num.Text
        Case 4
          frm_extRef.txt_startWvln.Text = txt_num.Text
        Case 5
          frm_extRef.txt_endWvln.Text = txt_num.Text
        Case 6
          frm_extRef.numInc_rotateIndexSteps.Text = txt_num.Text
        Case 7
          frm_extRef.numInc_rotateSpeed.Text = txt_num.Text
        Case 8
          frm_extRef.numInc_rotateStepSteps.Text = txt_num.Text
      End Select
#End If
      
    Case 17   ' frm_prd
      ' Determine which variable the data input is for
      Select Case unity_main.varfrom
        Case 3
          frm_prd.txt_modvar(3).Text = txt_num.Text
        Case 4
          frm_prd.txt_modvar(4).Text = txt_num.Text
        Case 5
          frm_prd.txt_modvar(5).Text = txt_num.Text
        Case 6
          frm_prd.txt_modvar(6).Text = txt_num.Text
        Case 7
          frm_prd.txt_modvar(7).Text = txt_num.Text
        Case 8
          frm_prd.txt_modvar(8).Text = txt_num.Text
        Case 9
          frm_prd.txt_modvar(9).Text = txt_num.Text
        Case 10
          frm_prd.txt_modvar(10).Text = txt_num.Text
        Case 11
          frm_prd.txt_modvar(11).Text = txt_num.Text
        Case 12
          frm_prd.txt_modvar(12).Text = txt_num.Text
        Case 13
          frm_prd.txt_modvar(13).Text = txt_num.Text
      End Select
      
    Case 18   ' frm_dynRptCfg
      ' Determine which variable the data input is for
      Select Case unity_main.varfrom
        Case 5
          frm_dynRptCfg.numInc_baseCounter.Text = txt_num.Text
        Case 6
          frm_dynRptCfg.numInc_dateCounter.Text = txt_num.Text
        Case 7
          frm_dynRptCfg.numInc_hdrNumFields.Text = txt_num.Text
        Case 8
          frm_dynRptCfg.numInc_usrNumFields.Text = txt_num.Text
        Case 9
          frm_dynRptCfg.numInc_recNumFields.Text = txt_num.Text
        Case 10
          frm_dynRptCfg.numInc_trlNumFields.Text = txt_num.Text
        Case 11
          frm_dynRptCfg.numInc_hdrMaxChars(1).Text = txt_num.Text
        Case 12
          frm_dynRptCfg.numInc_hdrMaxChars(2).Text = txt_num.Text
        Case 13
          frm_dynRptCfg.numInc_hdrMaxChars(3).Text = txt_num.Text
        Case 14
          frm_dynRptCfg.numInc_hdrMaxChars(4).Text = txt_num.Text
        Case 15
          frm_dynRptCfg.numInc_hdrMaxChars(5).Text = txt_num.Text
        Case 16
          frm_dynRptCfg.numInc_hdrMaxChars(6).Text = txt_num.Text
        Case 17
          frm_dynRptCfg.numInc_hdrMaxChars(7).Text = txt_num.Text
        Case 18
          frm_dynRptCfg.numInc_hdrMaxChars(8).Text = txt_num.Text
        Case 19
          frm_dynRptCfg.numInc_hdrMaxChars(9).Text = txt_num.Text
        Case 21
          frm_dynRptCfg.numInc_usrMaxChars(1).Text = txt_num.Text
        Case 22
          frm_dynRptCfg.numInc_usrMaxChars(2).Text = txt_num.Text
        Case 23
          frm_dynRptCfg.numInc_usrMaxChars(3).Text = txt_num.Text
        Case 24
          frm_dynRptCfg.numInc_usrMaxChars(4).Text = txt_num.Text
        Case 25
          frm_dynRptCfg.numInc_usrMaxChars(5).Text = txt_num.Text
        Case 26
          frm_dynRptCfg.numInc_usrMaxChars(6).Text = txt_num.Text
        Case 27
          frm_dynRptCfg.numInc_usrMaxChars(7).Text = txt_num.Text
        Case 28
          frm_dynRptCfg.numInc_usrMaxChars(8).Text = txt_num.Text
        Case 29
          frm_dynRptCfg.numInc_usrMaxChars(9).Text = txt_num.Text
        Case 31
          frm_dynRptCfg.numInc_recMaxChars(1).Text = txt_num.Text
        Case 32
          frm_dynRptCfg.numInc_recMaxChars(2).Text = txt_num.Text
        Case 33
          frm_dynRptCfg.numInc_recMaxChars(3).Text = txt_num.Text
        Case 34
          frm_dynRptCfg.numInc_recMaxChars(4).Text = txt_num.Text
        Case 35
          frm_dynRptCfg.numInc_recMaxChars(5).Text = txt_num.Text
        Case 36
          frm_dynRptCfg.numInc_recMaxChars(6).Text = txt_num.Text
        Case 37
          frm_dynRptCfg.numInc_recMaxChars(7).Text = txt_num.Text
        Case 38
          frm_dynRptCfg.numInc_recMaxChars(8).Text = txt_num.Text
        Case 39
          frm_dynRptCfg.numInc_recMaxChars(9).Text = txt_num.Text
        Case 41
          frm_dynRptCfg.numInc_trlMaxChars(1).Text = txt_num.Text
        Case 42
          frm_dynRptCfg.numInc_trlMaxChars(2).Text = txt_num.Text
        Case 43
          frm_dynRptCfg.numInc_trlMaxChars(3).Text = txt_num.Text
        Case 44
          frm_dynRptCfg.numInc_trlMaxChars(4).Text = txt_num.Text
        Case 45
          frm_dynRptCfg.numInc_trlMaxChars(5).Text = txt_num.Text
        Case 46
          frm_dynRptCfg.numInc_trlMaxChars(6).Text = txt_num.Text
        Case 47
          frm_dynRptCfg.numInc_trlMaxChars(7).Text = txt_num.Text
        Case 48
          frm_dynRptCfg.numInc_trlMaxChars(8).Text = txt_num.Text
        Case 49
          frm_dynRptCfg.numInc_trlMaxChars(9).Text = txt_num.Text
      End Select
      
#If SSTAR Then
    Case 19   ' frm_spectTreatCfg
      ' Determine which variable the data input is for
      Select Case unity_main.varfrom
        Case 1
          frm_spectTreatCfg.numInc_smoothNumPts.Text = txt_num.Text
        Case 2
          frm_spectTreatCfg.numInc_startSmoothNumPts.Text = txt_num.Text
        Case 3
          frm_spectTreatCfg.numInc_endSmoothNumPts.Text = txt_num.Text
        Case 4
          frm_spectTreatCfg.numInc_progSmoothRate.Text = txt_num.Text
        Case 5
          frm_spectTreatCfg.txt_startSmoothWvln.Text = txt_num.Text
        Case 6
          frm_spectTreatCfg.txt_endSmoothWvln.Text = txt_num.Text
      End Select
#End If

#If SSRCS Then
    Case 20   ' frm_ssrcsConnect
      ' Determine which variable the data input is for
      Select Case unity_main.varfrom
        Case 2
          frm_ssrcsConnect.numInc_connectTimeout.Text = txt_num.Text
        Case 3
          frm_ssrcsConnect.numInc_rspTimeout.Text = txt_num.Text
      End Select
#End If
  End Select
  
  Unload frm_numpad
  Exit Sub
  
problemwithentry:
  CWrap.ShowMessageBoxW MLSupport.GSS("frm_numpad", "errMsg2", "Problem in number pad entry or code, please try again, if it occurs again please contact Unity!"), vbCritical
End Sub

Private Sub cmd_clear_Click()
  
  unity_main.errorstring = "Number Entry screen Clear button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  txt_num.Text = ""
  txt_num.SetFocus
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub








