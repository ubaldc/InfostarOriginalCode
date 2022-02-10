VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_refPPT 
   Caption         =   "Internal Reference Qualification Setup"
   ClientHeight    =   3225
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
   ScaleHeight     =   3225
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   240
      Top             =   1680
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3225
      FormDesignWidth =   6600
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_maxWvln 
      Height          =   360
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
      Width           =   1065
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      Enabled         =   0   'False
      Locked          =   0   'False
      Text            =   "frm_refPPT.frx":0000
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
      MultiLine       =   0   'False
      Alignment       =   1
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frm_refPPT.frx":0020
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_refPPT.frx":0040
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_minWvln 
      Height          =   360
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   1065
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      Enabled         =   0   'False
      Locked          =   0   'False
      Text            =   "frm_refPPT.frx":005C
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
      MultiLine       =   0   'False
      Alignment       =   1
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frm_refPPT.frx":007C
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_refPPT.frx":009C
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_start 
      Height          =   650
      Left            =   960
      TabIndex        =   3
      Top             =   2400
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
      Caption         =   "frm_refPPT.frx":00B8
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_refPPT.frx":00EC
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_refPPT.frx":010C
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   3600
      TabIndex        =   0
      Top             =   2400
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
      Caption         =   "frm_refPPT.frx":0128
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_refPPT.frx":0154
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_refPPT.frx":0174
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_startWvln 
      Height          =   330
      Left            =   1560
      TabIndex        =   1
      Top             =   315
      Width           =   1065
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_refPPT.frx":0190
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
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
      Tip             =   "frm_refPPT.frx":01B0
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_refPPT.frx":01D0
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_endWvln 
      Height          =   360
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   1065
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_refPPT.frx":01EC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
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
      Tip             =   "frm_refPPT.frx":020C
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_refPPT.frx":022C
   End
   Begin HexUniControls.ctlUniLabel lbl_startWvln 
      Height          =   330
      Left            =   2880
      Top             =   360
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_refPPT.frx":0248
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_refPPT.frx":028E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_refPPT.frx":02AE
   End
   Begin HexUniControls.ctlUniLabel lbl_endWvln 
      Height          =   330
      Left            =   2880
      Top             =   840
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_refPPT.frx":02CA
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_refPPT.frx":030C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_refPPT.frx":032C
   End
   Begin HexUniControls.ctlUniLabel lbl_minWvln 
      Height          =   330
      Left            =   2880
      Top             =   1320
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_refPPT.frx":0348
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_refPPT.frx":038C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_refPPT.frx":03AC
   End
   Begin HexUniControls.ctlUniLabel lbl_maxWvln 
      Height          =   330
      Left            =   2880
      Top             =   1800
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_refPPT.frx":03C8
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_refPPT.frx":040C
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_refPPT.frx":042C
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   240
      Top             =   2280
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
      Top             =   1080
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_refPPT.frx":0448
   End
End
Attribute VB_Name = "frm_refPPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_selection As Integer

Private Sub cmd_cancel_Click()
  
  unity_main.errorstring = "Internal Reference Qualification Setup screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  m_selection = vbCancel
  Unload frm_refPPT
End Sub

Private Sub cmd_start_Click()
  Dim startWvln As Double
  Dim endWvln As Double
  Dim rc As Boolean
  Dim spcFilename As String
  Dim userReq As Integer

  unity_main.errorstring = "Internal Reference Qualification Setup screen Start Scan button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  txt_startWvln.Text = Trim(txt_startWvln.Text)
  txt_endWvln.Text = Trim(txt_endWvln.Text)

  On Error GoTo BAD_VALUE
  startWvln = CDbl(txt_startWvln.Text)
  endWvln = CDbl(txt_endWvln.Text)

  ' Check for valid wavelength range
  If (startWvln < txt_minWvln.Text) Or (startWvln > txt_maxWvln.Text) Then
    CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_refPPT.errMsg1", "Please enter a Starting Wavelength value between %1 and %2", txt_minWvln.Text, txt_maxWvln.Text), vbExclamation
    Exit Sub
  Else
    If (endWvln < txt_minWvln.Text) Or (endWvln > txt_maxWvln.Text) Then
      CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_refPPT.errMsg2", "Please enter a Ending Wavelength value between %1 and %2", txt_minWvln.Text, txt_maxWvln.Text), vbExclamation
      Exit Sub
    Else
      If (endWvln <= startWvln) Then
        CWrap.ShowMessageBoxW MLSupport.GSS("frm_refPPT", "errMsg1", "Please enter a Ending Wavelength value greater than the Starting Wavelength value"), vbExclamation
        Exit Sub
      Else
        ' Check wavelengths with instrument's table
#If SSRCS Then
        SSRCSClientError = unity_main.SSRCSClient.ChkWvlnRange(startWvln, endWvln, rc)
#Else
        rc = unity_main.MS11srv.ChkWvlnRange(startWvln, endWvln)
#End If
        ' Check if wavelengths changed by instrument, if so ask user if okay
        If (rc = False) Then
          userReq = CWrap.ShowMessageBoxW(MLSupport.GGS_Params("frm_refPPT.statMsg1", "Wavelengths were changed to %1 - %2. Do you want use these values?", CStr(startWvln), CStr(endWvln)), vbYesNo)
          
          If (userReq = vbNo) Then
            Exit Sub
          Else
            txt_startWvln.Text = startWvln
            txt_endWvln.Text = endWvln
          End If
        End If
        
        ' Check if internal reference qualification spectrum file already exists
        If (unity_main.check_int_ref_ppt_file(startWvln, endWvln, spcFilename) = True) Then
          userReq = CWrap.ShowMessageBoxW(MLSupport.GGS_Params("frm_refPPT.statMsg2", "%1 internal reference qualification file already exits for this wavelength range. Do you want to replace it?", spcFilename), vbYesNo)
          
          If (userReq = vbNo) Then
            Exit Sub
          End If
        End If
      End If
    End If
  End If
  
  unity_main.errorstring = ("User requested new Internal Reference Qualification file over wavelength range " & startWvln & "-" & endWvln)
  unity_main.write_error (LOG_DBG_LEVEL1)
  unity_main.m_intRefPPTStartWvln = startWvln
  unity_main.m_intRefPPTEndWvln = endWvln
  m_selection = vbOK
  Unload frm_refPPT
  Exit Sub
  
BAD_VALUE:
  CWrap.ShowMessageBoxW MLSupport.GSS("frm_refPPT", "errMsg2", "Please enter a valid number for Starting and/or Ending Wavelength"), vbExclamation
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub

Private Sub txt_endWvln_DblCLick(Button As Integer)
  
  unity_main.formfrom = 8
  unity_main.varfrom = 2
  frm_numpad.lbl_num.Caption = lbl_endWvln.Caption
  frm_numpad.txt_num.Text = txt_endWvln.Text
  frm_numpad.Show 1
End Sub

Private Sub txt_startWvln_DblCLick(Button As Integer)
  
  unity_main.formfrom = 8
  unity_main.varfrom = 1
  frm_numpad.lbl_num.Caption = lbl_startWvln.Caption
  frm_numpad.txt_num.Text = txt_startWvln.Text
  frm_numpad.Show 1
End Sub








