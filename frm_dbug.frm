VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_dbug 
   Caption         =   "Scan Log"
   ClientHeight    =   9015
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9945
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
   Icon            =   "frm_dbug.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr_dp 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1320
      Top             =   8160
   End
   Begin HexUniControls.ctlUniTextBoxXP txt_drawerPos 
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8190
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      BorderColor     =   -2147483633
      BackColor       =   -2147483633
      ForeColor       =   32768
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frm_dbug.frx":030A
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
      Alignment       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frm_dbug.frx":0338
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_dbug.frx":0358
   End
   Begin HexUniControls.ctlUniListBoxXP lst_status 
      Height          =   7215
      Left            =   5400
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   12726
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
      Tip             =   "frm_dbug.frx":0374
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   -1  'True
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_dbug.frx":0394
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin FPUSpreadADO.fpSpread ss_dbug 
      Height          =   7215
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   5055
      _Version        =   458752
      _ExtentX        =   8916
      _ExtentY        =   12726
      _StockProps     =   64
      ColHeaderDisplay=   0
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
      MaxRows         =   100
      OperationMode   =   1
      SelectBlockOptions=   0
      SpreadDesigner  =   "frm_dbug.frx":03B0
      UserResize      =   1
   End
   Begin HexUniControls.ctlUniLabel lbl_drawerPos 
      Height          =   375
      Left            =   2880
      Top             =   8160
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "frm_dbug.frx":0616
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   1
      VAlignment      =   1
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frm_dbug.frx":0654
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frm_dbug.frx":0674
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   240
      Top             =   7560
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9015
      FormDesignWidth =   9945
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_exit 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   7080
      TabIndex        =   0
      Top             =   7440
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
      Caption         =   "frm_dbug.frx":0690
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_dbug.frx":06B8
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_dbug.frx":06D8
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_testDrawer 
      Height          =   650
      Left            =   3975
      TabIndex        =   2
      Top             =   7440
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
      Caption         =   "frm_dbug.frx":06F4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_dbug.frx":072A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_dbug.frx":074A
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_refresh 
      Height          =   650
      Left            =   885
      TabIndex        =   1
      Top             =   7440
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
      Caption         =   "frm_dbug.frx":0766
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_dbug.frx":07A2
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_dbug.frx":07C2
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   840
      Top             =   8160
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
      Top             =   8040
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_dbug.frx":07DE
   End
End
Attribute VB_Name = "frm_dbug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_exit_Click()
  
  unity_main.errorstring = "Scan Log screen Exit button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  frm_dbug.tmr_dp.enabled = False
  Unload frm_dbug
End Sub

#If SSTAR Then
Private Sub cmd_testDrawer_Click()

  unity_main.errorstring = "Scan Log screen Test Drawer button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  tmr_dp.enabled = True
End Sub
#End If

Private Sub cmd_refresh_Click()
  
  unity_main.errorstring = "Scan Log screen Refresh Values button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Call frm_dbug.loaddbug
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me

#If ABBFT Then
    cmd_testDrawer.Visible = False
    lbl_drawerPos.Visible = False
    txt_drawerPos.Visible = False
#Else
  ' Check if drawer system
  If (MS11CfgData.devID = DTID_DRAWER0) Or (MS11CfgData.devID = DTID_DRAWER1) Then
    cmd_testDrawer.Visible = True
    lbl_drawerPos.Visible = True
    txt_drawerPos.Visible = True
    txt_drawerPos.Text = ""
  Else
    cmd_testDrawer.Visible = False
    lbl_drawerPos.Visible = False
    txt_drawerPos.Visible = False
  End If
#End If
  
  frm_dbug.ss_dbug.ColWidth(1) = 22
  frm_dbug.ss_dbug.ColWidth(2) = 22
End Sub

Sub loaddbug()
  Dim tempstring As String
  Dim trayfailed As Boolean
  
  frm_dbug.ss_dbug.Row = 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg3", "Report Filename")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.m_savePredFile

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg4", "Sample ID")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.txtsamplename.Text

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg5", "Outlier Format")
  frm_dbug.ss_dbug.Col = 2
  
  If (unity_main.m_olFormat = True) Then
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg29", "pics")
  Else
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg30", "values")
  End If

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg6", "Value Limiting (0=act value, 1=bound min, 2=bound max, 3=bound both)")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.m_valueBound

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg7", "M-Distance Alarm")
  frm_dbug.ss_dbug.Col = 2

  If (unity_main.m_alarmMD = 0) Then
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg2", "No")
  Else
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg1", "Yes")
  End If

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg8", "Residual Alarm")
  frm_dbug.ss_dbug.Col = 2
   
  If (unity_main.m_alarmRR = 0) Then
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg2", "No")
  Else
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg1", "Yes")
  End If

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg9", "Property Alarm")
  frm_dbug.ss_dbug.Col = 2
  
  If (unity_main.m_alarmProp = 0) Then
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg2", "No")
  Else
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg1", "Yes")
  End If

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg31", "ND Alarm")
  frm_dbug.ss_dbug.Col = 2

  If (unity_main.m_alarmND = 0) Then
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg2", "No")
  Else
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg1", "Yes")
  End If

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg23", "Save Scan Spectrums")
  frm_dbug.ss_dbug.Col = 2
  
  If (LCase(unity_main.m_saveIt) = "nosave") Then
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg2", "No")
  Else
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg1", "Yes")
  End If

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg10", "Spectrum Save Directory")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.m_saveDir

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg11", "Reference Freq (0=on demand, 1=every sample)")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.m_backFreq

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg22", "Reference Date/Time")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.backdate

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg12", "Sample Name Mode")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.m_nameScanType

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg13", "No. of Model Variables")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.m_numModelVars

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg14", "No. of Properties")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = frmedmod.numprops.Text

#If SSTAR Then
  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg15", "Rotate Dir (0=none, -1=counter clockwise, 1=clockwise)")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.m_rotateDir
#End If

#If SSTAR Then
  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg16", "Rotate Speed")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.m_rotateSpeed
#End If

#If SSTAR Then
  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg17", "Rotate Stepped Steps")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.m_rotateStepSteps
#End If

#If SSTAR Then
  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg18", "Rotate Indexed Steps")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.m_rotateIndexSteps
#End If

#If SSTAR Then
  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg19", "Subtract Dark")
  frm_dbug.ss_dbug.Col = 2
  
  If (unity_main.m_darkSub = False) Then
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg2", "No")
  Else
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg1", "Yes")
  End If
#End If

#If SSTAR Then
  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg20", "Min/Max Wavelengths (nm)")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.m_minWvln & " to " & unity_main.m_maxWvln
#End If

#If ABBFT Then
  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg32", "Starting/Ending Wavenumbers (1/cm)")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.m_mb3000.m_startWavenum & " to " & unity_main.m_mb3000.m_endWavenum
#Else
  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg21", "Starting/Ending Wavelengths (nm)")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.m_smplStartWvln & " to " & unity_main.m_smplEndWvln
#End If

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg24", "User Input Values?")
  frm_dbug.ss_dbug.Col = 2
  
  If (unity_main.m_useMIV = False) Then
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg2", "No")
  Else
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg1", "Yes")
  End If

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg25", "Have Scan Name?")
  frm_dbug.ss_dbug.Col = 2

  If (unity_main.gotscanname = False) Then
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg2", "No")
  Else
    frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg1", "Yes")
  End If

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg26", "No. of Repacks")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.m_repsAvg

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg27", "Repack Counter")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.repcounter

  frm_dbug.ss_dbug.Row = frm_dbug.ss_dbug.Row + 1
  frm_dbug.ss_dbug.Col = 1
  frm_dbug.ss_dbug.Text = MLSupport.GSS("frm_dbug", "debugMsg28", "Print Ticket (0=no, 1=always, 2=on demand)")
  frm_dbug.ss_dbug.Col = 2
  frm_dbug.ss_dbug.Text = unity_main.m_writeTkt
  
  frm_dbug.ss_dbug.MaxRows = frm_dbug.ss_dbug.Row
End Sub

Private Sub lst_status_DataLoaded()

  If (lst_status.ListCount > 30) Then
    lst_status.ListIndex = lst_status.ListCount - 30
  End If
End Sub

#If SSTAR Then
Private Sub tmr_dp_Timer()
  
  ' Get and determine current tray status
#If SSRCS Then
  SSRCSClientError = unity_main.SSRCSClient.GetTrayStatus(unity_main.m_trayStatus)
  
  Select Case (unity_main.m_trayStatus And &H300)
#Else
  Select Case (unity_main.MS11srv.trayStatus And &H300)
#End If
    Case DRWR_POS_OPENED
      txt_drawerPos.Text = MLSupport.GSS("frm_dbug", "txt_drawerPos1", "Opened")
    Case DRWR_POS_CLOSED
      txt_drawerPos.Text = MLSupport.GSS("frm_dbug", "txt_drawerPos2", "Closed")
    Case Else
      txt_drawerPos.Text = MLSupport.GSS("frm_dbug", "txt_drawerPos3", "Partial")
  End Select
End Sub
#End If








