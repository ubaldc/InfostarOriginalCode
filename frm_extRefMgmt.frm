VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_extRefMgmt 
   Caption         =   "External References Management"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6960
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
   ScaleHeight     =   5655
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniListBoxXP lst_refFileNames 
      Height          =   3090
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5450
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
      Tip             =   "frm_extRefMgmt.frx":0000
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   -1  'True
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_extRefMgmt.frx":0020
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_create 
      Height          =   645
      Left            =   4320
      TabIndex        =   1
      Top             =   3600
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
      Caption         =   "frm_extRefMgmt.frx":003C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_extRefMgmt.frx":0084
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_extRefMgmt.frx":00A4
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_edit 
      Height          =   645
      Left            =   720
      TabIndex        =   2
      Top             =   3600
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
      Caption         =   "frm_extRefMgmt.frx":00C0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_extRefMgmt.frx":00FC
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_extRefMgmt.frx":011C
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_delete 
      Height          =   645
      Left            =   720
      TabIndex        =   3
      Top             =   4440
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
      Caption         =   "frm_extRefMgmt.frx":0138
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_extRefMgmt.frx":0178
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_extRefMgmt.frx":0198
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_exit 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   645
      Left            =   4320
      TabIndex        =   4
      Top             =   4440
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
      Caption         =   "frm_extRefMgmt.frx":01B4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_extRefMgmt.frx":01DC
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_extRefMgmt.frx":01FC
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3600
      Top             =   5160
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5655
      FormDesignWidth =   6960
   End
   Begin HexUniControls.ctlUniFormCaption uniTitle 
      Left            =   2160
      Top             =   5160
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_extRefMgmt.frx":0218
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   2880
      Top             =   5160
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
Attribute VB_Name = "frm_extRefMgmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_create_Click()

  unity_main.errorstring = "External References Management screen Create New Reference button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  frm_extRef.init_ext_ref_settings
  frm_extRef.Show 1
End Sub

Private Sub cmd_delete_Click()
  Dim optVal As Integer
  Dim fileName As String
  Dim uniFile As New clsUniFile

  unity_main.errorstring = "External References Management screen Delete Reference button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)

  If (lst_refFileNames.ListIndex < 0) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_extRefMgmt", "errMsg1", "You must select a reference file from the list to delete"), vbOKOnly
    Exit Sub
  End If
  
  optVal = CWrap.ShowMessageBoxW(MLSupport.GSS("frm_extRefMgmt", "statMsg1", "Are you sure you want to delete this external reference file?"), vbYesNo)

  If (optVal = vbYes) Then
    fileName = (EXT_REFS_CFG_DIR & lst_refFileNames.List(lst_refFileNames.ListIndex) & CFG_FILE_EXT)
    uniFile.st_RmFile fileName
    frm_collect.build_ref_name_list "external"
    unity_main.errorstring = "User deleted external reference file " & fileName
    unity_main.write_error (LOG_DBG_LEVEL1)
  End If
End Sub

Private Sub cmd_edit_Click()

  unity_main.errorstring = "External References Management screen Edit Reference button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  If (lst_refFileNames.ListIndex < 0) Then
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_extRefMgmt", "errMsg2", "You must select a reference file from the list to edit"), vbOKOnly
    Exit Sub
  Else
    frm_extRef.load_ext_ref_cfg_file lst_refFileNames.List(lst_refFileNames.ListIndex), False
    frm_extRef.Show 1
  End If
  
End Sub

Private Sub cmd_exit_Click()
  
  unity_main.errorstring = "External References Management screen Exit button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  Call unity_main.load_prod_file("", False)
  frm_extRefMgmt.Visible = False
End Sub

Private Sub Form_Load()

  ' Apply language file to form
  MLSupport.ApplyToForm Me
End Sub






