VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "ReSize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_modtype 
   Caption         =   "Property Model Type"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7020
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
   Icon            =   "frm_modtype.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   240
      Top             =   3120
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4020
      FormDesignWidth =   7020
   End
   Begin HexUniControls.ctlUniListBoxXP lst_modtype 
      Height          =   2295
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4048
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TrapTab         =   0   'False
      Tip             =   "frm_modtype.frx":08CA
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   -1  'True
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_modtype.frx":08EA
      ManualStart     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_select 
      Height          =   650
      Left            =   960
      TabIndex        =   0
      Top             =   3000
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
      Caption         =   "frm_modtype.frx":0906
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_modtype.frx":0932
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_modtype.frx":0952
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_cancel 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   4080
      TabIndex        =   1
      Top             =   3000
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
      Caption         =   "frm_modtype.frx":096E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_modtype.frx":099A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_modtype.frx":09BA
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   240
      Top             =   3480
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
      Top             =   2640
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_modtype.frx":09D6
   End
End
Attribute VB_Name = "frm_modtype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub model_selected()
  Dim ff As Integer

  frm_1model.modtypex = (frm_modtype.lst_modtype.ListIndex + 1)
  Unload frm_modtype
    
  frmedmod.m_addProp = True
  frmedmod.grid_models.MaxRows = frmedmod.grid_models.MaxRows + 1
  
  Select Case (frm_1model.modtypex)
    Case 1          'pls model
      frm_1model.uniTitle = MLSupport.GSS("frm_1model", "uniTitle1", "GRAMS PLSIQ Model Property Configuration")
      
    Case 2          'mlr model
      frm_1model.uniTitle = MLSupport.GSS("frm_1model", "uniTitle2", "MLR Model Property Configuration")
      
      For ff = 6 To 9
        frm_1model.txt_modvar(ff).Visible = False
        frm_1model.Label1(ff).Visible = False
      Next ff
      
      frm_1model.Picture7.Visible = False
      frm_1model.Picture8.Visible = False
      frm_1model.Picture9.Visible = False
      frm_1model.Picture10.Visible = False
      
    Case 3          'secondary model
      frm_1model.uniTitle = MLSupport.GSS("frm_1model", "uniTitle3", "Secondary Model Property Configuration")
      
      For ff = 6 To 9
        frm_1model.txt_modvar(ff).Visible = False
        frm_1model.Label1(ff).Visible = False
      Next ff
      
      frm_1model.Picture7.Visible = False
      frm_1model.Picture8.Visible = False
      frm_1model.Picture9.Visible = False
      frm_1model.Picture10.Visible = False
      
    Case 4          ' Senslogic CalStar Model
      frmMain.txt_modvar(3).Text = "0" ' intercept
      frmMain.txt_modvar(4).Text = "1" ' slope
      frmMain.txt_modvar(5).Text = "2" ' sig figs
      frmMain.txt_modvar(6).Text = "3.0" ' m dist warn
      frmMain.txt_modvar(7).Text = "7.0" ' m dist fail
      frmMain.txt_modvar(8).Text = "3" ' resid rat warn
      frmMain.txt_modvar(9).Text = "7" ' rr fail
      frmMain.txt_modvar(10).Text = "1" ' prop low warn
      frmMain.txt_modvar(11).Text = "0.5" ' prop low fail
      frmMain.txt_modvar(12).Text = "99" ' prop hi warn
      frmMain.txt_modvar(13).Text = "100" ' prop hi fail
      frmMain.txtmrow.Text = frmedmod.grid_models.MaxRows
      frmMain.Show 1
  End Select
    
  If (frm_1model.modtypex < 4) Then
    frm_1model.txt_modvar(2).Text = "1" ' property index
    frm_1model.txt_modvar(3).Text = "0" ' intercept
    frm_1model.txt_modvar(4).Text = "1" ' slope
    frm_1model.txt_modvar(5).Text = "2" ' sig figs
    frm_1model.txt_modvar(6).Text = "3.0" ' m dist warn
    frm_1model.txt_modvar(7).Text = "7.0" ' m dist fail
    frm_1model.txt_modvar(8).Text = "3" ' resid rat warn
    frm_1model.txt_modvar(9).Text = "7" ' rr fail
    frm_1model.txt_modvar(10).Text = "1" ' prop low warn
    frm_1model.txt_modvar(11).Text = "0.5" ' prop low fail
    frm_1model.txt_modvar(12).Text = "99" ' prop hi warn
    frm_1model.txt_modvar(13).Text = "100" ' prop hi fail
    frm_1model.txtmrow.Text = frmedmod.grid_models.MaxRows
    frm_1model.Show 1
  End If
End Sub

Private Sub cmd_select_Click()
  
  unity_main.errorstring = "Property Model Type screen Select button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  If (frm_modtype.lst_modtype.ListIndex >= 0) Then
    model_selected
  Else
    CWrap.ShowMessageBoxW MLSupport.GSS("frm_modtype", "errMsg1", "You must select a model type from the list!"), vbOKOnly
  End If
End Sub

Private Sub cmd_cancel_Click()
  
  unity_main.errorstring = "Property Model Type screen Cancel button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  Unload frm_modtype
End Sub

Private Sub Form_Load()
  
  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
  frm_modtype.lst_modtype.Clear
  frm_modtype.lst_modtype.AddItem (MLSupport.GSS("frm_modtype", "lst_modtype1", "GRAMS PLSIQ Model") & " (*" & GRAMS_MODEL_FILE_EXT & ")")
  frm_modtype.lst_modtype.AddItem (MLSupport.GSS("frm_modtype", "lst_modtype2", "MLR Model") & " (*" & MLR_MODEL_FILE_EXT & ")")
  frm_modtype.lst_modtype.AddItem (MLSupport.GSS("frm_modtype", "lst_modtype3", "Secondary Model") & " (*" & SEC_MODEL_FILE_EXT & ")")

  If unity_main.calstar_enabled = True Then
    frm_modtype.lst_modtype.AddItem (MLSupport.GSS("frm_modtype", "lst_modtype4", "CalStar Model") & " (*" & CALSTAR_MODEL_FILE_EXT & ")")
  End If
  
End Sub

Private Sub lst_modtype_DblClick()
  
  model_selected
End Sub








