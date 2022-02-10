VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#4.1#0"; "resize32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.0#0"; "HexUniControls30.ocx"
Begin VB.Form frm_help 
   Caption         =   "InfoStar Help"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8070
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
   Icon            =   "frm_help.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin HexUniControls.ctlUniButtonImageXP cmd_exit 
      Cancel          =   -1  'True
      Height          =   650
      Left            =   2888
      TabIndex        =   0
      Top             =   8280
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
      Caption         =   "frm_help.frx":0442
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_help.frx":046A
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_help.frx":048A
   End
   Begin HexUniControls.ctlUniButtonImageXP cmd_manual 
      Height          =   650
      Left            =   2288
      TabIndex        =   1
      Top             =   7440
      Width           =   3495
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
      Caption         =   "frm_help.frx":04A6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frm_help.frx":04EE
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frm_help.frx":050E
   End
   Begin HexUniControls.ctlUniListBoxXP lst_unity 
      Height          =   1425
      Left            =   2400
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   240
      Width           =   4575
      _ExtentX        =   0
      _ExtentY        =   0
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
      Tip             =   "frm_help.frx":052A
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   -1  'True
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_help.frx":054A
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   2520
      _Version        =   262145
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   9285
      FormDesignWidth =   8070
   End
   Begin VB.PictureBox logo_pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   1440
      ScaleHeight     =   2295
      ScaleWidth      =   4935
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2040
      Width           =   4935
   End
   Begin HexUniControls.ctlUniListBoxXP lst_repinfo 
      Height          =   2670
      Left            =   600
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4560
      Width           =   6735
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   16711680
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
      TrapTab         =   0   'False
      Tip             =   "frm_help.frx":0566
      MultiSelect     =   0
      Sorted          =   0   'False
      HScroll         =   -1  'True
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      RoundedBorders  =   0   'False
      SelectorStyle   =   -1
      MousePointer    =   0
      MouseIcon       =   "frm_help.frx":0586
      ManualStart     =   0   'False
      Columns         =   0
   End
   Begin HexUniControls.ctlUniImage ctlUniImage1 
      Height          =   1545
      Left            =   120
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2725
      Picture         =   "frm_help.frx":05A2
      Tip             =   "frm_help.frx":54A9
      Enabled         =   -1  'True
      Border          =   0   'False
      BackColor       =   -2147483633
      BorderColor     =   -1
      RoundedBorders  =   -1  'True
      Stretch         =   -1  'True
      XTransp         =   0
      YTransp         =   0
      MousePointer    =   0
      MouseIcon       =   "frm_help.frx":54C9
   End
   Begin HexUniControls.ctlUniTooltipManager ctlUniTooltipManager1 
      Left            =   360
      Top             =   3600
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
      Left            =   360
      Top             =   3120
      _ExtentX        =   794
      _ExtentY        =   794
      Caption         =   "frm_help.frx":54E5
   End
End
Attribute VB_Name = "frm_help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_manual_Click()
  Dim viewer1, fileName As String
  Dim RetVal As Variant
  Dim i As Integer

  unity_main.errorstring = "InfoStar Help screen View InfoStar Manual button selected"
  unity_main.write_error (LOG_DBG_LEVEL3)
  
  On Error GoTo badload
  
  viewer1 = FileManagement.GetAssociatedApp(".pdf")
  
  i = InStr(viewer1, ".exe")
  If 1 >= 0 Then
    viewer1 = Mid(viewer1, 2, i + 2)
  End If
  
  fileName = (HELP_DIR & INFOSTAR_HELP_FILE & INFOSTAR_VER & "*" & HELP_FILE_EXT)
  RetVal = Shell(viewer1 + " " + fileName, 1)
  Exit Sub
  
badload:
  CWrap.ShowMessageBoxW MLSupport.GGS_Params("frm_help.errMsg1", "Problem loading help manual, either Adobe Acrobat Reader isn't installed or the help file %1 is not present or is corrupt!", fileName), vbCritical
End Sub

Private Sub cmd_exit_Click()
  
  Call unity_main.restart_loop(LOG_DBG_LEVEL3, "InfoStar Help screen Exit button selected")
  Unload frm_help
End Sub

Private Sub Form_Load()
  Dim fileName As String
  Dim uniFile As New clsUniFile
  Dim fEncoding As eFileEncoding
  Dim rc As Boolean
  Dim inString As String
  
  ' Apply language file to form
  MLSupport.ApplyToForm Me
  
  lst_unity.Clear
  lst_unity.AddItem ("Unity Scientific (Corporate Headquarters)")
  lst_unity.AddItem ("117 Old State Road")
  lst_unity.AddItem ("Brookfield, CT 06804")
  lst_unity.AddItem ("Phone: 203-740-2999")
  lst_unity.AddItem ("Fax: 203-740-2955")
  lst_unity.AddItem ("e-mail: info@unityscientific.com")
  lst_unity.AddItem ("Web: www.unityscientific.com")

  lst_repinfo.Clear
  fileName = (CFG_DIR & REP_INFO_FILE)

  If (uniFile.OpenFileRead(fileName) = True) Then
    On Error GoTo BAD_FILE
    fEncoding = uniFile.ReadBOM

    While Not (uniFile.EOF())
      If (fEncoding = fe_ANSI) Then
        rc = uniFile.ReadAnsiLine(inString)
      Else
        rc = uniFile.ReadUnicodeLine(inString)
      End If
      
      If (rc = False) Then GoTo BAD_FILE

      frm_help.lst_repinfo.AddItem Trim(inString)
    Wend
  End If
  
BAD_FILE:
  uniFile.CloseFile
  
  If (CFile.st_FileExist(GRAPHICS_DIR & "company_logo.jpg") = True) Then
    logo_pic.Picture = LoadPicture(GRAPHICS_DIR & "company_logo.jpg")
  End If

  If (frm_help.lst_repinfo.ListCount = 0) Then
    frm_help.lst_repinfo.Visible = False
  End If
End Sub








